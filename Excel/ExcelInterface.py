"""Main interface class for Excel->Python conversion

ExcelInterface provides the single resource for converting the data
in the Excel spreadsheet into a list of StreamNode classes for use
by the HeatSource model.
"""
# Builtin methods
from __future__ import division
from itertools import ifilter, izip, chain, repeat, count
from math import ceil, log, degrees, atan
from datetime import datetime, timedelta
from win32com.client import Dispatch
from os.path import exists, join, split, normpath
from os import unlink
from sys import exit
from win32gui import PumpWaitingMessages
from bisect import bisect
from time import strptime, ctime, gmtime
from calendar import timegm

# Heat Source Methods
from ..Dieties.IniParamsDiety import IniParams
from ..Stream.StreamNode import StreamNode
from ..Dieties.ChronosDiety import Chronos
from ExcelDocument import ExcelDocument
from ..Utils.Dictionaries import Interpolator
from ..Utils.easygui import buttonbox

from .. import opt
try:
    if opt(__name__):
        import psyco.classes
        object = psyco.classes.psyobj
except ImportError: pass

class ExcelInterface(ExcelDocument):
    """Defines an interface specific to the Current (version 8.x) HeatSource Excel interface.

    This class provides methods which seek knowingly through a correctly formatted Excel
    spreadsheet. It creates a list of StreamNode instances, and populates those in"""
    def __init__(self, filename=None, log=None, run_type=0):
        ExcelDocument.__init__(self, filename)
        self.run_type = run_type
        self.log = log
        self.Reach = {}
        #######################################################
        # Grab the initialization parameters from the Excel file.
        lst = {"name": "C4",
               "length": "C5",
               "outputdir": "C6",
               "date": "C8",
               "modelstart": "C9",
               "modelend": "C10",
               "end": "C11",
               "flushdays": "C12",
               "offset": "C13",
               "dt": "E4",
               "dx": "E5",
               "longsample": "E6",
               "transsample": "E7",
               "inflowsites": "E8",
               "contsites": "E9",
               "calcevap": "E11",
               "evapmethod": "E12",
               "wind_a": "E13",
               "wind_b": "E14",
               "calcalluvium": "E15",
               "alluviumtemp": "E16",
               "emergent": "E17",
               "lidar": "E18",
               "lcdensity": "E19",
               "lcoverhang": "E20"}
        for k,v in lst.iteritems():
            IniParams[k] = self.GetValue(v, "Heat Source Inputs")
        # These might be blank, make them zeros
        for key in ["inflowsites","flushdays","wind_a","wind_b"]:
            IniParams[key] = 0.0 if not IniParams[key] else IniParams[key]
        # Then make all of these integers because they're used later in for loops
        for key in ["inflowsites","flushdays","contsites"]:
            IniParams[key] = int(IniParams[key])
        # Set up our evaporation method
        IniParams["penman"] = False
        if IniParams["calcevap"]:
            IniParams["penman"] = True if IniParams["evapmethod"] == "Penman" else False
        # The offset should be negated to work around issues with internal date
        # representation. i.e. Pacific time is -7 from UTC, but the code needs a +7 to work.
        # TODO: This is probably a bug in ChronosDiety, not the time module.
        IniParams["offset"] = -1 * IniParams["offset"]
        # Make the dates into datetime instances of the start/stop dates
        IniParams["date"] = timegm(strptime(IniParams["date"].Format("%m/%d/%y %H:%M:%S"),"%m/%d/%y %H:%M:%S"))
        IniParams["end"] = timegm(strptime(IniParams["end"].Format("%m/%d/%y") + " 23:59:59","%m/%d/%y %H:%M:%S"))
        if IniParams["modelstart"] is None:
            IniParams["modelstart"] = IniParams["date"]
        else:
            IniParams["modelstart"] = timegm(strptime(IniParams["modelstart"].Format("%m/%d/%y %H:%M:%S"),"%m/%d/%y %H:%M:%S"))
        if IniParams["modelend"] is None:
            IniParams["modelend"] = IniParams["end"]
        else:
            IniParams["modelend"] = timegm(strptime(IniParams["modelend"].Format("%m/%d/%y") + " 23:59:59","%m/%d/%y %H:%M:%S"))
        IniParams["flushtimestart"] = IniParams["modelstart"] - IniParams["flushdays"]*86400
        # make sure alluvium temp is present and a floating point number.
        IniParams["alluviumtemp"] = 0.0 if not IniParams["alluviumtemp"] else float(IniParams["alluviumtemp"])
        # make sure that the timestep divides into 60 minutes, or we may not land squarely on each hour's starting point.
        if 60%IniParams["dt"] > 1e-7:
            raise Exception("I'm sorry, your timestep (%0.2f) must evenly divide into 60 minutes." % IniParams["dt"])
        else:
            IniParams["dt"] = IniParams["dt"]*60 # make dt measured in seconds
        # Make sure the output directory ends in a slash (VB chokes if not)
        if IniParams["outputdir"][-1] != "\\":
            raise Exception("Output directory needs to have a trailing backslash")
        # Set up the log file in the outputdir
        self.log.SetFile(normpath(join(IniParams["outputdir"],"outfile.log")))

        # Make empty Dictionaries for the boundary conditions
        self.Q_bc = Interpolator()
        self.T_bc = Interpolator()

        # List of kilometers with continuous data nodes assigned.
        self.ContDataSites = []

        # the distance step must be an exact, greater or equal to one, multiple of the sample rate.
        if (IniParams["dx"]%IniParams["longsample"]
            or IniParams["dx"]<IniParams["longsample"]):
            raise Exception("Distance step must be a multiple of the Longitudinal transfer rate")
        # Some convenience variables
        self.dx = IniParams["dx"]
        self.multiple = int(self.dx/IniParams["longsample"]) #We have this many samples per distance step

        # Get the list of times in the flow and continuous data sheets- we make no assumptions
        # that they equal each other.
        self.flowtimelist = self.GetTimelist("Flow Data")
        self.continuoustimelist = self.GetTimelist("Continuous Data")
        self.flushtimelist = self.GetFlushTimelist()

        #####################
        # Now we start through the steps of building a reach full of StreamNodes
        self.GetBoundaryConditions()
        self.BuildNodes()
        if IniParams["lidar"]: self.BuildZonesLidar()
        else: self.BuildZonesNormal()
        self.GetTributaryData()
        self.GetContinuousData()
        self.SetAtmosphericData()
        self.OrientNodes()

    def OrientNodes(self):
        self.PB("Initializing StreamNodes")
        # Now we manually set each nodes next and previous kilometer values by stepping through the reach
        l = sorted(self.Reach.keys(), reverse=True)
        head = self.Reach[max(l)] # The headwater node
        # Set the previous and next kilometer of each node.
        slope_problems = []
        for i in xrange(len(l)):
            key = l[i] # The current node's key
            # Then, set pointers to the next and previous nodes
            if i == 0: pass
            else: self.Reach[key].prev_km = self.Reach[l[i-1]] # At first node, there's no previous
            try:
                self.Reach[key].next_km = self.Reach[l[i+1]]
            except IndexError:
            ##!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            ## For last node (mouth) we set the downstream node equal to self, this is because
            ## we want to access the node's temp if there's no downstream, and this safes us an
            ## if statement.
                self.Reach[key].next_km = self.Reach[key]
            # Set a headwater node
            self.Reach[key].head = head
            self.Reach[key].Initialize()
            # check for a zero slope. We store all of them before checking so we can print a lengthy error that no-one will ever read.
            if self.Reach[key].S <= 0.0: slope_problems.append(key)
        if self.run_type != 1: # zeros are alright in shade calculations
            if len(slope_problems):
                raise Exception ("The following reaches have zero slope. Kilometers: %s" %",".join(['%0.3f'%i for i in slope_problems]))

    def close(self):
        del self.T_bc, self.Reach

    def CheckEarlyQuit(self):
        """Checks a value to see whether the user wants to stop the model before we completely set everything up"""
        if exists("c:\\quit_heatsource"):
            unlink("c:\\quit_heatsource")
            self.QuitMessage()

    def SetAtmosphericData(self):
        """For each node without continuous data, use closest (up or downstream) node's data"""
        self.CheckEarlyQuit()
        self.PB("Setting Atmospheric Data")
        sites = self.ContDataSites # Localize the variable for speed
        sites.sort() #Sort is necessary for the bisect module
        c = count()
        l = self.Reach.keys()
        # This routine bisects the reach and searches the difference between us and the upp
        for km, node in self.Reach.iteritems():
            if km not in sites:
                # Kilometer's downstream and upstream
                lower = bisect(sites,km)-1 if bisect(sites,km)-1 > 0 else 0 # zero is the lowest (protect against value of -1)
                # bisect returns the length of a list when the bisecting number is greater than the greatest value.
                # Here we protect by max-ing out at the length of the list.
                upper = min([bisect(sites,km),len(sites)-1])
                # Use the indexes to get the kilometers from the sites list
                down = sites[lower]
                up = sites[upper]
                datasite = self.Reach[up] # Initialize to upstream's continuous data
                if km-down < up-km: # Only if the distance to the downstream node is closer do we use that
                    datasite = self.Reach[down]
                self.Reach[km].ContData = datasite.ContData
                self.PB("Setting Atmospheric Data", c.next(), len(l))

    def GetBoundaryConditions(self):
        """Get the boundary conditions from the "Continuous Data" page"""
        self.CheckEarlyQuit()
        # Get the columns, which is faster than accessing cells
        self.PB("Reading boundary conditions")
        sheetname = "Continuous Data"
        timelist = self.continuoustimelist
        Rstart, Cstart = 5,5
        Rend = Rstart + len(timelist) - 1
        Cend = 7
        rng = ((Rstart, Cstart),(Rend, Cend))
        # the data block is a tuple of tuples, each corresponding to a timestamp.
        data = self.GetValue(rng, sheetname)
        # Check out GetTributaryData() for details on this reformatting of the data
        # for the progress bar
        length = len(data)
        c = count()
        # Now set the discharge and temperature boundary condition dictionaries.

        for i in xrange(len(timelist)):
            time = timelist[i]
            t, flow, temp = data[i]
            # Get the flow boundary condition
            if flow == 0 or not flow:
                if self.run_type != 1:
                    raise Exception("Missing flow boundary condition for day %s " % ctime(time))
                else: flow = 0
            self.Q_bc[time] = flow
            # Temperature boundary condition
            t_val = temp if temp is not None else 0.0
            self.T_bc[time] = t_val
            self.PB("Reading boundary conditions",c.next(),length)

        # Next we expand or revise the dictionary to account for the flush period
        # Flush flow: model start value over entire flush period
        for i in xrange(len(self.flushtimelist)):
            time = self.flushtimelist[i]
            self.Q_bc[time] = self.Q_bc[IniParams["modelstart"]]
        # Flush temperature: first 24 hours repeated over flush period
        first_day_time = IniParams["modelstart"]
        second_day = IniParams["modelstart"] + 86400
        for i in xrange(len(self.flushtimelist)):
            time = self.flushtimelist[i]
            self.T_bc[time] = self.T_bc[first_day_time]
            first_day_time += 3600
            if first_day_time >= second_day:
                first_day_time = IniParams["modelstart"]


        self.Q_bc = self.Q_bc.View(IniParams["flushtimestart"], IniParams["modelend"], aft=1)
        self.T_bc = self.T_bc.View(IniParams["flushtimestart"], IniParams["modelend"], aft=1)

    def GetTimelist(self, sheet):
        """Return list of floating point time values corresponding to the data available in the sheet"""
                # Sheetname: (column, starting row)
        nums = {'Continuous Data': (5, 4),
                'Flow Data': (11, 3)}
        col, row = nums[sheet]
        timelist = [i for i in ifilter(None,self.GetColumn(col,sheet)[row:])]

        timelist2 = []
        for t in timelist:
            # strptime returns a tuple, we only want the first 8 elements as a list, then we want to
            # add a zero to the end of the list. We append this list to timelist2 and ship it off
            # as a tuple
            tm = [i for i in strptime(t.Format("%m/%d/%y %H:%M:%S"),"%m/%d/%y %H:%M:%S")[0:8]] + [0]
            timelist2.append(timegm(tm))
        return tuple(timelist2)

    def GetFlushTimelist(self):
        #Build a timelist that represents the flushing period
        #This assumes that data is hourly, not tested with variable input timesteps
        flushtimelist = []
        flushtime = IniParams["flushtimestart"]
        while flushtime < IniParams["modelstart"]:
            flushtimelist += flushtime,
            flushtime += 3600
        return tuple(flushtimelist)

    def GetLocations(self,sheetname):
        """Return a list of kilometers corresponding to the inflow or continuous data sites"""
        #                        Number of sites, row, column
        d = {'Continuous Data': (IniParams["contsites"], 5, 3),
             'Flow Data': (IniParams["inflowsites"], 4, 9)}
        t = ()
        l = self.Reach.keys()
        l.sort()
        ini, row, col = d[sheetname]
        for site in xrange(ini):
            km = self.GetValue((site + row, col),sheetname)
            if km is None or not isinstance(km, float):
                # This is a bad dataset if there's no kilometer
                raise Exception("Must have a stream kilometer (e.g. 15.3) for each node in %s page!" % sheetname)
            key = bisect(l,km)-1
            t += l[key], # Index by kilometer
        return t

    def GetTributaryData(self):
        """Populate the tributary flow and temperature values for nodes from the Flow Data page"""
        self.CheckEarlyQuit()
        self.PB("Reading inflow data")
        sheetname = "Flow Data"
        timelist = self.flowtimelist
        # Get a list of the timestamps that we have data for, and use that to grab the data block
        Rstart, Cstart = 4,12
        Rend = Rstart + len(timelist) - 1
        Cend = IniParams["inflowsites"]*2 + Cstart - 1
        rng = ((Rstart, Cstart),(Rend, Cend))
        # the data block is a tuple of tuples, each corresponding to a timestamp.
        data = self.GetValue(rng, sheetname)
        # Current data is in the form:
        # | Site 1   | Site 2   | Site 3   | ...
        # ((0.3, 15.7, 0.3, 17.7, 0.02, 18.2), (, ...)
        # Where every tuple is a data record corresponding to a time, and every
        # two numbers in the tuple refer to a site's flow rate and temp. We want
        # to change this to the form:
        # | Site 1     | Site 2     | Site 3       | ...
        # [((0.3, 15.7), (0.3, 17.7), (0.02, 18.2)), (, ...))]
        # To facilitate each site having it's own two item tuple.
        # The calls to tuple() just ensure that we are not making lists, which can
        # be changed accidentally. Without them, the line is easier to understand:
        # [zip(line[0:None:2],line[1:None:2]) for line in data]
        data = [tuple(zip(line[0:None:2],line[1:None:2])) for line in data]

        # Get a tuple of kilometers to use as keys to the location of each tributary
        kms = self.GetLocations("Flow Data")
        length = len(timelist)
        tm = count() # Which datapoint time are we recording
        nodelist = [] # Quick list of nodes with flow data
        for time in timelist:
            line = data.pop(0)
            # Error checking?! Naw!!
            c = count()
            for flow, temp in line:
                i = c.next()
                node = self.Reach[kms[i]] # Index by kilometer
                if node not in nodelist or not len(nodelist): nodelist.append(node)
                if flow is None or (flow > 0 and temp is None):
                    raise Exception("Cannot have a tributary with blank flow or temperature conditions")
                # Here, we actually set the tribs library, appending to a tuple. Q_ and T_tribs are
                # tuples of values because we may have more than one input for a given node
                node.Q_tribs[time] += flow, #Append to tuple
                #print node, time, flow, node.Q_tribs[time]
                node.T_tribs[time] += temp,
            self.PB("Reading inflow data",tm.next(), length)

        # Next we expand or revise the dictionary to account for the flush period
        # Flush flow: model start value over entire flush period
        for i in xrange(len(self.flushtimelist)):
            time = self.flushtimelist[i]
            for node in nodelist:
                node.Q_tribs[time] = node.Q_tribs[IniParams["modelstart"]]
        # Flush temperature: first 24 hours repeated over flush period
        first_day_time = IniParams["modelstart"]
        second_day = IniParams["modelstart"] + 86400
        for i in xrange(len(self.flushtimelist)):
            time = self.flushtimelist[i]
            for node in nodelist:
                node.T_tribs[time] = node.T_tribs[first_day_time]
            first_day_time += 3600
            if first_day_time >= second_day:
                first_day_time = IniParams["modelstart"]

        # Now we strip out the unnecessary values from the dictionaries. This is placed here
        # at the end so we can dispose of it easily if necessary
        for node in nodelist:
            node.Q_tribs = node.Q_tribs.View(IniParams["flushtimestart"], IniParams["modelend"], aft=1)
            node.T_tribs = node.T_tribs.View(IniParams["flushtimestart"], IniParams["modelend"], aft=1)

    def GetContinuousData(self):
        """Get data from the "Continuous Data" page"""
        # This is remarkably similar to GetInflowData. We get a block of data, then set the dictionary of the node
        self.CheckEarlyQuit()
        self.PB("Reading Continuous Data")
        sheetname = "Continuous Data"
        Rstart,Cstart = 5,9
        timelist = self.continuoustimelist
        Rend = Rstart + len(timelist) - 1
        #We need five columns because stream temp data (which we ignore in heat source)
        Cend = IniParams["contsites"]*5 + Cstart-1
        rng = ((Rstart,Cstart),(Rend,Cend))
        data = self.GetValue(rng,"Continuous Data")
        # See GetTributaryData() for info on this crazy one-liner
        data = [tuple(zip(line[0:None:5],line[1:None:5],line[2:None:5],line[3:None:5])) for line in data]
        kms = self.GetLocations("Continuous Data")
        tm = count() # Which datapoint time are we recording
        length = len(timelist)
        for time in timelist:
            line = data.pop(0)
            c = count()
            for cloud, wind, humid, air in line:
                i = c.next()
                node = self.Reach[kms[i]] # Index by kilometer
                # Append this node to a list of all nodes which have continuous data
                if node.km not in self.ContDataSites:
                    self.ContDataSites.append(node.km)
                # Perform some tests for data accuracy and validity
                if cloud is None: cloud = 0.0
                if wind is None: wind = 0.0
                if cloud < 0 or cloud > 1:
                    if self.run_type == 1: # Alright in shade-a-lator
                        cloud = 0.0
                    else: raise Exception("Cloudiness (value of '%s' in Continuous Data) must be greater than zero and less than one." % `cloud`)
                if humid < 0 or humid is None or humid > 1:
                    if self.run_type == 1: # Alright in shade-a-lator
                        humid = 0.0
                    else: raise Exception("Humidity (value of '%s' in Continuous Data) must be greater than zero and less than one." % `hum_val`)
                if air is None or air < -90 or air > 58:
                    if self.run_type == 1: # Alright in shade-a-lator
                        air = 0.0
                    else: raise Exception("Air temperature input (value of '%s' in Continuous Data) outside of world records, -89 to 58 deg C." % `air`)
                node.ContData[time] = cloud, wind, humid, air
            self.PB("Reading continuous data", tm.next(), length)

        # Flush meteorology: first 24 hours repeated over flush period
        first_day_time = IniParams["modelstart"]
        second_day = IniParams["modelstart"] + 86400
        for i in xrange(len(self.flushtimelist)):
            time = self.flushtimelist[i]
            for km in self.ContDataSites:
                node = self.Reach[km]
                node.ContData[time] = node.ContData[first_day_time]
            first_day_time += 3600
            if first_day_time >= second_day:
                first_day_time = IniParams["modelstart"]

        # Now we strip out the unnecessary values from the dictionaries. This is placed here
        # at the end so we can dispose of it easily if necessary
        self.PB("Subsetting the Continuous Data")
        tm = count()
        length = len(self.ContDataSites)
        for km in self.ContDataSites:
            node = self.Reach[km]
            node.ContData = node.ContData.View(IniParams["flushtimestart"], IniParams["modelend"], aft=1)
            self.PB("Subsetting the Continuous Data",tm.next(), length)
    def zipper(self,iterable,mul=2):
        """Zippify list by grouping <mul> consecutive elements together

        Zipper returns a list of lists where the internal lists are groups of <mul> consecutive
        elements from the input list. For example:
        >>> lst = [0,1,2,3,4,5,6,7,8,9]
        >>> zipper(lst)
        [[0],[1,2],[3,4][5,6],[7,8],[9]]
        The first element is a length 1 list because we assume that it is a single element node (headwaters).
        Note that the last element, 9, is alone as well, this method will figure out when there are not
        enough elements to make n equal length lists, and modify itself appropriately so that the remaining list
        will contain all leftover elements. The usefulness of this method is that it will allow us to average over each <mul> consecutive elements
        """
        # From itertools recipes... We use all but the first (boundary node) element
        lst = [i for i in izip(*[chain(iterable[1:], repeat(None, mul-1))]*mul)]
        # Then we tack on the boundary node element
        lst.insert(0,(iterable[0],))
        # Then strip off the None values from the last (if any)
        lst[-1] = tuple(ifilter(lambda x: x is not None,lst[-1]))
        return self.numify(lst)

    def numify(self, lst):
        """Take a list of iterables and remove all values of None or empty strings"""
        # Remove None values at the end of each individual list
        for i in xrange(len(lst)):
            # strip out values of None from the tuple, returning a new tuple
            lst[i] = [x for x in ifilter(lambda x: x is not None, lst[i])]
        # Remove blank strings from within the list
        for l in lst:
            n = []
            for i in xrange(len(l)):
                if l[i] == "": n.append(i)
            n.reverse()
            for i in n: del l[i]
        # Make sure there are no zero length lists because they'll fail if we average
        for i in xrange(len(lst)):
            if len(lst[i]) == 0: lst[i].append(0.0)
        return lst

    def multiplier(self, iterable, predicate=lambda x:x):
        """Return an iterable that was run through the zipper

        Take an iterable and strip the values of None, then send to the zipper
        and apply predicate to each value returned (zipper returns a list)"""
        # This is a way to safely apply a generic lambda function to an iterable.
        # If I were paying attention to design, instead of just hacking away, I would
        # have done this with decorators to modify the function. Now I'm too lazy to
        # re-write it (well, not lazy, but I'm not paid as a programmer, and so I have
        # "better" things to do than optimize our code.)
        # First we strip off the None values.
        stripNone = lambda y: [i for i in ifilter(lambda x: x is not None, y)]
        return [predicate(stripNone(x)) for x in self.zipper(iterable,self.multiple)]

    def zeroOutList(self, lst):
        """Replace blank values in a list with zeros"""
        test = lambda x: 0.0 if x=="" else x
        return [test(i) for i in lst]

    def GetColumnarData(self):
        """return a dictionary of attributes that are averaged or summed as appropriate"""
        self.CheckEarlyQuit()
        # Pages we grab columns from, and the columns that we grab
        ttools = ["km","Longitude","Latitude"]
        morph = ["Elevation","S","W_b","z","n","SedThermCond","SedThermDiff","SedDepth",
                 "hyp_percent","phi","FLIR_Time","FLIR_Temp","Q_cont","d_cont"]
        flow = ["Q_in","T_in","Q_out"]
        # Ways that we grab the columns
        sums = ["hyp_percent","Q_in","Q_out"] # These are summed, not averaged
        mins = ["km"]
        aves = ["Longitude","Latitude","Elevation","S","W_b","z","n","SedThermCond",
                "SedThermDiff","SedDepth","phi", "Q_cont","d_cont","T_in"]
        ignore = ["FLIR_Temp","FLIR_Time"] # Ignore in the loop, set them manually

        data = {}
        # Get all the columnar data from the sheets
        for i in xrange(len(ttools)):
            data[ttools[i]] = self.GetColumn(1+i, "TTools Data")[5:]
        for i in xrange(len(morph)):
            data[morph[i]] = self.GetColumn(2+i, "Morphology Data")[5:]
        for i in xrange(len(flow)):
            data[flow[i]] = self.GetColumn(2+i, "Flow Data")[3:]
        #Longitude check
        if max(data[ttools[1]]) > 180 or min(data[ttools[1]]) < -180:
            raise Exception("Longitude must be greater than -180 and less than 180 degrees")
        #Latitude check
        if max(data[ttools[2]]) > 90 or min(data[ttools[2]]) < -90:
            raise Exception("Latitude must be greater than -90 and less than 90 degrees")

        # Then sum and average things as appropriate. multiplier() takes a tuple
        # and applies the given lambda function to that tuple.
        for attr in sums:
            data[attr] = self.multiplier(data[attr],lambda x:sum(x))
        for attr in aves:
            data[attr] = self.multiplier(data[attr],lambda x:sum(x)/len(x))
        for attr in mins:
            data[attr] = self.multiplier(data[attr],lambda x:min(x))
        return data

    def BuildNodes(self):
        # This is the worst of the methods. At some point, dealing with the collection of data
        # from an excel spreadsheet is going to cause trouble. I tried to keep the trouble to a
        # minimum, but this is one of the bad methods of our interface with Excel.
        self.CheckEarlyQuit()
        self.PB("Building Stream Nodes")
        Q_mb = 0.0
        # Grab all of the data in a dictionary
        data = self.GetColumnarData()
        #################################
        # Build a boundary node
        node = StreamNode(run_type=self.run_type,Q_mb=Q_mb)
        # Then set the attributes for everything in the dictionary
        for k,v in data.iteritems():
            setattr(node,k,v[0])
        # set the flow and temp boundary conditions for the boundary node
        node.Q_bc = self.Q_bc
        node.T_bc = self.T_bc
        self.InitializeNode(node)
        node.dx = IniParams["longsample"]
        self.Reach[node.km] = node
        ############################################

        #Figure out how many nodes we should have downstream. We use math.ceil() because
        # if we end up with a fraction, that means that there's a node at the end that
        # is not a perfect multiple of the sample distance. We might end up ending at
        # stream kilometer 0.5, for instance, in that case
        vars = (IniParams["length"] * 1000)/IniParams["longsample"]

        num_nodes = int(ceil((vars)/self.multiple))
        for i in range(0, num_nodes):
            node = StreamNode(run_type=self.run_type,Q_mb=Q_mb)
            for k,v in data.iteritems():
                setattr(node,k,v[i+1])# Add one to ignore boundary node
            self.InitializeNode(node)
            self.Reach[node.km] = node
            self.PB("Building Stream Nodes", i, vars/self.multiple)
        # Find the mouth node and calculate the actual distance
        mouth = self.Reach[min(self.Reach.keys())]
        mouth_dx = (vars)%self.multiple or 1.0 # number of extra variables if we're not perfectly divisible
        mouth.dx = IniParams["longsample"] * mouth_dx


    def BuildZonesNormal(self):
        """This method builds the sampled vegzones in the case of non-lidar datasets"""
        # Hide your straight razors. This implementation will make you want to use them on your wrists.
        self.CheckEarlyQuit()
        LC = self.GetLandCoverCodes() # Pull the LULC data from the appropriate sheet
        vheight = []
        vdensity = []
        overhang = []
        elevation = []
        average = lambda x:sum(x)/len(x)

        keys = self.Reach.keys()
        keys.sort(reverse=True) # Downstream sorted list of stream kilometers
        self.PB("Translating LULC Data")
        for i in xrange(7, 36): # For each column of LULC data
            col = self.GetColumn(i, "TTools Data")[5:] # LULC column
            elev = self.GetColumn(i+28,"TTools Data")[5:] # Shift by 28 to get elevation column
            # Make a list from the LC codes from the column, then send that to the multiplier
            # with a lambda function that averages them appropriately. Note, we're averaging over
            # the values (e.g. density) not the actual code, which would be meaningless.
            try:
                vheight.append(self.multiplier([LC[x][0] for x in col], average))
                vdensity.append(self.multiplier([LC[x][1] for x in col], average))
                overhang.append(self.multiplier([LC[x][2] for x in col], average))
            except KeyError, (stderr):
                raise Exception("At least one land cover code from the 'TTools Data' worksheet is blank or not in 'Land Cover Codes' worksheet (Code: %s)." % stderr.message)
            if i>7:  #We don't want to read in column AJ -Dan
                elevation.append(self.multiplier(elev, average))
            self.PB("Translating LULC Data", i, 36)
        # We have to set the emergent vegetation, so we strip those off of the iterator
        # before we record the zones.
        for i in xrange(len(keys)):
            node = self.Reach[keys[i]]
            node.VHeight = vheight[0][i]
            node.VDensity = vdensity[0][i]
            node.Overhang = overhang[0][i]

        # Average over the topo values
        topo_w = self.multiplier(self.GetColumn(4, "TTools Data")[5:], average)
        topo_s = self.multiplier(self.GetColumn(5, "TTools Data")[5:], average)
        topo_e = self.multiplier(self.GetColumn(6, "TTools Data")[5:], average)

        # ... and you thought things were crazy earlier! Here is where we build up the
        # values for each node. This is culled from earlier version's VB code and discussions
        # to try to simplify it... yeah, you read that right, simplify it... you should've seen in earlier!
        for h in xrange(len(keys)):
            self.PB("Building VegZones", h, len(keys))
            node = self.Reach[keys[h]]
            VTS_Total = 0 #View to sky value
            LC_Angle_Max = 0
            # Now we set the topographic elevations in each direction
            node.TopoFactor = (topo_w[h] + topo_s[h] + topo_e[h])/(90*3) # Topography factor Above Stream Surface
            # This is basically a list of directions, each with a sort of average topography
            ElevationList = (topo_e[h],
                             topo_e[h],
                             0.5*(topo_e[h]+topo_s[h]),
                             topo_s[h],
                             0.5*(topo_s[h]+topo_w[h]),
                             topo_w[h],
                             topo_w[h])
            # Sun comes down and can be full-on, blocked by veg, or blocked by topography. Earlier implementations
            # calculated each case on the fly. Here we chose a somewhat more elegant solution and calculate necessary
            # angles. Basically, there is a minimum angle for which full sun is calculated (top of trees), and the
            # maximum angle at which full shade is calculated (top of topography). Anything in between these is an
            # angle for which sunlight is passing through trees. So, for each direction, we want to calculate these
            # two angles so that late we can test whether we are between them, and only do the shading calculations
            # if that is true.

            for i in xrange(7): # Iterate through each direction
                T_Full = () # lowest angle necessary for full sun
                T_None = () # Highest angle necessary for full shade
                rip = () # Riparian extinction, basically the amount of loss due to vegetation shading
                for j in xrange(4): # Iterate through each of the 4 zones
                    Vheight = vheight[i*4+j+1][h]
                    Vdens = vdensity[i*4+j+1][h]
                    Overhang = overhang[i*4+j+1][h]
                    Elev = elevation[i*4+j][h]

                    if not j: # We are at the stream edge, so start over
                        LC_Angle_Max = 0 # New value for each direction
                    else:
                        Overhang = 0 # No overhang away from the stream
                    ##########################################################
                    # Calculate the relative ground elevation. This is the
                    # vertical distance from the stream surface to the land surface
                    SH = Elev - node.Elevation
                    # Then calculate the relative vegetation height
                    VH = Vheight + SH
                    # Calculate the riparian extinction value
                    try:
                        RE = -log(1-Vdens)/10
                    except OverflowError:
                        if Vdens == 1: RE = 1 # cannot take log of 0, RE is full if it's zero
                        else: raise
                    # Calculate the node distance
                    LC_Distance = IniParams["transsample"] * (j + 0.5) #This is "+ 0.5" because j starts at 0.
                    # We shift closer to the stream by the amount of overhang
                    # This is a rather ugly cludge.
                    if not j: LC_Distance -= Overhang
                    if LC_Distance <= 0:
                        LC_Distance = 0.00001
                    # Calculate the minimum sun angle needed for full sun
                    T_Full += degrees(atan(VH/LC_Distance)), # It gets added to a tuple of full sun values
                    # Now get the maximum of bank shade and topographic shade for this
                    # direction
                    T_None += degrees(atan(SH/LC_Distance)), # likewise, a tuple of values
                    ##########################################################
                    # Now we calculate the view to sky value
                    # LC_Angle is the vertical angle from the surface to the land-cover top. It's
                    # multiplied by the density as a kludge
                    LC_Angle = degrees(atan(VH / LC_Distance) * Vdens)
                    if not j or LC_Angle_Max < LC_Angle:
                        LC_Angle_Max = LC_Angle
                    if j == 3: VTS_Total += LC_Angle_Max # Add angle at end of each zone calculation
                    rip += RE,
                node.ShaderList += (max(T_Full), ElevationList[i], max(T_None), rip, T_Full),
            node.ViewToSky = 1 - VTS_Total / (7 * 90)

    def BuildZonesLidar(self):
        """Build zones if we are using LiDAR data"""
        #self.CheckEarlyQuit()
        #raise NotImplementedError("LiDAR not yet implemented")
        ################### under construction - copied from BuildZonesNormal
        #Tried to keep in the same general form as BuildZonesNormal so blame Metta
        self.CheckEarlyQuit()
        vheight = []
        elevation = []
        average = lambda x:sum(x)/len(x)

        keys = self.Reach.keys()
        keys.sort(reverse=True) # Downstream sorted list of stream kilometers
        self.PB("Translating LULC Data")
        for i in xrange(7, 36): # For each column of LULC data
            col = self.GetColumn(i, "TTools Data")[5:] # LULC column
            elev = self.GetColumn(i+28,"TTools Data")[5:] # Shift by 28 to get elevation column
            # Make a list from the LC codes from the column, then send that to the multiplier
            # with a lambda function that averages them appropriately. Note, we're averaging over
            # the values (e.g. density) not the actual code, which would be meaningless.
            try:
                vheight.append(self.multiplier([x for x in col], average))
            except KeyError, (stderr):
                raise Exception("Vegetation height error" % stderr.message)
            if i>7:  #We don't want to read in column AJ -Dan
                elevation.append(self.multiplier(elev, average))
            self.PB("Reading vegetation heights", i, 36)
        # We have to set the emergent vegetation, so we strip those off of the iterator
        # before we record the zones.
        for i in xrange(len(keys)):
            node = self.Reach[keys[i]]
            node.VHeight = vheight[0][i]
            node.VDensity = IniParams["lcdensity"]
            node.Overhang = IniParams["lcoverhang"]

        # Average over the topo values
        topo_w = self.multiplier(self.GetColumn(4, "TTools Data")[5:], average)
        topo_s = self.multiplier(self.GetColumn(5, "TTools Data")[5:], average)
        topo_e = self.multiplier(self.GetColumn(6, "TTools Data")[5:], average)

        # ... and you thought things were crazy earlier! Here is where we build up the
        # values for each node. This is culled from earlier version's VB code and discussions
        # to try to simplify it... yeah, you read that right, simplify it... you should've seen in earlier!
        for h in xrange(len(keys)):
            self.PB("Building VegZones", h, len(keys))
            node = self.Reach[keys[h]]
            VTS_Total = 0 #View to sky value
            LC_Angle_Max = 0
            # Now we set the topographic elevations in each direction
            node.TopoFactor = (topo_w[h] + topo_s[h] + topo_e[h])/(90*3) # Topography factor Above Stream Surface
            # This is basically a list of directions, each with a sort of average topography
            ElevationList = (topo_e[h],
                             topo_e[h],
                             0.5*(topo_e[h]+topo_s[h]),
                             topo_s[h],
                             0.5*(topo_s[h]+topo_w[h]),
                             topo_w[h],
                             topo_w[h])
            # Sun comes down and can be full-on, blocked by veg, or blocked by topography. Earlier implementations
            # calculated each case on the fly. Here we chose a somewhat more elegant solution and calculate necessary
            # angles. Basically, there is a minimum angle for which full sun is calculated (top of trees), and the
            # maximum angle at which full shade is calculated (top of topography). Anything in between these is an
            # angle for which sunlight is passing through trees. So, for each direction, we want to calculate these
            # two angles so that late we can test whether we are between them, and only do the shading calculations
            # if that is true.

            for i in xrange(7): # Iterate through each direction
                T_Full = () # lowest angle necessary for full sun
                T_None = () # Highest angle necessary for full shade
                rip = () # Riparian extinction, basically the amount of loss due to vegetation shading
                for j in xrange(4): # Iterate through each of the 4 zones
                    Vheight = vheight[i*4+j+1][h]
                    if Vheight < 0 or Vheight is None or Vheight > 120:
                        raise Exception("Vegetation height (value of %s in TTools Data) must be greater than zero and less than 120 meters (when LiDAR = True)" % `Vheight`)
                    Vdens = IniParams["lcdensity"]
                    Overhang = IniParams["lcoverhang"]
                    Elev = elevation[i*4+j][h]

                    if not j: # We are at the stream edge, so start over
                        LC_Angle_Max = 0 # New value for each direction
                    else:
                        Overhang = 0 # No overhang away from the stream
                    ##########################################################
                    # Calculate the relative ground elevation. This is the
                    # vertical distance from the stream surface to the land surface
                    SH = Elev - node.Elevation
                    # Then calculate the relative vegetation height
                    VH = Vheight + SH
                    # Calculate the riparian extinction value
                    try:
                        RE = -log(1-Vdens)/10
                    except OverflowError:
                        if Vdens == 1: RE = 1 # cannot take log of 0, RE is full if it's zero
                        else: raise
                    # Calculate the node distance
                    LC_Distance = IniParams["transsample"] * (j + 0.5) #This is "+ 0.5" because j starts at 0.
                    # We shift closer to the stream by the amount of overhang
                    # This is a rather ugly cludge.
                    if not j: LC_Distance -= Overhang
                    if LC_Distance <= 0:
                        LC_Distance = 0.00001
                    # Calculate the minimum sun angle needed for full sun
                    T_Full += degrees(atan(VH/LC_Distance)), # It gets added to a tuple of full sun values
                    # Now get the maximum of bank shade and topographic shade for this
                    # direction
                    T_None += degrees(atan(SH/LC_Distance)), # likewise, a tuple of values
                    ##########################################################
                    # Now we calculate the view to sky value
                    # LC_Angle is the vertical angle from the surface to the land-cover top. It's
                    # multiplied by the density as a kludge
                    LC_Angle = degrees(atan(VH / LC_Distance) * Vdens)
                    if not j or LC_Angle_Max < LC_Angle:
                        LC_Angle_Max = LC_Angle
                    if j == 3: VTS_Total += LC_Angle_Max # Add angle at end of each zone calculation
                    rip += RE,
                node.ShaderList += (max(T_Full), ElevationList[i], max(T_None), rip, T_Full),
            node.ViewToSky = 1 - VTS_Total / (7 * 90)

    def GetLandCoverCodes(self):
        """Return the codes from the Land Cover Codes worksheet as a dictionary of dictionaries"""
        self.CheckEarlyQuit()
        codes = self.GetColumn(1, "Land Cover Codes")[3:]
        height = self.GetColumn(2, "Land Cover Codes")[3:]
        dens = self.GetColumn(3, "Land Cover Codes")[3:]
        over = self.GetColumn(4, "Land Cover Codes")[3:]
        # make a list of lists with values: [(height[0], dens[0], over[0]), (height[1],...),...]
        vals = [tuple([j for j in i]) for i in zip(height,dens,over)]
        data = {}

        for i in xrange(len(codes)):
            # Each code is a tuple in the form of (VHeight, VDensity, Overhang)
            data[codes[i]] = vals[i]
            if vals[i][0] != None and (vals[i][1] < 0 or vals[i][1] > 1):
                raise Exception("Vegetation Density (value of %s in Land Cover Codes) must be >= 0.0 and <= 1.0" % `vals[i][1]`)
        return data

    def InitializeNode(self, node):
        """Perform some initialization of the StreamNode, and write some values to spreadsheet"""
        # Initialize each nodes tribs dictionary to a tuple
        for time in self.flowtimelist:
            node.Q_tribs[time] = ()
            node.T_tribs[time] = ()
        ##############################################################
        #Now that we have a stream node, we set the node's dx value, because
        # we have most nodes that are long-sample-distance times multiple,
        node.dx = IniParams["dx"] # Nodes distance step.
        node.dt = IniParams["dt"] # Set the node's timestep... this may have to be adjusted to comply with stability
        # Find the earliest temperature boundary condition
        mindate = min(self.T_bc.keys())
        if self.run_type == 2: # Running hydraulics only
            node.T, node.T_prev, node.T_sed = 0.0, 0.0, 0.0
        else:
            if self.T_bc[mindate] is None:
                # Shade-a-lator doesn't need a boundary condition
                if self.run_type == 1: self.T_bc[mindate] = 0.0
                else:  raise Exception("Boundary temperature conditions cannot be blank")
            node.T = self.T_bc[mindate]
            node.T_prev = self.T_bc[mindate]
            node.T_sed = self.T_bc[mindate]
        #we're in shadealator if the runtype is 1. Since much of the heat
        # math is coupled to the shade math, we have to make sure the hydraulic
        # values are not zero or blank because they'll raise ZeroDivisionError
        if self.run_type ==1:
            for attr in ["d_w", "A", "P_w", "W_w", "U", "Disp","Q_prev","Q",
                         "SedThermDiff","SedDepth","SedThermCond"]:
                if (getattr(node, attr) is None) or (getattr(node, attr) == 0):
                    setattr(node, attr, 0.01)
        node.Q_hyp = 0.0 # Assume zero hyporheic flow unless otherwise calculated
        node.E = 0 # Same for evaporation
    def QuitMessage(self):
        b = buttonbox("Do you really want to quit Heat Source", "Quit Heat Source", ["Cancel", "Quit"])
        if b == "Quit":
            raise Exception("Model stopped user.")
        else: return
