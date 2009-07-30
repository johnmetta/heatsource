from __future__ import division
from time import ctime, strftime, gmtime
from os.path import join, exists
from os import makedirs
from copy import deepcopy
from pywintypes import Time as pyTime


from ..Dieties.IniParamsDiety import IniParams
from ..Dieties.ChronosDiety import Chronos

from .. import opt
try:
    if opt(__name__):
        import psyco.classes
        object = psyco.classes.psyobj
except ImportError: pass

class Output(object):
    """Data and fileobject storage class"""
    def __init__(self, reach, start_time, run_type):
        # Store a sorted list of StreamNodes. This all could be a bit more abstracted.
        self.nodes = sorted(reach.itervalues(),reverse=True)
        # A reference to the model's starting time (i.e. when spin-up is over)
        self.start_time = start_time

        # run_type is a bit hack-y. If we are running only hydraulics,
        # we fail on division of solar parameters- if running only solar,
        # we fail on hydraulics. This is an easy way to prevent that.
        self.run_type = run_type #0=HS, 1=solar, 2=hydraulics
        # Our first time through, we ignore daily data and don't have stream
        # geometry calculated, so we have switches for those (which is a bit dumb)
        self.first_hour = True
        self.first_day = True

        # Filenames and descriptions for each of the output files
        desc = {}
        if run_type < 2:
            desc["Heat_Cond"] = "Streambed Conduction Flux (w/sq m)"
            desc["Heat_Conv"] = "Convection Flux (w/sq m)"
            desc["Heat_Evap"] = "Evaporation Flux (w/sq m)"
            desc["Heat_SR1"] = "Potential Solar Radiation Flux (w/sq m)"
            desc["Heat_SR4"] = "Surface Solar Radiation Flux (w/sq m)"
            desc["Heat_SR6"] = "Received Solar Radiation Flux (w/sq m)"
            desc["Heat_TR"] = "Thermal Radiation Flux (w/sq m)"
            desc["Shade"] = "Effective Shade"
            desc["VTS"] = "View to Sky"
        if run_type != 1:
            desc["Hyd_DA"] = "Ave Depth (m)"
            desc["Hyd_DM"] = "Max Depth (m)"
            desc["Hyd_Flow"] = "Flow Rate (cms)"
            desc["Hyd_Hyp"] = "Hyporheic Exchange (cms)"
            desc["Hyd_Vel"] = "Flow Velocity (m/s)"
            desc["Hyd_WT"] = "Top Width (m)"
        if not run_type:
            desc["Rate_Evap"] = "Evaporation Rate (mm/hr)"
            desc["Temp_H20"] = "Stream Temperature (*C)"
            desc["Temp_Sed"] = "Sediment Temperature (*C)"
            desc["Hyd_Disp"] = "Hydraulic Dispersion (m2/s)"

        # Storage dictionary for the data.
        self.data = {}
        for name in desc.keys():
            self.data[name] = {}
        # make a deepcopy of the empty variables dictionary for use later
        self.empty_vars = deepcopy(self.data)
        # Empty dictionary to store file objects
        self.files = {}

        # Here we build up the self.files attribute by cycling through the
        # filenames and descriptions
        for key in desc.iterkeys():
            # String concatenation takes up a bit of time, but still a lot less
            # than writing to a file each time.
            header = "Heat Source Hourly Output File:  "
            header += desc[key]
            header += "     File created on %s\n\n" % ctime()
            header += "Datetime".ljust(14)
            # Grab a joined list of left justified strings of all the kilometers
            header += "".join([("%0.3f" % x.km).ljust(14) for x in self.nodes])
            header += "\n"
            # Now create a file object in the dictionary, and write the header
            self.files[key] = open(join(IniParams["outputdir"], key + ".txt"), 'w')
            self.files[key].write(header)

    def close(self):
        # Flush the rest of the values from the dataset by flushing the
        # daily values and by calling the write() method
        # self.write(self.run_type < 2)  #commented out this line so shade wouldn't output last day twice - DT
        # Then close all of the file objects cleanly
        [f.close() for f in self.files.itervalues()]

    def __call__(self, time, hour):
        """Call the storage method with a time and an hour"""
        # Ignore this if we're still spinning up of if this is the first
        # hour run (because we don't have channel geometry calculated).
        if time < self.start_time: return
        if self.first_hour:
            self.first_hour = False
            #return
        # Create an Excel-friendly time string
        timestamp = ("%0.6f" % float(time/86400 + 25569)).ljust(14)
        # Localize variables to save a bit of time
        nodes = self.nodes
        data = self.data
        # Cycle through each datatype, creating a list of values
        # corresponding to the nodes for this timestamp. Thus, each
        # timestamp conforms to a single line, and each element in the
        # list comprehension conforms to a column. List comprehensions
        # are generally fast (more optimized by the underlying C code)
        # than for loops.

        # Run only with solar
        if self.run_type < 2:
            data["Heat_Cond"][timestamp] = [x.F_Conduction for x in nodes]
            data["Heat_Conv"][timestamp] = [x.F_Convection for x in nodes]
            data["Heat_Evap"][timestamp] = [x.F_Evaporation for x in nodes]
            data["Heat_SR1"][timestamp] = [x.F_Solar[1] for x in nodes]
            data["Heat_SR4"][timestamp] = [x.F_Solar[4] for x in nodes]
            data["Heat_SR6"][timestamp] = [x.F_Solar[6] for x in nodes]
            data["Heat_TR"][timestamp] = [x.F_Longwave for x in nodes]
        # Run only with hydro
        if self.run_type != 1:
            data["Hyd_DA"][timestamp] = [(x.A / x.W_w) for x in nodes]
            data["Hyd_DM"][timestamp] = [x.d_w for x in nodes]
            data["Hyd_Flow"][timestamp] = [x.Q for x in nodes]
            data["Hyd_Hyp"][timestamp] = [x.Q_hyp for x in nodes]
            data["Hyd_Vel"][timestamp] = [x.U for x in nodes]
            data["Hyd_WT"][timestamp] = [x.W_w for x in nodes]
        # Run only with both solar and hydro
        if not self.run_type:
            data["Rate_Evap"][timestamp] = [(x.E / x.dx / x.W_w * 3600 * 1000) for x in nodes] #TODO: Check
            data["Temp_H20"][timestamp] = [x.T for x in nodes]
            data["Temp_Sed"][timestamp] = [x.T_sed for x in nodes]
            data["Hyd_Disp"][timestamp] = [x.Disp for x in nodes]

        # Zero for an hour means a new day, so we add daily outputs
        # and write to the file. Writing only every day saves us
        # 24xF file accesses where F=len(self.files). Each file access
        # has quite a bit of overhead, so we lump them. It's "A Good Thing."
        if hour == 23:
            self.write(self.run_type < 2)

    def daily(self, timestamp):
        """Compile and store data that is collected every hour"""
        nodes = self.nodes
        self.data["Shade"][timestamp] = [((x.F_DailySum[1] - x.F_DailySum[4]) / x.F_DailySum[1]) for x in nodes]
        self.data["VTS"][timestamp] = [x.ViewToSky for x in nodes]
        # If there's no hour, we're at the beginning of a day, so we write the values
        # to a file.

    def write(self, daily):
        if daily: # don't call for hydraulics
            self.daily(("%0.6f" % float(pyTime(Chronos()))).ljust(14))
        # localize the
        data = self.data
        # Cycle through the file objects
        for name, fileobj in self.files.iteritems():
            # Each time is a single line, so we want to iterate over all the times
            # stored so far. We can do this because everytime we store data, we
            # append the time string to self.times
            line = ""
            timelist = sorted(data[name].keys())
            for timestamp in timelist:
                line += timestamp
                line += "".join([("%0.4f" % x).ljust(14) for x in data[name][timestamp]])
                line += "\n"
            # finally, write all the lines to the file
            fileobj.write(line)
        del data
        # Now empty out the dictionary by simply copying a new one.
        self.data = deepcopy(self.empty_vars)