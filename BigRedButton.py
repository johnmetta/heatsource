"""BigRedButton holds the main Heat Source model controls

BigRedButton holds the main Heat Source model controls.
The ModelControl class is basically a one-off hack that
loads and controls the model run. The other functions are
supplied for use by the Excel interface, which imports this
module in order to run the model from Excel.
"""
from __future__ import with_statement, division

# Built-in modules
from itertools import count
from traceback import print_exc, format_tb
from sys import exc_info
from os.path import join, exists
from os import unlink
from win32com.client import Dispatch
from win32gui import PumpWaitingMessages
from Utils.easygui import msgbox, buttonbox
from time import time as Time
from time import ctime, gmtime

# Heat Source modules
from Dieties.IniParamsDiety import IniParams
from Excel.ExcelInterface import ExcelInterface
from Dieties.ChronosDiety import Chronos
from Utils.Logger import Logger
from Utils.Output import Output as O
from HSmodule import HeatSourceError
from __version__ import version_info

from . import opt

class ModelControl(object):
    """Main model control class for Heat Source.

    While it works, this class is basically a one-off hack
    that was designed as an interim solution to model control.
    As it is, it's a fairly dumb method which simply grabs
    a list of StreamNodes from the ExcelInterface class,
    and loops through them using the iterator inside
    ChronosDiety to advance the time. Ideally, this would be
    a more sophisticated system that could hold and control
    multiple stream reaches, modeling upstream reaches
    and using the results as tributary inputs to the down-
    stream reaches. For that, we'd likely need a separate
    Reach class. Since this was essentially an interim
    solution to the problem, don't hesitate to improve it.
    """
    def __init__(self, spreadsheet, run_type=0):
        """ModelControl(spreadsheet, run_type) -> Class instance

        Spreadsheet is the path to an excel sheet containing the data.
        run_type is one of 0,1,2 for Heat Source, Solar only, or
        hydraulics only, respectively.
        """
        # TODO: Fix the logger so it actually works
        self.ErrLog = Logger

        # Create an ExcelInterface instance. Here, we could just grab
        # the Reach and PB (progress bar) attributes and then release it,
        # but internal use has suggested that it's nice to keep ownership
        # of the sheet throughout the model run.
        self.HS = ExcelInterface(spreadsheet, self.ErrLog, run_type)

        # This is the list of StreamNode instances- we sort it in reverse
        # order because we number stream kilometer from the mouth to the
        # headwater, but we want to run the model from headwater to mouth.
        self.reachlist = sorted(self.HS.Reach.itervalues(), reverse=True)

        # This if statement prevents us from having to test every timestep
        # We just call self.run_all(), which is a classmethod pointing to
        # the correct method.
        if run_type == 0: self.run_all = self.run_hs
        elif run_type == 1: self.run_all = self.run_sh
        elif run_type == 2: self.run_all = self.run_hy
        else: raise Exception("Bad run_type: %i. Must be 0, 1 or 2" %`self.run_type`)
        # Create a Chronos iterator that controls all model time.
        Chronos.Start(start = IniParams["modelstart"],
                      stop = IniParams["modelend"],
                      dt = IniParams["dt"],
                      spin = IniParams["flushdays"],
                      offset = IniParams["offset"])
        # This is the output class, which is essentially just a list
        # of file objects and an append method which writes to them
        # every so often.
        self.Output = O(self.HS.Reach, IniParams["modelstart"], run_type)

    def Run(self):
        """Run the model one time

        Use the Chronos instance and list of StreamNodes to cycle
        through each timestep and spacestep, calling the appropriate
        StreamNode functions to calculate heat and hydraulics."""
        time = Chronos.TheTime # Current time of the Chronos clock (i.e. this timestep)
        stop = Chronos.stop # Stop time for Chronos
        start = Chronos.start # Start time for Chronos, model start, not flush/spin start.
        flush = start-(IniParams["flushdays"]*86400) # in seconds
        # Number of timesteps is based on the division of the timesteps into the hour. In other words
        # 1 day with a 1 minute dt is 1440 timesteps, while a 3 minute dt is only 480 timesteps. Thus,
        # We define the timesteps by dividing dt (now in seconds) by 3600
        timesteps = (stop-flush)/IniParams["dt"]
        cnt = count() # Counter iterator for counting current timesteps passed
        out = 0 # Volume of water flowing out of mouth (for simple mass balance)
        time1 = Time() # Current computer time- for estimating total model runtime
        # Localize run_type for a bit more speed
        ################################################################
        # So, it's simple and stupid. We basically just cycle through the time
        # until we get to the model stop time, and each cycle, we call a method
        # that cycles through the StreamNodes, and calculates them in order.
        # A smarter way would be to thread this so we can start calculating
        # the second timestep (using another CPU or core) while the first one
        # is still unfinished.
        while time <= stop:
            year, month, day, hour, minute, second, JD, offset, JDC = Chronos.TimeTuple()
            # zero hour+minute+second means first timestep of new day
            # We want to zero out the daily flux sum at this point.
            if not (hour + minute + second):
                for nd in self.reachlist: nd.F_DailySum = [0]*5

            # Back to every timestep level of the loop. Here we wrap the call to
            # run_all() in a try block to catch the exceptions thrown.
            try:
                # Note that all of the run methods have to have the same signature
                self.run_all(time, hour, minute, second, JD, JDC)
            # Shit, there's a problem, throw an exception up using a graphical window.
            except HeatSourceError, (stderr):
                msg = "At %s and time %s\n"%(self, Chronos.PrettyTime())
                try:
                    msg += stderr+"\nThe model run has been halted. You may ignore any further error messages."
                except TypeError:
                    msg += `stderr`+"\nThe model run has been halted. You may ignore any further error messages."
                msgbox(msg)
                # Then just die
                raise SystemExit
                        # If minute and second are both zero, we are at the top of the hour. Performing

            # The following house keeping tasks each hours saves us enormous amounts of
            # runtime overhead over doing it every timestep.
            if not (minute + second):
                ts = cnt.next() # Number of actual timesteps per tick
                hr = 60/(IniParams["dt"]/60) # Number of timesteps in one hour
                # This writes a line to the status bar of Excel.
                self.HS.PB("%i of %i timesteps"% (ts*hr, timesteps))
                # Update the Excel status bar when the queue is free
                PumpWaitingMessages()
                # Call the Output class to update the textfiles. We call this every
                # hour and store the data, then we write to file every day. Limiting
                # disk access saves us considerable time.
                self.Output(time, hour)
                # Check to see if the user pressed the stop button. Pretty crappy kludge here- VB code writing an
                # empty file- but I basically got to lazy to figure out how to interact with the underlying
                # COM API without using a threading interface.
                if exists("c:\\quit_heatsource"):
                    unlink("c:\\quit_heatsource")
                    if QuitMessage():
                        break

            # We've made it through the entire stream without an error, so we update our mass balance
            # by adding the discharge of the mouth...
            out += self.reachlist[-1].Q
            # and tell Chronos that we're moving time forward.
            time = Chronos(True)

        # So, here we are at the end of a model run. First we calculate how long all of this took
        total_time = (Time() - time1) / 60
        # Calculate the mass balance inflow
        balances = [x.Q_mass for x in self.reachlist]
        total_inflow = sum(balances)
        # Calculate how long the model took to run each timestep for each spacestep.
        # This is how long (on average) it took to calculate a single StreamNode's information
        # one time. Ideally, for performance and impatience reasons, we want this to be somewhere
        # around or less than 1 microsecond.
        microseconds = (total_time/timesteps/len(self.reachlist))*1e6
        message = "Finished in %0.1f minutes (spent %0.3f microseconds in each stream node). Water Balance: %0.3f/%0.3f" %\
                    (total_time, microseconds, total_inflow, out)
        # Close out Output (there's a bit of a lag when it writes the file,
        # so we do this before the final message so people don't accidentally
        # access the file and screw up the buffer)
        self.Output.close()
        # write that final message to the Excel status bar
        self.HS.PB(message)
        # Hopefully, Python's cyclic garbage collection takes care of the rest :)

    #############################################################
    # three different versions of the run() routine, depending on the run_type
    # We use list comprehension because it's slightly faster than a for loop,
    # and we want to eek out all the speed we can.
    def run_hs(self, time, H, M, S, JD, JDC):
        """Call both hydraulic and solar routines for each StreamNode"""
        [x.CalcDischarge(time) for x in self.reachlist]
        [x.CalcHeat(time, H, M, S, JD, JDC) for x in self.reachlist]
        [x.MacCormick2(time) for x in self.reachlist]

    def run_hy(self, time, H, M, S, JD, JDC):
        """Call hydraulic routines for each StreamNode"""
        [x.CalcDischarge(time) for x in self.reachlist]

    def run_sh(self, time, H, M, S, JD, JDC):
        """Call solar routines for each StreamNode"""
        [x.CalcHeat(time, H, M, S, JD, JDC, True) for x in self.reachlist]


def QuitMessage():
    """Throw up a confirmation box to make sure we didn't hit the quit button accidentally"""
    b = buttonbox("Do you really want to quit Heat Source", "Quit Heat Source", ["Cancel", "Quit"])
    if b == "Quit": return True
    else: return False


def RunHS(sheet):
    """Run full model"""
    try:
        HSP = ModelControl(sheet)
        HSP.Run()
        del HSP
    except Exception, stderr:
        f = open("c:\\HSError.txt", "w")
        print_exc(file=f)
        f.close()
        msgbox("".join(format_tb(exc_info()[2]))+"\nSynopsis: %s"%stderr, "HeatSource Error", err=True)
def RunSH(sheet):
    """Run solar routines only"""
    try:
        HSP = ModelControl(sheet, 1)
        HSP.Run()
    except Exception, stderr:
        f = open("c:\\HSError.txt", "w")
        print_exc(file=f)
        f.close()
        msgbox("".join(format_tb(exc_info()[2]))+"\nSynopsis: %s"%stderr, "HeatSource Error", err=True)
def RunHY(sheet):
    """Run hydraulics only"""
    try:
        HSP = ModelControl(sheet, 2)
        HSP.Run()
    except Exception, stderr:
        f = open("c:\\HSError.txt", "w")
        print_exc(file=f)
        f.close()
        msgbox("".join(format_tb(exc_info()[2]))+"\nSynopsis: %s"%stderr, "HeatSource Error", err=True)

try:
    if opt(__name__):
        import psyco
        psyco.bind(ModelControl.Run)
        psyco.bind(ModelControl.run_sh)
        psyco.bind(ModelControl.run_hy)
        psyco.bind(ModelControl.run_hs)
except ImportError: pass
