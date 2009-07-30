import time

class LoggerDiety(object):
    def __init__(self):
        self._last = None
        self._file = None
    def __del__(self):
        try:
            self._file.close()
        except AttributeError:
            pass
    def SetFile(self, filename):
        try:
            self._file = open(filename,"w")
        except IOError:
            raise IOError("Opening output directory failed. Make sure directory exists before running.")
    def __call__(self, message, n=None,t=None): self.write(message)
    def write(self, message):
        if message != self._last:
            t = time.strftime("%H:%M:%S",time.gmtime(time.time()))
            self._file.write("%s-> %s\n"% (t,message))
            self._last = message
    def progress(self): self._file.write(".")
Logger = LoggerDiety()
