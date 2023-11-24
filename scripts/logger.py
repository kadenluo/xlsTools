import utils
from datetime import datetime


@utils.singleton
class Logger():
    def __init__(self, outputObj = None):
        super().__init__()
        self.outputObj = outputObj

    def _printMsg(self, level, pattern, *args):
        detail = "{} [{}] {}".format(datetime.now().strftime('%Y-%m-%d %H:%M:%S'), level, pattern.format(*args))
        if self.outputObj is None:
            print(detail)
        else: 
            self.outputObj.append(detail)

    def info(self, pattern, *args):
        self._printMsg("INFO", pattern, *args)

    def error(self, pattern, *args):
        self._printMsg("ERROR", pattern, *args)

    def critical(self, pattern, *args):
        self._printMsg("CRIT", pattern, *args)

    def clear(self):
        if self.outputObj is not None:
            self.outputObj.clear()