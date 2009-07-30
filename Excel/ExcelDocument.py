"""A class containing methods to manipulate Excel documents
"""
from __future__ import division
from win32com.client import constants, Dispatch, gencache
from pythoncom import CoInitialize, CoUninitialize
from pywintypes import com_error
from os.path import exists
from os import remove

from ..Dieties.IniParamsDiety import IniParams

borderTop = 3
borderBottom = 4
borderLeft = 1
borderRight = 2
borderSolid = 1
borderDashed = 2
borderDotted = 3
colorBlack = 1
directionUp = -4162
directionDown = -4121
directionLeft = -4131
directionRight = -4152

from .. import opt
try:
    if opt(__name__):
        import psyco.classes
        Tobject = psyco.classes.psyobj
    else: Tobject = object
except ImportError: Tobject = object

class TextPB(Tobject):
    def __init__(self):
        self.bar = "---->"
        self.text = list(self.bar) + [" "]*60

    def __call__(self, msg="", num=None, total=None):
        self.text.insert(0, self.text.pop())
        if num and total:
            msg = "%s %i%%" %(msg, (num/total)*100)
        elif num:
            msg = "%s %i%%" %(msg, num)
        return "%s  | %s" %("".join(self.text), msg)

# psyco fills memory if this class is optimized, so we just use
# the normal Python object
class ExcelDocument(object):
    """
    This is a recipe class culled from ASPN (ActiveState). It implements
    many of the necessary methods to manipulate an Excel spreadsheet using
    the COM interface and the Python win32com library. The methods
    excelize and deExcelize were taken from some other, now forgotten, source.
    """
    def __init__(self, filename):
        # The following code assumes that there is an active workbook (i.e. that Excel is running
        # and a workbook is open and active. This is because we should be calling this from the VB
        # macro which would activate the Heat Source workbook. The reason we do not call Open(excelfile)
        # is that we want to be able to catch unsaved changes, which is possible only if we catch
        # a reference to the active workbook.

        # EnsureDispatch() will start the Excel.Application, optionally doing a little makepy magic
        # and generating some behind the scenes including some constants and classes that we use
        # such as xlLastCell. Calling only Dispatch() will assume we have run makepy by hand, which
        # may not be the case.
        self.app = gencache.EnsureDispatch("Excel.Application", 0)
        # This flag tells us whether we should quit the Excel application when this class' destructor
        # is called. Doing so makes Excel shut down when the user may not want this to occur, so
        # we've set it to False, but kept the flag for future changes.
        self.quit_excel = False
        # TextPB is a progress bar like class that creates a moving arrow and a message in the status
        # bar.
        self.PBtext = TextPB()
        # If we don't have an active workbook, open one
        if not self.app.ActiveWorkbook:
            self.quit_excel = True
            self.Open(filename)
            self.app.Visible = True

    def __del__(self):
        if self.quit_excel:
            self.app.Quit()
        del self.app
        del self.PBtext
    def PB(self, message, num=None, divisor=None):
        """Send a message to the status bar, or push the arrow forward one tick"""
        try:
            self.app.StatusBar = self.PBtext(message, num, divisor)
        except com_error: # win32 throws an exception if we are playing with excel and try to update the statusbar
            pass

    def New(self, filename=None):
        """
        Create a new Excel workbook. If 'filename' specified,
        use the file as a template.
        """
        self.app.Workbooks.Add(filename)

    def Open(self, filename):
        """
        Open an existing Excel workbook for editing.
        """
        self.app.Workbooks.Open(filename)

    def SetSheet(self, sheet):
        """
        Set the active worksheet.
        """
        if not isinstance(sheet, str) and not isinstance(sheet, int):
            raise Exception("Sheet must be set to an integer or a name")
        self.sheet = sheet

    def GetSheet(self, sheet):
        """
        Return reference to the sheet
        """
        return self.app.ActiveWorkbook.Worksheets(sheet)

    def GetColumn(self, col, sheet):
        """
        Return a column of data
        """
        strip = lambda x: x[0] if len(x) == 1 else x
        lst = [strip(i) for i in self.GetValue((1, col, self.LastRow(sheet), col), sheet)]
        return tuple(lst)

    def GetRange(self, range, sheet=None):
        """
        Get a range object for the specified range or single cell.
        """
        sheet = sheet if sheet else self.sheet
        if isinstance(range, list) or isinstance(range, tuple):
            #We have a sequence, lets make sure it's either (r1,c1,r2,c2) or ((r1,c1),(r2,c2))
            if len(range) == 4: # (r1,c1,r2,c2)
                rng = "%s%i:%s%i" % (self.excelize(range[1]), range[0], self.excelize(range[3]), range[2])
            elif len(range) == 2: # ((r1,c1),(r2,c2))
                if (isinstance(range[0], list) or isinstance(range[0], tuple)) and \
                    (isinstance(range[1], list) or isinstance(range[1], tuple)):
                    rng = "%s%i:%s%i" % (self.excelize(range[0][1]), range[0][0], \
                                         self.excelize(range[1][1]), range[1][0])
                elif isinstance(range[0], int) and isinstance(range[1], int):
                    rng = "%s%i" % (self.excelize(range[1]), range[0])
                else: raise Exception
            else: raise Exception
        elif isinstance(range, str): rng = range
        else: raise Exception

        return self.app.ActiveWorkbook.Sheets(sheet).Range(rng)

    def SetValue(self, cell, value='', sheet=None):
        """
        Set the value of 'cell' to 'value'.
        """
        self.GetRange(cell, sheet).Value = value

    def GetValue(self, cell, sheet=None):
        """
        Get the value of 'cell'.
        """
        return self.GetRange(cell, sheet).Value

    def GetUsedRange(self, sheet=None):
        """
        Return the data for the entire used range.
        """
        return self.app.ActiveWorkbook.Sheets(sheet).UsedRange.Value

    def LastRow(self, sheet=None):
        return self.app.ActiveWorkbook.Sheets(sheet).Cells.SpecialCells(constants.xlLastCell).Row
    def LastColumn(self, sheet=None):
        return self.app.ActiveWorkbook.Sheets(sheet).Cells.SpecialCells(constants.xlLastCell).Column

    def UsedRange(self, sheet=None):
        """
        Return the used range of the data in the form of (endcol,endrow)
        """
        return (self.LastColumn(sheet), self.LastRow(sheet))

    def Clear(self, cell, sheet=None):
        self.GetRange(cell, sheet).Clear()

    def SetBorder(self, range, side, line_style=borderSolid, color=colorBlack):
        """
        Set a border on the specified range of cells or single cell.
        'range' = range of cells or single cell
        'side' = one of borderTop, borderBottom, borderLeft, borderRight
        'line_style' = one of borderSolid, borderDashed, borderDotted, others?
        'color' = one of colorBlack, others?
        """
        range = self.GetRange(range).Borders(side)
        range.LineStyle = line_style
        range.Color = color

    def Sort(self, range, key_cell):
        """
        Sort the specified 'range' of the activeworksheet by the
        specified 'key_cell'.
        """
        range.Sort(Key1=self.GetRange(key_cell), Order1=1, Header=0, OrderCustom=1, MatchCase=False, Orientation=1)

    def HideRow(self, row, hide=True):
        """
        Hide the specified 'row'.
        Specify hide=False to show the row.
        """
        self.GetRange('a%s' % row).EntireRow.Hidden = hide

    def HideColumn(self, column, hide=True):
        """
        Hide the specified 'column'.
        Specify hide=False to show the column.
        """
        self.GetRange('%s1' % column).EntireColumn.Hidden = hide

    def DeleteRow(self, row, shift=directionUp):
        """
        Delete the entire 'row'.
        """
        self.GetRange('a%s' % row).EntireRow.Delete(Shift=shift)

    def DeleteColumn(self, column, shift=directionLeft):
        """
        Delete the entire 'column'.
        """
        self.GetRange('%s1' % column).EntireColumn.Delete(Shift=shift)

    def FitColumn(self, column):
        """
        Resize the specified 'column' to fit all its contents.
        """
        self.GetRange('%s1' % column).EntireColumn.AutoFit()

    def Save(self):
        """
        Save the active workbook.
        """
        self.app.ActiveWorkbook.Save()

    def SaveAs(self, filename, delete_existing=False):
        """
        Save the active workbook as a different filename.
        If 'delete_existing' is specified and the file already
        exists, it will be deleted before saving.
        """
        if delete_existing and exists(filename):
            remove(filename)
        self.app.ActiveWorkbook.SaveAs(filename)

    def PrintOut(self):
        """
        Print the active workbook.
        """
        self.app.Application.PrintOut()

    def Close(self):
        """
        Close the active workbook.
        """
        if self.app.ActiveWorkbook: self.app.ActiveWorkbook.Close(SaveChanges=0)

    def Quit(self):
        """
        Quit Excel.
        """
        return self.app.Quit()

    def chars(self):
        """
        Returns an iterator object that yields each charector of the
        english alphabet in capitals.
        """
        for i in range(26):
            yield chr(65 + i)
    def excelIter(self):
        """
        Returns an iterator that yields each excel formated column
        number in ascending order.
        """
        for ch in self.chars():
            yield ch
        for exCh in excelIter():
            for ch in self.chars():
                yield exCh+ch
    def excelize(self, n):
        """
        Returns excel formated column number for n>=0.
        """
        div = int(n/26)
        if div==0:
            if isinstance(n, float): n = int(n)
            return chr(65+n)
        else:
            return self.excelize(div-1)+chr(65+n%26)
    def deExcelize(self, s):
        """
        Returns an integer value for an excel formated column value.
        Expects a string containing only English letters
        """
        s = s.upper() if not s.isupper() else s
        rem = s[:-1]
        if rem == "":
            return ord(s) - 65
        else:
            return 26*(self.deExcelize(s[:-1])+1) + ord(s[-1]) - 65
