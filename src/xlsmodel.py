
import globals, node, formula


class ModelType:
    Workbook = 0
    Unknown = 999


class ModelBase(object):
    def __init__ (self, modelType=ModelType.Unknown):
        self.modelType = modelType


class Workbook(ModelBase):

    def __init__ (self):
        ModelBase.__init__(self, ModelType.Workbook)

        # public members
        self.encrypted = False

        # private members
        self.__sheets = []

    def appendSheet (self):
        n = len(self.__sheets)
        if n == 0:
            self.__sheets.append(WorkbookGlobal())
        else:
            self.__sheets.append(Worksheet(n-1))

        return self.getCurrentSheet()

    def getWorkbookGlobal (self):
        return self.__sheets[0]

    def getCurrentSheet (self):
        return self.__sheets[-1]

    def createDOM (self):
        nd = node.Element('workbook')
        nd.setAttr('encrypted', self.encrypted)
        n = len(self.__sheets)
        if n == 0:
            return

        wbglobal = self.__sheets[0]
#       nd.appendChild(wbglobal.createDOM(self))
        for i in xrange(1, n):
            sheet = self.__sheets[i]
            sheetNode = sheet.createDOM(self)
            nd.appendChild(sheetNode)
            if i > 0:
                data = wbglobal.getSheetData(i-1)
                sheetNode.setAttr('name', data.name)
                sheetNode.setAttr('visible', data.visible)

        return nd


class SheetBase(object):

    class Type:
        WorkbookGlobal = 0
        Worksheet = 1

    def __init__ (self, modelType):
        self.modelType = modelType
        self.version = None

    def createDOM (self, wb):
        nd = node.Element('sheet')
        return nd


class Supbook(object):
    class Type:
        Self     = 0
        AddIn    = 1
        External = 2
        DDE      = 3
        OLE      = 4
        Unused   = 5

    def __init__ (self, sbType=None):
        self.sbType = sbType

    def getType (self):
        return self.sbType


class SupbookSelf(Supbook):
    def __init__ (self, sheetCount):
        Supbook.__init__(self, Supbook.Type.Self)
        self.sheetCount = sheetCount


class SupbookExternal(Supbook):
    def __init__ (self):
        Supbook.__init__(self, Supbook.Type.External)


class WorkbookGlobal(SheetBase):
    class SheetData:
        def __init__ (self):
            self.name = None
            self.visible = True

    def __init__ (self):
        SheetBase.__init__(self, SheetBase.Type.WorkbookGlobal)

        self.__sheetData = []
        self.__sharedStrings = []
        self.__supbooks = []
        self.__externSheets = []  # tuple (book ID, sheet begin ID, sheet end ID)
        self.__dbRanges = {}      # key: sheet ID (0-based), value: range tokens

    def createDOM (self, wb):
        nd = node.Element('workbook-global')
        return nd

    def appendSheetData (self, data):
        self.__sheetData.append(data)

    def getSheetData (self, i):
        return self.__sheetData[i]

    def appendSharedString (self, sst):
        self.__sharedStrings.append(sst)

    def getSharedString (self, strID):
        if len(self.__sharedStrings) <= strID:
            return None
        return self.__sharedStrings[strID]

    def appendSupbook (self, sb):
        self.__supbooks.append(sb)

    def getSupbook (self, sbID):
        if len(self.__supbooks) <= sbID:
            return None
        return self.__supbooks[sbID]

    def appendExternSheet (self, bookID, sheetBeginID, sheetEndID):
        self.__externSheets.append((bookID, sheetBeginID, sheetEndID))

    def getExternSheet (self, xtiID):
        if len(self.__externSheets) <= xtiID:
            return None
        return self.__externSheets[xtiID]

    def setFilterRange (self, sheetID, tokens):
        self.__dbRanges[sheetID] = tokens

    def getFilterRange (self, sheetID):
        if not self.__dbRanges.has_key(sheetID):
            return None

        return self.__dbRanges[sheetID]


class Worksheet(SheetBase):

    def __init__ (self, sheetID):
        SheetBase.__init__(self, SheetBase.Type.Worksheet)
        self.__rows = {}
        self.__autoFilterArrows = []
        self.sheetID = sheetID

    def setAutoFilterArrowSize (self, arrowSize):
        arrows = []
        for i in xrange(0, arrowSize):
            arrows.append(None)

        # Swap with the new and empty list.
        self.__autoFilterArrows = arrows

    def setCell (self, col, row, cell):
        if not self.__rows.has_key(row):
            self.__rows[row] = {}

        self.__rows[row][col] = cell

    def createDOM (self, wb):
        nd = node.Element('worksheet')
        nd.setAttr('version', self.version)

        # cells
        rows = self.__rows.keys()
        rows.sort()
        for row in rows:
            rowNode = nd.appendElement('row')
            rowNode.setAttr('id', row)
            cols = self.__rows[row].keys()
            for col in cols:
                cell = self.__rows[row][col]
                cellNode = cell.createDOM(wb)
                rowNode.appendChild(cellNode)
                cellNode.setAttr('col', col)

        # autofilter (if exists)
        self.__appendAutoFilterNode(wb, nd)

        return nd

    def __appendAutoFilterNode (self, wb, baseNode):
        if len(self.__autoFilterArrows) <= 0:
            return

        wbg = wb.getWorkbookGlobal()
        tokens = wbg.getFilterRange(self.sheetID)
        parser = formula.FormulaParser2(None, tokens, False)
        parser.parse()
        tokens = parser.getTokens()
        if len(tokens) != 1 or tokens[0].tokenType != formula.TokenType.Area3d:
            return


        tk = tokens[0]
        cellRange = tk.cellRange

        elem = baseNode.appendElement('autofilter')
        elem.setAttr('range', "(col=%d,row=%d)-(col=%d,row=%d)"%(cellRange.firstCol, cellRange.firstRow, cellRange.lastCol, cellRange.lastRow))

        for col in xrange(cellRange.firstCol, cellRange.lastCol+1):
            arrow = elem.appendElement('arrow')
            arrow.setAttr('col', col)
            arrow.setAttr('row', cellRange.firstRow)

class CellBase(object):

    class Type:
        Label   = 0
        Number  = 1
        Formula = 2
        Unknown = 999

    def __init__ (self, modelType):
        self.modelType = modelType


class LabelCell(CellBase):
    def __init__ (self):
        CellBase.__init__(self, CellBase.Type.Label)
        self.strID = None

    def createDOM (self, wb):
        nd = node.Element('label-cell')
        if self.strID != None:
            sst = wb.getWorkbookGlobal().getSharedString(self.strID)
            if sst != None:
                nd.setAttr('value', sst.baseText)
        return nd


class NumberCell(CellBase):
    def __init__ (self, value):
        CellBase.__init__(self, CellBase.Type.Number)
        self.value = value

    def createDOM (self, wb):
        nd = node.Element('number-cell')
        nd.setAttr('value', self.value)
        return nd


class FormulaCell(CellBase):
    def __init__ (self):
        CellBase.__init__(self, CellBase.Type.Formula)
        self.tokens = None

    def createDOM (self, wb):
        nd = node.Element('formula-cell')
        if self.tokens != None:
            s = globals.getRawBytes(self.tokens, True, False)
            nd.setAttr('token-bytes', s)
        return nd
