
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
        self.__sheetID = sheetID
        self.__firstDefinedCell = None
        self.__firstFreeCell = None

    def setFirstDefinedCell (self, col, row):
        self.__firstDefinedCell = formula.CellAddress(col, row)

    def setFirstFreeCell (self, col, row):
        self.__firstFreeCell = formula.CellAddress(col, row)

    def setAutoFilterArrowSize (self, arrowSize):
        arrows = []
        for i in xrange(0, arrowSize):
            arrows.append(None)

        # Swap with the new and empty list.
        self.__autoFilterArrows = arrows

    def setAutoFilterArrow (self, filterID, obj):
        self.__autoFilterArrows[filterID] = obj

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

        # table dimension
        if self.__firstDefinedCell != None:
            nd.setAttr('first-defined-cell', self.__firstDefinedCell.getName())

        if self.__firstFreeCell != None:
            nd.setAttr('first-free-cell', self.__firstFreeCell.getName())

        # autofilter (if exists)
        self.__appendAutoFilterNode(wb, nd)

        return nd

    def __appendAutoFilterNode (self, wb, baseNode):
        if len(self.__autoFilterArrows) <= 0:
            # No autofilter in this sheet.
            return

        wbg = wb.getWorkbookGlobal()
        tokens = wbg.getFilterRange(self.__sheetID)
        parser = formula.FormulaParser2(None, tokens)
        parser.parse()
        tokens = parser.getTokens()
        if len(tokens) != 1 or tokens[0].tokenType != formula.TokenType.Area3d:
            # We assume that the database range only has one range token, otherwise
            # we bail out.
            return

        tk = tokens[0]
        cellRange = tk.cellRange

        elem = baseNode.appendElement('autofilter')
        elem.setAttr('range', cellRange.getName())

        for i in xrange(0, len(self.__autoFilterArrows)):
            arrowObj = self.__autoFilterArrows[i]
            if arrowObj == None:
                arrow = elem.appendElement('arrow')
                cell = formula.CellAddress(cellRange.firstCol+i, cellRange.firstRow)
                arrow.setAttr('pos', cell.getName())
            else:
                elem.appendChild(arrowObj.createDOM(wb, cellRange))


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
            parser = formula.FormulaParser2(None, self.tokens)
            parser.parse()
            nd.setAttr('formula', parser.getText())
            s = globals.getRawBytes(self.tokens, True, False)
            nd.setAttr('token-bytes', s)
        return nd


class AutoFilterArrow(object):

    def __init__ (self, filterID):
        self.filterID = filterID
        self.isActive = False
        self.equalString1 = None
        self.equalString2 = None

    def createDOM (self, wb, filterRange):
        nd = node.Element('arrow')
        col = self.filterID + filterRange.firstCol
        row = filterRange.firstRow
        cell = formula.CellAddress(col, row)
        nd.setAttr('pos', cell.getName())
        nd.setAttr('active', self.isActive)
        eqStr = ''
        if self.equalString1 != None:
            eqStr = self.equalString1
        if self.equalString2 != None:
            eqStr += ',' + self.equalString2
        nd.setAttr('equals', eqStr)
        return nd
