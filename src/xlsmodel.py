
import globals, node


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
        if len(self.__sheets) == 0:
            self.__sheets.append(WorkbookGlobal())
        else:
            self.__sheets.append(Worksheet())

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
        nd.appendChild(wbglobal.createDOM())
        for i in xrange(1, n):
            sheet = self.__sheets[i]
            sheetNode = sheet.createDOM()
            nd.appendChild(sheetNode)
            if i > 0:
                data = wbglobal.getSheetData(i-1)
                sheetNode.setAttr('name', data.name)
                sheetNode.setAttr('visible', data.visible)

        return nd


class SheetModelType:
    WorkbookGlobal = 0
    Worksheet = 1


class SheetBase(object):
    def __init__ (self, modelType):
        self.modelType = modelType
        self.version = None

    def createDOM (self):
        nd = node.Element('sheet')
        return nd


class WorkbookGlobal(SheetBase):
    class SheetData:
        def __init__ (self):
            self.name = None
            self.visible = True

    def __init__ (self):
        SheetBase.__init__(self, SheetModelType.WorkbookGlobal)

        self.__sheetData = []

    def createDOM (self):
        nd = node.Element('workbook-global')
        return nd

    def appendSheetData (self, data):
        self.__sheetData.append(data)

    def getSheetData (self, i):
        return self.__sheetData[i]


class Worksheet(SheetBase):

    def __init__ (self):
        SheetBase.__init__(self, SheetModelType.Worksheet)
        self.rows = {}

    def setCell (self, col, row, cell):
        if not self.rows.has_key(row):
            self.rows[row] = {}

        self.rows[row][col] = cell

    def createDOM (self):
        nd = node.Element('worksheet')
        nd.setAttr('version', self.version)
        rows = self.rows.keys()
        rows.sort()
        for row in rows:
            rowNode = nd.appendElement('row')
            rowNode.setAttr('id', row)
        return nd


class CellModelType:
    Label = 0
    Unknown = 999


class CellBase(object):
    def __init__ (self, modelType):
        self.modelType = modelType

class LabelCell(CellBase):
    def __init__ (self):
        CellBase.__init__(self, CellModelType.Label)

