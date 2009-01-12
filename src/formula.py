
import struct, sys
import globals

class InvalidCellAddress(Exception): pass

def toColName (colID):
    if colID > 255:
        globals.error("Column ID greater than 255")
        raise InvalidCellAddress
    n1 = colID % 26
    n2 = int(colID/26)
    name = struct.pack('b', n1 + ord('A'))
    if n2 > 0:
        name += struct.pack('b', n2 + ord('A'))
    return name

def toAbsName (name, isRelative):
    if not isRelative:
        name = '$' + name
    return name

class CellAddress(object):
    def __init__ (self, col=0, row=0, colRel=False, rowRel=False):
        self.col = col
        self.row = row
        self.isColRelative = colRel
        self.isRowRelative = rowRel

    def getName (self):
        colName = toAbsName(toColName(self.col), self.isColRelative)
        rowName = toAbsName("%d"%(self.row+1),   self.isRowRelative)
        return colName + rowName

class CellRange(object):
    def __init__ (self):
        self.firstRow = 0
        self.lastRow = 0
        self.firstCol = 0
        self.lastCol = 0
        self.isFirstRowRelative = False
        self.isLastRowRelative = False
        self.isFirstColRelative = False
        self.isLastColRelative = False

    def getName (self):
        col1 = toColName(self.firstCol)
        col2 = toColName(self.lastCol)
        row1 = "%d"%(self.firstRow+1)
        row2 = "%d"%(self.lastRow+1)
        col1 = toAbsName(col1, self.isFirstColRelative)
        col2 = toAbsName(col2, self.isLastColRelative)
        row1 = toAbsName(row1, self.isFirstRowRelative)
        row2 = toAbsName(row2, self.isLastRowRelative)
        return col1 + row1 + ':' + col2 + row2

def parseCellAddress (bytes):
    if len(bytes) != 4:
        globals.error("Byte size is %d but expected 4 bytes for cell address.\n"%len(bytes))
        raise InvalidCellAddress

    row = globals.getSignedInt(bytes[0:2])
    col = globals.getSignedInt(bytes[2:4])
    colRelative = ((col & 0x4000) != 0)
    rowRelative = ((col & 0x8000) != 0)
    col = (col & 0x00FF)
    obj = CellAddress(col, row, colRelative, rowRelative)
    return obj

def parseCellRangeAddress (bytes):
    if len(bytes) != 8:
        raise InvalidCellAddress

    obj = CellRange()
    obj.firstRow = globals.getSignedInt(bytes[0:2])
    obj.lastRow  = globals.getSignedInt(bytes[2:4])
    obj.firstCol = globals.getSignedInt(bytes[4:6])
    obj.lastCol  = globals.getSignedInt(bytes[6:8])

    obj.isFirstColRelative = ((obj.firstCol & 0x4000) != 0)
    obj.isFirstRowRelative = ((obj.firstCol & 0x8000) != 0)
    obj.firstCol = (obj.firstCol & 0x00FF)

    obj.isLastColRelative = ((obj.lastCol & 0x4000) != 0)
    obj.isLastRowRelative = ((obj.lastCol & 0x8000) != 0)
    obj.lastCol = (obj.lastCol & 0x00FF)
    return obj


def makeSheetName (sheet1, sheet2):
    if sheet2 == None or sheet1 == sheet2:
        sheetName = "sheetID='%d'"%sheet1
    else:
        sheetName = "sheetID='%d-%d'"%(sheet1, sheet2)
    return sheetName


class TokenBase(object):
    """base class for token handler

Derive a class from this base class to create a token handler for a formula
token.  

The parse method takes the token array position that points to the first 
token to be processed, and returns the position of the laste token that has 
been processed.  So, if the handler processes only one token, it should 
return the same value it receives without incrementing it.  

"""
    def __init__ (self, header, tokens):
        self.header = header
        self.tokens = tokens
        self.size = len(self.tokens)
        self.init()

    def init (self):
        """initializer for a derived class"""
        pass

    def parse (self, i):
        return i

    def getText (self):
        return ''

class Add(TokenBase): pass
class Sub(TokenBase): pass
class Mul(TokenBase): pass
class Div(TokenBase): pass
class Power(TokenBase): pass
class Concat(TokenBase): pass
class LT(TokenBase): pass
class LE(TokenBase): pass
class EQ(TokenBase): pass
class GE(TokenBase): pass
class GT(TokenBase): pass
class NE(TokenBase): pass
class Isect(TokenBase): pass
class List(TokenBase): pass
class Range(TokenBase): pass

class Plus(TokenBase): pass
class Minus(TokenBase): pass
class Percent(TokenBase): pass

class NameX(TokenBase):
    """external name"""

    def parse (self, i):
        i += 1
        self.refID = globals.getSignedInt(self.tokens[i:i+2])
        i += 2
        self.nameID = globals.getSignedInt(self.tokens[i:i+2])
        i += 2
        return i

    def getText (self):
        return "<externname externSheetID='%d' nameID='%d'>"%(self.refID, self.nameID)


class Ref3dR(TokenBase):
    """3D reference or external reference to a cell"""

    def init (self):
        self.cell = None
        self.sheet1 = None
        self.sheet2 = None

    def parse (self, i):
        try:
            i += 1
            self.sheet1 = globals.getSignedInt(self.tokens[i:i+2])
            i += 2
            if self.header == 0x0023:
                # 3A in EXTERNNAME expects a 2nd sheet index
                self.sheet2 = globals.getSignedInt(self.tokens[i:i+2])
                i += 2
            self.cell = parseCellAddress(self.tokens[i:i+4])
            i += 4
        except InvalidCellAddress:
            pass
        return i

    def getText (self):
        if self.cell == None:
            return ''
        cellName = self.cell.getName()
        sheetName = makeSheetName(self.sheet1, self.sheet2)
        return "<3dref %s cellAddress='%s'>"%(sheetName, cellName)


class Ref3dV(TokenBase):
    """3D reference or external reference to a cell"""

    def init (self):
        self.cell = None

    def parse (self, i):
        try:
            i += 1
            self.extSheetId = globals.getSignedInt(self.tokens[i:i+2])
            i += 2
            self.cell = parseCellAddress(self.tokens[i:i+4])
            i += 4
        except InvalidCellAddress:
            pass
        return i

    def getText (self):
        if self.cell == None:
            return ''
        cellName = self.cell.getName()
        return "<3dref externSheetID=%d cellAddress='%s'>"%(self.extSheetId, cellName)


class Ref3dA(Ref3dV):
    def __init__ (self, header, tokens):
        Ref3dV.__init__(self, header, tokens)


class Area3d(TokenBase):

    def parse (self, i):
        self.cellrange = None
        try:
            op = self.tokens[i]
            i += 1
            self.extSheetId = globals.getSignedInt(self.tokens[i:i+2])
            i += 2
            self.cellrange = parseCellRangeAddress(self.tokens[i:i+8])
        except InvalidCellAddress:
            pass
        return i

    def getText (self):
        if self.cellrange == None:
            return ''
        cellRangeName = self.cellrange.getName()
        return "<3drange externSheetID=%d rangeAddress='%s'>"%(self.extSheetId, cellRangeName)

class Error(TokenBase):

    def parse (self, i):
        i += 1 # skip opcode
        self.errorNum = globals.getSignedInt(self.tokens[i:i+1])
        i += 1
        return i

    def getText (self):
        errorText = ''
        if self.errorNum == 0x17:
            errorText = '#REF!'
        return "<error code='0x%2.2X' text='%s'>"%(self.errorNum, errorText)

tokenMap = {
    # binary operator
    0x03: Add,
    0x04: Sub,
    0x05: Mul,
    0x06: Div,
    0x07: Power,
    0x08: Concat,
    0x09: LT,
    0x0A: LE,
    0x0B: EQ,
    0x0C: GE,
    0x0D: GT,
    0x0E: NE,
    0x0F: Isect,
    0x10: List,
    0x11: Range,

    # unary operator
    0x12: Plus,
    0x13: Minus,
    0x14: Percent,

    # operand tokens
    0x39: NameX,
    0x59: NameX,
    0x79: NameX,

    # 3d reference (TODO: There is a slight variation in how a cell reference
    # is represented between 0x3A and 0x5A).
    0x3A: Ref3dR,
    0x5A: Ref3dV,
    0x7A: Ref3dA,

    0x3B: Area3d,
    0x5B: Area3d,
    0x7B: Area3d,

    0x1C: Error,

    # last item
  0xFFFF: None
}

class FormulaParser(object):
    """formula parser for token bytes

This class receives a series of bytes that represent formula tokens through
the constructor.  That series of bytes must also include the formula length
which is usually the first 2 bytes.
"""
    def __init__ (self, header, tokens, sizeField=True):
        self.header = header
        self.tokens = tokens
        self.text = ''
        self.sizeField = sizeField

    def parse (self):
        ftokens = self.tokens
        length = len(ftokens)
        if self.sizeField:
            # first 2-bytes contain the length of the formula tokens
            length = globals.getSignedInt(self.tokens[0:2])
            if length <= 0:
                return
            ftokens = self.tokens[2:2+length]
            length = len(ftokens)

        i = 0
        while i < length:
            tk = ftokens[i]

            if type(tk) == type('c'):
                # get the ordinal of the character.
                tk = ord(tk)

            if not tokenMap.has_key(tk):
                # no token handler
                i += 1
                continue

            # token handler exists.
            o = tokenMap[tk](self.header, ftokens)
            i = o.parse(i)
            self.text += o.getText() + ' '

            i += 1


    def getText (self):
        return self.text
