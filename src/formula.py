########################################################################
#
#  Copyright (c) 2010 Kohei Yoshida
#  
#  Permission is hereby granted, free of charge, to any person
#  obtaining a copy of this software and associated documentation
#  files (the "Software"), to deal in the Software without
#  restriction, including without limitation the rights to use,
#  copy, modify, merge, publish, distribute, sublicense, and/or sell
#  copies of the Software, and to permit persons to whom the
#  Software is furnished to do so, subject to the following
#  conditions:
#  
#  The above copyright notice and this permission notice shall be
#  included in all copies or substantial portions of the Software.
#  
#  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
#  EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
#  OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
#  NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
#  HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
#  WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
#  FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
#  OTHER DEALINGS IN THE SOFTWARE.
#
########################################################################

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
        self.strm = globals.ByteStream(tokens)
        self.text = ''
        self.sizeField = sizeField

    def parse (self):
        length = self.strm.getSize()
        if self.sizeField:
            # first 2-bytes contain the length of the formula tokens
            length = self.strm.readUnsignedInt(2)
            if length <= 0:
                return
            ftokens = self.strm.readBytes(length)
            length = len(ftokens)
        else:
            ftokens = self.strm.readRemainingBytes()

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

# ============================================================================

class TokenType:
    Area3d = 0
    Unknown = 9999

class _TokenBase(object):
    def __init__ (self, strm, opcode1, opcode2=None):
        self.opcode1 = opcode1
        self.opcode2 = opcode2
        self.strm = strm
        self.tokenType = TokenType.Unknown

    def parse (self):
        self.parseBytes()
        self.strm = None # no need to hold reference to the stream.

    def parseBytes (self):
        # derived class should overwrite this method.
        pass

    def getText (self):
        return ''

class _Int(_TokenBase):
    def parseBytes (self):
        self.value = self.strm.readUnsignedInt(2)

    def getText (self):
        return "%d"%self.value

class _Area3d(_TokenBase):
    def parseBytes (self):
        self.xti = self.strm.readUnsignedInt(2)
        self.cellRange = parseCellRangeAddress(self.strm.readBytes(8))
        self.tokenType = TokenType.Area3d

    def getText (self):
        return "(xti=%d,"%self.xti + self.cellRange.getName() + ")"

class _FuncVar(_TokenBase):

    funcTab = {
        0x0000: 'COUNT',
        0x0001: 'IF',
        0x0002: 'ISNA',
        0x0003: 'ISERROR',
        0x0004: 'SUM',
        0x0005: 'AVERAGE',
        0x0006: 'MIN',
        0x0007: 'MAX',
        0x0008: 'ROW',
        0x0009: 'COLUMN',
        0x000A: 'NA',
        0x000B: 'NPV',
        0x000C: 'STDEV',
        0x000D: 'DOLLAR',
        0x000E: 'FIXED',
        0x000F: 'SIN',
        0x0010: 'COS',
        0x0011: 'TAN',
        0x0012: 'ATAN',
        0x0013: 'PI',
        0x0014: 'SQRT',
        0x0015: 'EXP',
        0x0016: 'LN',
        0x0017: 'LOG10',
        0x0018: 'ABS',
        0x0019: 'INT',
        0x001A: 'SIGN',
        0x001B: 'ROUND',
        0x001C: 'LOOKUP',
        0x001D: 'INDEX',
        0x001E: 'REPT',
        0x001F: 'MID',
        0x0020: 'LEN',
        0x0021: 'VALUE',
        0x0022: 'TRUE',
        0x0023: 'FALSE',
        0x0024: 'AND',
        0x0025: 'OR',
        0x0026: 'NOT',
        0x0027: 'MOD',
        0x0028: 'DCOUNT',
        0x0029: 'DSUM',
        0x002A: 'DAVERAGE',
        0x002B: 'DMIN',
        0x002C: 'DMAX',
        0x002D: 'DSTDEV',
        0x002E: 'VAR',
        0x002F: 'DVAR',
        0x0030: 'TEXT',
        0x0031: 'LINEST',
        0x0032: 'TREND',
        0x0033: 'LOGEST',
        0x0034: 'GROWTH',
        0x0035: 'GOTO',
        0x0036: 'HALT',
        0x0037: 'RETURN',
        0x0038: 'PV',
        0x0039: 'FV',
        0x003A: 'NPER',
        0x003B: 'PMT',
        0x003C: 'RATE',
        0x003D: 'MIRR',
        0x003E: 'IRR',
        0x003F: 'RAND',
        0x0040: 'MATCH',
        0x0041: 'DATE',
        0x0042: 'TIME',
        0x0043: 'DAY',
        0x0044: 'MONTH',
        0x0045: 'YEAR',
        0x0046: 'WEEKDAY',
        0x0047: 'HOUR',
        0x0048: 'MINUTE',
        0x0049: 'SECOND',
        0x004A: 'NOW',
        0x004B: 'AREAS',
        0x004C: 'ROWS',
        0x004D: 'COLUMNS',
        0x004E: 'OFFSET',
        0x004F: 'ABSREF',
        0x0050: 'RELREF',
        0x0051: 'ARGUMENT',
        0x0052: 'SEARCH',
        0x0053: 'TRANSPOSE',
        0x0054: 'ERROR',
        0x0055: 'STEP',
        0x0056: 'TYPE',
        0x0057: 'ECHO',
        0x0058: 'SET.NAME',
        0x0059: 'CALLER',
        0x005A: 'DEREF',
        0x005B: 'WINDOWS',
        0x005C: 'SERIES',
        0x005D: 'DOCUMENTS',
        0x005E: 'ACTIVE.CELL',
        0x005F: 'SELECTION',
        0x0060: 'RESULT',
        0x0061: 'ATAN2',
        0x0062: 'ASIN',
        0x0063: 'ACOS',
        0x0064: 'CHOOSE',
        0x0065: 'HLOOKUP',
        0x0066: 'VLOOKUP',
        0x0067: 'LINKS',
        0x0068: 'INPUT',
        0x0069: 'ISREF',
        0x006A: 'GET.FORMULA',
        0x006B: 'GET.NAME',
        0x006C: 'SET.VALUE',
        0x006D: 'LOG',
        0x006E: 'EXEC',
        0x006F: 'CHAR',
        0x0070: 'LOWER',
        0x0071: 'UPPER',
        0x0072: 'PROPER',
        0x0073: 'LEFT',
        0x0074: 'RIGHT',
        0x0075: 'EXACT',
        0x0076: 'TRIM',
        0x0077: 'REPLACE',
        0x0078: 'SUBSTITUTE',
        0x0079: 'CODE',
        0x007A: 'NAMES',
        0x007B: 'DIRECTORY',
        0x007C: 'FIND',
        0x007D: 'CELL',
        0x007E: 'ISERR',
        0x007F: 'ISTEXT',
        0x0080: 'ISNUMBER',
        0x0081: 'ISBLANK',
        0x0082: 'T',
        0x0083: 'N',
        0x0084: 'FOPEN',
        0x0085: 'FCLOSE',
        0x0086: 'FSIZE',
        0x0087: 'FREADLN',
        0x0088: 'FREAD',
        0x0089: 'FWRITELN',
        0x008A: 'FWRITE',
        0x008B: 'FPOS',
        0x008C: 'DATEVALUE',
        0x008D: 'TIMEVALUE',
        0x008E: 'SLN',
        0x008F: 'SYD',
        0x0090: 'DDB',
        0x0091: 'GET.DEF',
        0x0092: 'REFTEXT',
        0x0093: 'TEXTREF',
        0x0094: 'INDIRECT',
        0x0095: 'REGISTER',
        0x0096: 'CALL',
        0x0097: 'ADD.BAR',
        0x0098: 'ADD.MENU',
        0x0099: 'ADD.COMMAND',
        0x009A: 'ENABLE.COMMAND',
        0x009B: 'CHECK.COMMAND',
        0x009C: 'RENAME.COMMAND',
        0x009D: 'SHOW.BAR',
        0x009E: 'DELETE.MENU',
        0x009F: 'DELETE.COMMAND',
        0x00A0: 'GET.CHART.ITEM',
        0x00A1: 'DIALOG.BOX',
        0x00A2: 'CLEAN',
        0x00A3: 'MDETERM',
        0x00A4: 'MINVERSE',
        0x00A5: 'MMULT',
        0x00A6: 'FILES',
        0x00A7: 'IPMT',
        0x00A8: 'PPMT',
        0x00A9: 'COUNTA',
        0x00AA: 'CANCEL.KEY',
        0x00AB: 'FOR',
        0x00AC: 'WHILE',
        0x00AD: 'BREAK',
        0x00AE: 'NEXT',
        0x00AF: 'INITIATE',
        0x00B0: 'REQUEST',
        0x00B1: 'POKE',
        0x00B2: 'EXECUTE',
        0x00B3: 'TERMINATE',
        0x00B4: 'RESTART',
        0x00B5: 'HELP',
        0x00B6: 'GET.BAR',
        0x00B7: 'PRODUCT',
        0x00B8: 'FACT',
        0x00B9: 'GET.CELL',
        0x00BA: 'GET.WORKSPACE',
        0x00BB: 'GET.WINDOW',
        0x00BC: 'GET.DOCUMENT',
        0x00BD: 'DPRODUCT',
        0x00BE: 'ISNONTEXT',
        0x00BF: 'GET.NOTE',
        0x00C0: 'NOTE',
        0x00C1: 'STDEVP',
        0x00C2: 'VARP',
        0x00C3: 'DSTDEVP',
        0x00C4: 'DVARP',
        0x00C5: 'TRUNC',
        0x00C6: 'ISLOGICAL',
        0x00C7: 'DCOUNTA',
        0x00C8: 'DELETE.BAR',
        0x00C9: 'UNREGISTER',
        0x00CC: 'USDOLLAR',
        0x00CD: 'FINDB',
        0x00CE: 'SEARCHB',
        0x00CF: 'REPLACEB',
        0x00D0: 'LEFTB',
        0x00D1: 'RIGHTB',
        0x00D2: 'MIDB',
        0x00D3: 'LENB',
        0x00D4: 'ROUNDUP',
        0x00D5: 'ROUNDDOWN',
        0x00D6: 'ASC',
        0x00D7: 'DBCS',
        0x00D8: 'RANK',
        0x00DB: 'ADDRESS',
        0x00DC: 'DAYS360',
        0x00DD: 'TODAY',
        0x00DE: 'VDB',
        0x00DF: 'ELSE',
        0x00E0: 'ELSE.IF',
        0x00E1: 'END.IF',
        0x00E2: 'FOR.CELL',
        0x00E3: 'MEDIAN',
        0x00E4: 'SUMPRODUCT',
        0x00E5: 'SINH',
        0x00E6: 'COSH',
        0x00E7: 'TANH',
        0x00E8: 'ASINH',
        0x00E9: 'ACOSH',
        0x00EA: 'ATANH',
        0x00EB: 'DGET',
        0x00EC: 'CREATE.OBJECT',
        0x00ED: 'VOLATILE',
        0x00EE: 'LAST.ERROR',
        0x00EF: 'CUSTOM.UNDO',
        0x00F0: 'CUSTOM.REPEAT',
        0x00F1: 'FORMULA.CONVERT',
        0x00F2: 'GET.LINK.INFO',
        0x00F3: 'TEXT.BOX',
        0x00F4: 'INFO',
        0x00F5: 'GROUP',
        0x00F6: 'GET.OBJECT',
        0x00F7: 'DB',
        0x00F8: 'PAUSE',
        0x00FB: 'RESUME',
        0x00FC: 'FREQUENCY',
        0x00FD: 'ADD.TOOLBAR',
        0x00FE: 'DELETE.TOOLBAR',
        0x00FF: 'User Defined Function',
        0x0100: 'RESET.TOOLBAR',
        0x0101: 'EVALUATE',
        0x0102: 'GET.TOOLBAR',
        0x0103: 'GET.TOOL',
        0x0104: 'SPELLING.CHECK',
        0x0105: 'ERROR.TYPE',
        0x0106: 'APP.TITLE',
        0x0107: 'WINDOW.TITLE',
        0x0108: 'SAVE.TOOLBAR',
        0x0109: 'ENABLE.TOOL',
        0x010A: 'PRESS.TOOL',
        0x010B: 'REGISTER.ID',
        0x010C: 'GET.WORKBOOK',
        0x010D: 'AVEDEV',
        0x010E: 'BETADIST',
        0x010F: 'GAMMALN',
        0x0110: 'BETAINV',
        0x0111: 'BINOMDIST',
        0x0112: 'CHIDIST',
        0x0113: 'CHIINV',
        0x0114: 'COMBIN',
        0x0115: 'CONFIDENCE',
        0x0116: 'CRITBINOM',
        0x0117: 'EVEN',
        0x0118: 'EXPONDIST',
        0x0119: 'FDIST',
        0x011A: 'FINV',
        0x011B: 'FISHER',
        0x011C: 'FISHERINV',
        0x011D: 'FLOOR',
        0x011E: 'GAMMADIST',
        0x011F: 'GAMMAINV',
        0x0120: 'CEILING',
        0x0121: 'HYPGEOMDIST',
        0x0122: 'LOGNORMDIST',
        0x0123: 'LOGINV',
        0x0124: 'NEGBINOMDIST',
        0x0125: 'NORMDIST',
        0x0126: 'NORMSDIST',
        0x0127: 'NORMINV',
        0x0128: 'NORMSINV',
        0x0129: 'STANDARDIZE',
        0x012A: 'ODD',
        0x012B: 'PERMUT',
        0x012C: 'POISSON',
        0x012D: 'TDIST',
        0x012E: 'WEIBULL',
        0x012F: 'SUMXMY2',
        0x0130: 'SUMX2MY2',
        0x0131: 'SUMX2PY2',
        0x0132: 'CHITEST',
        0x0133: 'CORREL',
        0x0134: 'COVAR',
        0x0135: 'FORECAST',
        0x0136: 'FTEST',
        0x0137: 'INTERCEPT',
        0x0138: 'PEARSON',
        0x0139: 'RSQ',
        0x013A: 'STEYX',
        0x013B: 'SLOPE',
        0x013C: 'TTEST',
        0x013D: 'PROB',
        0x013E: 'DEVSQ',
        0x013F: 'GEOMEAN',
        0x0140: 'HARMEAN',
        0x0141: 'SUMSQ',
        0x0142: 'KURT',
        0x0143: 'SKEW',
        0x0144: 'ZTEST',
        0x0145: 'LARGE',
        0x0146: 'SMALL',
        0x0147: 'QUARTILE',
        0x0148: 'PERCENTILE',
        0x0149: 'PERCENTRANK',
        0x014A: 'MODE',
        0x014B: 'TRIMMEAN',
        0x014C: 'TINV',
        0x014E: 'MOVIE.COMMAND',
        0x014F: 'GET.MOVIE',
        0x0150: 'CONCATENATE',
        0x0151: 'POWER',
        0x0152: 'PIVOT.ADD.DATA',
        0x0153: 'GET.PIVOT.TABLE',
        0x0154: 'GET.PIVOT.FIELD',
        0x0155: 'GET.PIVOT.ITEM',
        0x0156: 'RADIANS',
        0x0157: 'DEGREES',
        0x0158: 'SUBTOTAL',
        0x0159: 'SUMIF',
        0x015A: 'COUNTIF',
        0x015B: 'COUNTBLANK',
        0x015C: 'SCENARIO.GET',
        0x015D: 'OPTIONS.LISTS.GET',
        0x015E: 'ISPMT',
        0x015F: 'DATEDIF',
        0x0160: 'DATESTRING',
        0x0161: 'NUMBERSTRING',
        0x0162: 'ROMAN',
        0x0163: 'OPEN.DIALOG',
        0x0164: 'SAVE.DIALOG',
        0x0165: 'VIEW.GET',
        0x0166: 'GETPIVOTDATA',
        0x0167: 'HYPERLINK',
        0x0168: 'PHONETIC',
        0x0169: 'AVERAGEA',
        0x016A: 'MAXA',
        0x016B: 'MINA',
        0x016C: 'STDEVPA',
        0x016D: 'VARPA',
        0x016E: 'STDEVA',
        0x016F: 'VARA',
        0x0170: 'BAHTTEXT',
        0x0171: 'THAIDAYOFWEEK',
        0x0172: 'THAIDIGIT',
        0x0173: 'THAIMONTHOFYEAR',
        0x0174: 'THAINUMSOUND',
        0x0175: 'THAINUMSTRING',
        0x0176: 'THAISTRINGLENGTH',
        0x0177: 'ISTHAIDIGIT',
        0x0178: 'ROUNDBAHTDOWN',
        0x0179: 'ROUNDBAHTUP',
        0x017A: 'THAIYEAR',
        0x017B: 'RTD'
    }

    def parseBytes (self):
        self.dataType = (self.opcode1 & 0x60)/32  # 0x1 = reference, 0x2 = value, 0x3 = array
        self.argCount = self.strm.readUnsignedInt(1)
        tab = self.strm.readUnsignedInt(2)
        self.funcType = (tab & 0x7FFF)
        self.isCeTab = (tab & 0x8000) != 0

    def getText (self):
        if self.isCeTab:
            # I'll support this later.
            return ''

        if not _FuncVar.funcTab.has_key(self.funcType):
            # unknown function name
            return '#NAME!'

        if self.argCount > 0:
            # I'll support functions with arguments later.
            return ''

        return _FuncVar.funcTab[self.funcType] + "()"

_tokenMap = {
    0x1E: _Int,
    0x3B: _Area3d,
    0x5B: _Area3d,
    0x7B: _Area3d,

    0x42: _FuncVar
}

class FormulaParser2(object):
    """This is a new formula parser that will eventually replace the old one.

Once replaced, I'll change the name to FormulaParser and the names of the 
associated token classes will be without the leading underscore (_)."""


    def __init__ (self, header, bytes):
        self.header = header
        self.tokens = []
        self.strm = globals.ByteStream(bytes)

    def parse (self):
        while not self.strm.isEndOfRecord():
            b = self.strm.readUnsignedInt(1)
            if not _tokenMap.has_key(b):
                # Unknown token.  Stop parsing.
                return

            token = _tokenMap[b](self.strm, b)
            token.parse()
            self.tokens.append(token)

    def getText (self):
        s = ''
        for tk in self.tokens:
            s += tk.getText()
        return s

    def getTokens (self):
        return self.tokens
