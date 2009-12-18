
import struct
import globals, formula

# -------------------------------------------------------------------
# record handler classes

def getValueOrUnknown (list, idx):
    listType = type(list)
    if listType == type([]):
        # list
        if idx < len(list):
            return list[idx]
    elif listType == type({}):
        # dictionary
        if list.has_key(idx):
            return list[idx]

    return '(unknown)'

class BaseRecordHandler(globals.ByteStream):

    def __init__ (self, header, size, bytes, strmData):
        globals.ByteStream.__init__(self, bytes)
        self.header = header
        self.lines = []
        self.strmData = strmData

    def parseBytes (self):
        """Parse the original bytes and generate human readable output.

The derived class should only worry about overwriting this function.  The
bytes are given as self.bytes, and call self.appendLine([new line]) to
append a line to be displayed.
"""
        pass

    def output (self):
        self.parseBytes()
        print("%4.4Xh: %s"%(self.header, "-"*61))
        for line in self.lines:
            print("%4.4Xh: %s"%(self.header, line))

    def appendLine (self, line):
        self.lines.append(line)

    def appendMultiLine (self, line):
        charWidth = 61
        singleLine = ''
        testLine = ''
        for word in line.split():
            testLine += word + ' '
            if len(testLine) > charWidth:
                self.lines.append(singleLine)
                testLine = word + ' '
            singleLine = testLine

        if len(singleLine) > 0:
            self.lines.append(singleLine)

    def appendLineBoolean (self, name, value):
        text = "%s: %s"%(name, self.getYesNo(value))
        self.appendLine(text)

    def appendCellPosition (self, col, row):
        text = "cell position: (col: %d; row: %d)"%(col, row)
        self.appendLine(text)

    def getYesNo (self, boolVal):
        if boolVal:
            return 'yes'
        else:
            return 'no'

    def getTrueFalse (self, boolVal):
        if boolVal:
            return 'true'
        else:
            return 'false'

    def getEnabledDisabled (self, boolVal):
        if boolVal:
            return 'enabled'
        else:
            return 'disabled'


class BOF(BaseRecordHandler):

    Type = {
        0x0005: "Workbook globals",
        0x0006: "Visual Basic module",
        0x0010: "Worksheet or dialog sheet",
        0x0020: "Chart",
        0x0040: "Excel 4.0 macro sheet",
        0x0100: "Workspace file"
    }

    # TODO: Add more build identifiers.
    buildId = {
        0x0DBB: 'Excel 97',
        0x0EDE: 'Excel 97',
        0x2775: 'Excel XP'
    }

    def getBuildIdName (self, value):
        if BOF.buildId.has_key(value):
            return BOF.buildId[value]
        else:
            return '(unknown)'

    def parseBytes (self):
        # BIFF version
        ver = self.readUnsignedInt(2)
        s = 'not BIFF8'
        if ver == 0x0600:
            s = 'BIFF8'
        self.appendLine("BIFF version: %s"%s)

        # Substream type
        dataType = self.readUnsignedInt(2)
        self.appendLine("type: %s"%BOF.Type[dataType])

        # build ID and year
        buildID = self.readUnsignedInt(2)
        self.appendLine("build ID: %s (%4.4Xh)"%(self.getBuildIdName(buildID), buildID))
        buildYear = self.readUnsignedInt(2)
        self.appendLine("build year: %d"%buildYear)

        # file history flags
        flags = self.readUnsignedInt(4)
        win     = (flags & 0x00000001)
        risc    = (flags & 0x00000002)
        beta    = (flags & 0x00000004)
        winAny  = (flags & 0x00000008)
        macAny  = (flags & 0x00000010)
        betaAny = (flags & 0x00000020)
        riscAny = (flags & 0x00000100)
        self.appendLine("last edited by Excel on Windows: %s"%self.getYesNo(win))
        self.appendLine("last edited by Excel on RISC: %s"%self.getYesNo(risc))
        self.appendLine("last edited by beta version of Excel: %s"%self.getYesNo(beta))
        self.appendLine("has ever been edited by Excel for Windows: %s"%self.getYesNo(winAny))
        self.appendLine("has ever been edited by Excel for Macintosh: %s"%self.getYesNo(macAny))
        self.appendLine("has ever been edited by beta version of Excel: %s"%self.getYesNo(betaAny))
        self.appendLine("has ever been edited by Excel on RISC: %s"%self.getYesNo(riscAny))

        lowestExcelVer = self.readSignedInt(4)
        self.appendLine("earliest Excel version that can read all records: %d"%lowestExcelVer)


class BoundSheet(BaseRecordHandler):

    hiddenStates = {0x00: 'visible', 0x01: 'hidden', 0x02: 'very hidden'}

    sheetTypes = {0x00: 'worksheet or dialog sheet',
                  0x01: 'Excel 4.0 macro sheet',
                  0x02: 'chart',
                  0x06: 'Visual Basic module'}

    @staticmethod
    def getHiddenState (flag):
        if BoundSheet.hiddenStates.has_key(flag):
            return BoundSheet.hiddenStates[flag]
        else:
            return 'unknown'

    @staticmethod
    def getSheetType (flag):
        if BoundSheet.sheetTypes.has_key(flag):
            return BoundSheet.sheetTypes[flag]
        else:
            return 'unknown'

    def parseBytes (self):
        posBOF = self.readUnsignedInt(4)
        flags = self.readUnsignedInt(2)
        textLen = self.readUnsignedInt(1)
        text, textLen = globals.getRichText(self.readRemainingBytes(), textLen)
        self.appendLine("BOF position in this stream: %d"%posBOF)
        self.appendLine("sheet name: %s"%text)

        hiddenState = (flags & 0x0003)
        self.appendLine("hidden state: %s"%BoundSheet.getHiddenState(hiddenState))

        sheetType = (flags & 0xFF00)
        self.appendLine("sheet type: %s"%BoundSheet.getSheetType(sheetType))


class Dimensions(BaseRecordHandler):

    def parseBytes (self):
        rowMic = self.readUnsignedInt(4)
        rowMac = self.readUnsignedInt(4)
        colMic = self.readUnsignedInt(2)
        colMac = self.readUnsignedInt(2)

        self.appendLine("first defined row: %d"%rowMic)
        self.appendLine("last defined row plus 1: %d"%rowMac)
        self.appendLine("first defined column: %d"%colMic)
        self.appendLine("last defined column plus 1: %d"%colMac)


class FilePass(BaseRecordHandler):

    def parseBytes (self):
        mode = self.readUnsignedInt(2)    # mode: 0 = BIFF5  1 = BIFF8
        self.readUnsignedInt(2)           # ignore 2 bytes.
        subMode = self.readUnsignedInt(2) # submode: 1 = standard encryption  2 = strong encryption

        modeName = 'unknown'
        if mode == 0:
            modeName = 'BIFF5'
        elif mode == 1:
            modeName = 'BIFF8'

        encType = 'unknown'
        if subMode == 1:
            encType = 'standard'
        elif subMode == 2:
            encType = 'strong'
        
        self.appendLine("mode: %s"%modeName)
        self.appendLine("encryption type: %s"%encType)
        self.appendLine("")
        self.appendMultiLine("NOTE: Since this document appears to be encrypted, the dumper will not parse the record contents from this point forward.")


class Formula(BaseRecordHandler):

    def parseBytes (self):
        row  = globals.getSignedInt(self.bytes[0:2])
        col  = globals.getSignedInt(self.bytes[2:4])
        xf   = globals.getSignedInt(self.bytes[4:6])
        fval = globals.getDouble(self.bytes[6:14])

        flags          = globals.getSignedInt(self.bytes[14:16])
        recalc         = (flags & 0x0001) != 0
        calcOnOpen     = (flags & 0x0002) != 0
        sharedFormula  = (flags & 0x0008) != 0

        tokens = self.bytes[20:]
        fparser = formula.FormulaParser(self.header, tokens)
        fparser.parse()
        ftext = fparser.getText()

        self.appendCellPosition(col, row)
        self.appendLine("XF record ID: %d"%xf)
        self.appendLine("formula result: %g"%fval)
        self.appendLine("recalculate always: %d"%recalc)
        self.appendLine("calculate on open: %d"%calcOnOpen)
        self.appendLine("shared formula: %d"%sharedFormula)
        self.appendLine("formula bytes: %s"%globals.getRawBytes(tokens, True, False))
        self.appendLine("tokens: "+ftext)


class Array(BaseRecordHandler):

    def parseBytes (self):
        row1 = self.readUnsignedInt(2)
        row2 = self.readUnsignedInt(2)
        col1 = self.readUnsignedInt(1)
        col2 = self.readUnsignedInt(1)

        flags = self.readUnsignedInt(2)
        alwaysCalc = (0x01 & flags)
        calcOnLoad = (0x02 & flags)

        # Ignore these bits when reading a BIFF file.  When a BIFF file is 
        # being written, chn must be 00000000.
        chn = self.readUnsignedInt(4)

        fmlLen = self.readUnsignedInt(2)
        tokens = self.readBytes(fmlLen)

        fparser = formula.FormulaParser(self.header, tokens)
        fparser.parse()
        ftext = fparser.getText()

        self.appendLine("rows: %d - %d"%(row1, row2))
        self.appendLine("columns: %d - %d"%(col1, col2))
        self.appendLine("always calculate formula: " + self.getTrueFalse(alwaysCalc))
        self.appendLine("calculate formula on file load: " + self.getTrueFalse(calcOnLoad))
        self.appendLine("formula bytes: %s"%globals.getRawBytes(tokens, True, False))
        self.appendLine("tokens: " + ftext)

        if self.getCurrentPos() >= len(self.bytes):
            return

        # cached values
        cols = self.readUnsignedInt(1) + 1
        rows = self.readUnsignedInt(2) + 1
        self.appendLine("array size: cols=%d, rows=%d"%(cols, rows))
        for row in xrange(0, rows):
            for col in xrange(0, cols):
                msg = "(row=%d, col=%d): "%(row, col)
                valtype = self.readUnsignedInt(1)
                if valtype == 0x00:
                    # empty - ignore 8 bytes.
                    self.readUnsignedInt(8)
                    msg += "empty"
                elif valtype == 0x01:
                    # double
                    val = self.readDouble()
                    msg += "double: %g"%val
                elif valtype == 0x02:
                    # string
                    strLen = self.readUnsignedInt(2) + 1
                    text, strLen = globals.getRichText(self.readBytes(strLen), strLen)
                    msg += "text: '%s'"%text
                elif valtype == 0x04:
                    # bool
                    val = self.readUnsignedInt(1)
                    msg += "bool: " + self.getTrueFalse(val)
                elif valtype == 0x10:
                    # error
                    val = self.readUnsignedInt(1)
                    msg += "error: %d"%val
                self.appendLine(msg)

        return


class Label(BaseRecordHandler):

    def parseBytes (self):
        col = self.readUnsignedInt(2)
        row = self.readUnsignedInt(2)
        xfIdx = self.readUnsignedInt(2)
        textLen = self.readUnsignedInt(2)
        text, textLen = globals.getRichText(self.readRemainingBytes(), textLen)
        self.appendCellPosition(col, row)
        self.appendLine("XF record ID: %d"%xfIdx)
        self.appendLine("label text: %s"%text)


class LabelSST(BaseRecordHandler):

    def parseBytes (self):
        col = self.readUnsignedInt(2)
        row = self.readUnsignedInt(2)
        xfIdx = self.readUnsignedInt(2)
        strId = self.readUnsignedInt(4)
        self.appendCellPosition(col, row)
        self.appendLine("XF record ID: %d"%xfIdx)
        self.appendLine("string ID in SST: %d"%strId)


class Number(BaseRecordHandler):

    def parseBytes (self):
        row = globals.getSignedInt(self.bytes[0:2])
        col = globals.getSignedInt(self.bytes[2:4])
        xf  = globals.getSignedInt(self.bytes[4:6])
        fval = globals.getDouble(self.bytes[6:14])
        self.appendCellPosition(col, row)
        self.appendLine("XF record ID: %d"%xf)
        self.appendLine("value: %g"%fval)


class Obj(BaseRecordHandler):

    ftEnd      = 0x00 # End of OBJ record
                      # 0x01 - 0x03 (reserved)
    ftMacro    = 0x04 # Fmla-style macro
    ftButton   = 0x05 # Command button
    ftGmo      = 0x06 # Group marker
    ftCf       = 0x07 # Clipboard format
    ftPioGrbit = 0x08 # Picture option flags
    ftPictFmla = 0x09 # Picture fmla-style macro
    ftCbls     = 0x0A # Check box link
    ftRbo      = 0x0B # Radio button
    ftSbs      = 0x0C # Scroll bar
    ftNts      = 0x0D # Note structure
    ftSbsFmla  = 0x0E # Scroll bar fmla-style macro
    ftGboData  = 0x0F # Group box data
    ftEdoData  = 0x10 # Edit control data
    ftRboData  = 0x11 # Radio button data
    ftCblsData = 0x12 # Check box data
    ftLbsData  = 0x13 # List box data
    ftCblsFmla = 0x14 # Check box link fmla-style macro
    ftCmo      = 0x15 # Common object data

    class Cmo:
        Types = [
            'Group',                   # 0x00
            'Line',                    # 0x01
            'Rectangle',               # 0x02
            'Oval',                    # 0x03
            'Arc',                     # 0x04
            'Chart',                   # 0x05
            'Text',                    # 0x06
            'Button',                  # 0x07
            'Picture',                 # 0x08
            'Polygon',                 # 0x09
            '(Reserved)',              # 0x0A
            'Check box',               # 0x0B
            'Option button',           # 0x0C
            'Edit box',                # 0x0D
            'Label',                   # 0x0E
            'Dialog box',              # 0x0F
            'Spinner',                 # 0x10
            'Scroll bar',              # 0x11
            'List box',                # 0x12
            'Group box',               # 0x13
            'Combo box',               # 0x14
            '(Reserved)',              # 0x15
            '(Reserved)',              # 0x16
            '(Reserved)',              # 0x17
            '(Reserved)',              # 0x18
            'Comment',                 # 0x19
            '(Reserved)',              # 0x1A
            '(Reserved)',              # 0x1B
            '(Reserved)',              # 0x1C
            '(Reserved)',              # 0x1D
            'Microsoft Office drawing' # 0x1E
        ]

        @staticmethod
        def getType (typeID):
            if len(Obj.Cmo.Types) > typeID:
                return Obj.Cmo.Types[typeID]
            return "(unknown) (0x%2.2X)"%typeID

    def parseBytes (self):
        while not self.isEndOfRecord():
            fieldType = self.readUnsignedInt(2)
            fieldSize = self.readUnsignedInt(2)
            if fieldType == Obj.ftEnd:
                # reached the end of OBJ record.
                return

            if fieldType == Obj.ftCmo:
                self.parseCmo(fieldSize)
            else:
                fieldBytes = self.readBytes(fieldSize)
                self.appendLine("field 0x%2.2X: %s"%(fieldType, globals.getRawBytes(fieldBytes, True, False)))

    def parseCmo (self, size):
        if size != 18:
            # size of Cmo must be 18.  Something is wrong here.
            self.readBytes(size)
            globals.error("parsing of common object field in OBJ failed due to invalid size.")
            return

        objType = self.readUnsignedInt(2)
        objID  = self.readUnsignedInt(2)
        flag   = self.readUnsignedInt(2)

        # the rest of the bytes are reserved & should be all zero.
        self.readBytes(12)

        self.appendLine("common object: ")
        self.appendLine("  type: %s"%Obj.Cmo.getType(objType))
        self.appendLine("  object ID: %d"%objID)

        # 0    0001h fLocked    =1 if the object is locked when the sheet is protected
        # 3-1  000Eh (Reserved) Reserved; must be 0 (zero)
        # 4    0010h fPrint     =1 if the object is printable
        # 12-5 1FE0h (Reserved) Reserved; must be 0 (zero)
        # 13   2000h fAutoFill  =1 if the object uses automatic fill style
        # 14   4000h fAutoLine  =1 if the object uses automatic line style
        # 15   8000h (Reserved) Reserved; must be 0 (zero)

        locked    = (flag & 0x0001)
        printable = (flag & 0x0010)
        autoFill  = (flag & 0x2000)
        autoLine  = (flag & 0x4000)
        self.appendLineBoolean("  locked", locked)
        self.appendLineBoolean("  printable", printable)
        self.appendLineBoolean("  automatic fill style", autoFill)
        self.appendLineBoolean("  automatic line style", autoLine)


class RK(BaseRecordHandler):
    """Cell with encoded integer or floating-point value"""

    def parseBytes (self):
        row = globals.getSignedInt(self.bytes[0:2])
        col = globals.getSignedInt(self.bytes[2:4])
        xf  = globals.getSignedInt(self.bytes[4:6])

        rkval = globals.getSignedInt(self.bytes[6:10])
        multi100  = ((rkval & 0x00000001) != 0)
        signedInt = ((rkval & 0x00000002) != 0)
        realVal   = (rkval & 0xFFFFFFFC)

        if signedInt:
            # for integer, perform right-shift by 2 bits.
            realVal = realVal/4
        else:
            # for floating-point, convert the value back to the bytes,
            # pad the bytes to make it 8-byte long, and convert it back
            # to the numeric value.
            tmpBytes = struct.pack('<L', realVal)
            tmpBytes = struct.pack('xxxx') + tmpBytes
            realVal = struct.unpack('<d', tmpBytes)[0]

        if multi100:
            realVal /= 100

        self.appendCellPosition(col, row)
        self.appendLine("XF record ID: %d"%xf)
        self.appendLine("multiplied by 100: %d"%multi100)
        if signedInt:
            self.appendLine("type: signed integer")
        else:
            self.appendLine("type: floating point")
        self.appendLine("value: %g"%realVal)


class String(BaseRecordHandler):
    """Cached string formula result for preceding formula record."""

    def parseBytes (self):
        strLen = globals.getSignedInt(self.bytes[0:1])
        name, byteLen = globals.getRichText(self.bytes[2:], strLen)
        self.appendLine("string value: '%s'"%name)


class SST(BaseRecordHandler):

    def parseBytes (self):
        refCount = self.readSignedInt(4) # total number of references in workbook
        strCount = self.readSignedInt(4) # total number of unique strings.
        self.appendLine("total number of references: %d"%refCount)
        self.appendLine("total number of unique strings: %d"%strCount)
        for i in xrange(0, strCount):
            extText, bytesRead = globals.getUnicodeRichExtText(self.bytes[self.getCurrentPos():])
            self.readBytes(bytesRead) # advance current position.
        return


class Blank(BaseRecordHandler):

    def parseBytes (self):
        row = globals.getSignedInt(self.bytes[0:2])
        col = globals.getSignedInt(self.bytes[2:4])
        xf  = globals.getSignedInt(self.bytes[4:6])
        self.appendCellPosition(col, row)
        self.appendLine("XF record ID: %d"%xf)


class DBCell(BaseRecordHandler):

    def parseBytes (self):
        rowRecOffset = self.readUnsignedInt(4)
        self.appendLine("offset to first ROW record: %d"%rowRecOffset)
        while not self.isEndOfRecord():
            cellOffset = self.readUnsignedInt(2)
            self.appendLine("offset to CELL record: %d"%cellOffset)
        return


class DefColWidth(BaseRecordHandler):

    def parseBytes (self):
        w = self.readUnsignedInt(2)
        self.appendLine("default column width (in characters): %d"%w)


class ColInfo(BaseRecordHandler):

    def parseBytes (self):
        colFirst = self.readUnsignedInt(2)
        colLast  = self.readUnsignedInt(2)
        coldx    = self.readUnsignedInt(2)
        ixfe     = self.readUnsignedInt(2)
        flags    = self.readUnsignedInt(2)

        isHidden = (flags & 0x0001)
        outlineLevel = (flags & 0x0700)/4
        isCollapsed = (flags & 0x1000)/4

        self.appendLine("formatted columns: %d - %d"%(colFirst,colLast))
        self.appendLine("column width (in 1/256s of a char): %d"%coldx)
        self.appendLine("XF record index: %d"%ixfe)
        self.appendLine("hidden: %s"%self.getYesNo(isHidden))
        self.appendLine("outline level: %d"%outlineLevel)
        self.appendLine("collapsed: %s"%self.getYesNo(isCollapsed))


class Row(BaseRecordHandler):

    def parseBytes (self):
        row  = self.readUnsignedInt(2)
        col1 = self.readUnsignedInt(2)
        col2 = self.readUnsignedInt(2)

        rowHeightBits = self.readUnsignedInt(2)
        rowHeight     = (rowHeightBits & 0x7FFF)
        defaultHeight = ((rowHeightBits & 0x8000) == 1)

        self.appendLine("row: %d; col: %d - %d"%(row, col1, col2))
        self.appendLine("row height (twips): %d"%rowHeight)
        if defaultHeight:
            self.appendLine("row height type: default")
        else:
            self.appendLine("row height type: custom")

        irwMac = self.readUnsignedInt(2)
        self.appendLine("optimize flag (0 for BIFF): %d"%irwMac)

        dummy = self.readUnsignedInt(2)
        flags = self.readUnsignedInt(2)
        outLevel   = (flags & 0x0007)
        collapsed  = (flags & 0x0010)
        zeroHeight = (flags & 0x0020)
        unsynced   = (flags & 0x0040)
        ghostDirty = (flags & 0x0080)
        self.appendLine("outline level: %d"%outLevel)
        self.appendLine("collapsed: %s"%self.getYesNo(collapsed))
        self.appendLine("zero height: %s"%self.getYesNo(zeroHeight))
        self.appendLine("unsynced: %s"%self.getYesNo(unsynced))
        self.appendLine("ghost dirty: %s"%self.getYesNo(ghostDirty))


class Name(BaseRecordHandler):
    """internal defined name"""

    def __getInt (self, offset, size):
        return globals.getSignedInt(self.bytes[offset:offset+size])

    def __parseOptionFlags (self, flags):
        self.appendLine("option flags:")
        isHidden       = (flags & 0x0001) != 0
        isFuncMacro    = (flags & 0x0002) != 0
        isVBMacro      = (flags & 0x0004) != 0
        isMacroName    = (flags & 0x0008) != 0
        isComplFormula = (flags & 0x0010) != 0
        isBuiltinName  = (flags & 0x0020) != 0
        funcGrp        = (flags & 0x0FC0) / 64
        isBinary       = (flags & 0x1000) != 0

        if isHidden:
            self.appendLine("  hidden")
        else:
            self.appendLine("  visible")

        if isMacroName:
            self.appendLine("  macro name")
            if isFuncMacro:
                self.appendLine("  function macro")
                self.appendLine("  function group: %d"%funcGrp)
            else:
                self.appendLine("  command macro")
            if isVBMacro:
                self.appendLine("  visual basic macro")
            else:
                self.appendLine("  sheet macro")
        else:
            self.appendLine("  standard name")

        if isComplFormula:
            self.appendLine("  complex formula")
        else:
            self.appendLine("  simple formula")
        if isBuiltinName:
            self.appendLine("  built-in name")
        else:
            self.appendLine("  user-defined name")
        if isBinary:
            self.appendLine("  binary data")
        else:
            self.appendLine("  formula definition")


    def parseBytes (self):
        optionFlags = self.__getInt(0, 2)

        keyShortCut = self.__getInt(2, 1)
        nameLen     = self.__getInt(3, 1)
        formulaLen  = self.__getInt(4, 2)
        sheetId     = self.__getInt(8, 2)

        # these optional texts may come after the formula token bytes.
        menuTextLen = self.__getInt(10, 1)
        descTextLen = self.__getInt(11, 1)
        helpTextLen = self.__getInt(12, 1)
        statTextLen = self.__getInt(13, 1)

        name, byteLen = globals.getRichText(self.bytes[14:], nameLen)
        name = globals.decodeName(name)
        tokenPos = 14 + byteLen
        tokenText = globals.getRawBytes(self.bytes[tokenPos:tokenPos+formulaLen], True, False)
        o = formula.FormulaParser(self.header, self.bytes[tokenPos:tokenPos+formulaLen], False)
        o.parse()
        self.appendLine("name: %s"%name)
        self.__parseOptionFlags(optionFlags)

        self.appendLine("sheet ID: %d"%sheetId)
        self.appendLine("menu text length: %d"%menuTextLen)
        self.appendLine("description length: %d"%descTextLen)
        self.appendLine("help tip text length: %d"%helpTextLen)
        self.appendLine("status bar text length: %d"%statTextLen)
        self.appendLine("formula length: %d"%formulaLen)
        self.appendLine("formula bytes: " + tokenText)
        self.appendLine("formula: " + o.getText())



class SupBook(BaseRecordHandler):
    """Supporting workbook"""

    def __parseSpecial (self):
        if self.bytes[2:4] == [0x01, 0x04]:
            # internal reference
            num = globals.getSignedInt(self.bytes[0:2])
            self.appendLine("sheet name count: %d (internal reference)"%num)
        elif self.bytes[0:4] == [0x00, 0x01, 0x01, 0x3A]:
            # add-in function
            self.appendLine("add-in function name stored in the following EXERNNAME record.")

    def __parseDDE (self):
        # not implemented yet.
        pass

    def parseBytes (self):
        if self.size == 4:
            self.__parseSpecial()
            return

        if self.bytes[0:2] == [0x00, 0x00]:
            self.__parseDDE()
            return

        num = globals.getSignedInt(self.bytes[0:2])
        self.appendLine("sheet name count: %d"%num)
        i = 2
        isFirst = True
        while i < self.size:
            nameLen = globals.getSignedInt(self.bytes[i:i+2])
            i += 2
            flags = globals.getSignedInt(self.bytes[i:i+1])
            i += 1
            name = globals.getTextBytes(self.bytes[i:i+nameLen])
            name = globals.decodeName(name)
            i += nameLen
            if isFirst:
                isFirst = False
                self.appendLine("document URL: %s"%name)
            else:
                self.appendLine("sheet name: %s"%name)


class ExternSheet(BaseRecordHandler):

    def parseBytes (self):
        num = globals.getSignedInt(self.bytes[0:2])
        for i in xrange(0, num):
            offset = 2 + i*6
            book = globals.getSignedInt(self.bytes[offset:offset+2])
            firstSheet = globals.getSignedInt(self.bytes[offset+2:offset+4])
            lastSheet  = globals.getSignedInt(self.bytes[offset+4:offset+6])
            self.appendLine("SUPBOOK record ID: %d  (sheet ID range: %d - %d)"%(book, firstSheet, lastSheet))


class ExternName(BaseRecordHandler):

    def __parseOptionFlags (self, flags):
        self.isBuiltinName = (flags & 0x0001) != 0
        self.isAutoDDE     = (flags & 0x0002) != 0
        self.isStdDocName  = (flags & 0x0008) != 0
        self.isOLE         = (flags & 0x0010) != 0
        # 5 - 14 bits stores last successful clip format
        self.lastClipFormat = (flags & 0x7FE0)

    def parseBytes (self):
        # first 2 bytes are option flags for external name.
        optionFlags = globals.getSignedInt(self.bytes[0:2])
        self.__parseOptionFlags(optionFlags)

        if self.isOLE:
            # next 4 bytes are 32-bit storage ID
            storageID = globals.getSignedInt(self.bytes[2:6])
            nameLen = globals.getSignedInt(self.bytes[6:7])
            name, byteLen = globals.getRichText(self.bytes[7:], nameLen)
            self.appendLine("type: OLE")
            self.appendLine("storage ID: %d"%storageID)
            self.appendLine("name: %s"%name)
        else:
            # assume external defined name (could be DDE link).
            # TODO: differentiate DDE link from external defined name.

            supbookID = globals.getSignedInt(self.bytes[2:4])
            nameLen = globals.getSignedInt(self.bytes[6:7])
            name, byteLen = globals.getRichText(self.bytes[7:], nameLen)
            tokenText = globals.getRawBytes(self.bytes[7+byteLen:], True, False)
            self.appendLine("type: defined name")
            if supbookID == 0:
                self.appendLine("sheet ID: 0 (global defined names)")
            else:
                self.appendLine("sheet ID: %d"%supbookID)
            self.appendLine("name: %s"%name)
            self.appendLine("formula bytes: %s"%tokenText)

            # parse formula tokens
            o = formula.FormulaParser(self.header, self.bytes[7+byteLen:])
            o.parse()
            ftext = o.getText()
            self.appendLine("formula: %s"%ftext)


class Xct(BaseRecordHandler):

    def parseBytes (self):
        crnCount = globals.getSignedInt(self.bytes[0:2])
        sheetIndex = globals.getSignedInt(self.bytes[2:4])
        self.appendLine("CRN count: %d"%crnCount)
        self.appendLine("index of referenced sheet in the SUPBOOK record: %d"%sheetIndex)


class Crn(BaseRecordHandler):

    def parseBytes (self):
        lastCol = globals.getSignedInt(self.bytes[0:1])
        firstCol = globals.getSignedInt(self.bytes[1:2])
        rowIndex = globals.getSignedInt(self.bytes[2:4])
        self.appendLine("first column: %d"%firstCol)
        self.appendLine("last column:  %d"%lastCol)
        self.appendLine("row index: %d"%rowIndex)

        i = 4
        n = len(self.bytes)
        while i < n:
            typeId = self.bytes[i]
            i += 1
            if typeId == 0x00:
                # empty value
                i += 8
                self.appendLine("* empty value")
            elif typeId == 0x01:
                # number
                val = globals.getDouble(self.bytes[i:i+8])
                i += 8
                self.appendLine("* numeric value (%g)"%val)
            elif typeId == 0x2:
                # string
                text, length = globals.getRichText(self.bytes[i:])
                i += length
                text = globals.decodeName(text)
                self.appendLine("* string value (%s)"%text)
            elif typeId == 0x04:
                # boolean
                val = self.bytes[i]
                i += 7 # next 7 bytes not used
                self.appendLine("* boolean value (%d)"%val)
            elif typeId == 0x10:
                # error value
                val = self.bytes[i]
                i += 7 # not used
                self.appendLine("* error value (%d)"%val)
            else:
                sys.stderr.write("error parsing CRN record")
                sys.exit(1)
            

class RefreshAll(BaseRecordHandler):

    def parseBytes (self):
        boolVal = globals.getSignedInt(self.bytes[0:2])
        strVal = "no"
        if boolVal:
            strVal = "yes"
        self.appendLine("refresh all external data ranges and pivot tables: %s"%strVal)


class Hyperlink(BaseRecordHandler):

    def parseBytes (self):
        rowFirst = self.readUnsignedInt(2)
        rowLast = self.readUnsignedInt(2)
        colFirst = self.readUnsignedInt(2)
        colLast = self.readUnsignedInt(2)
        # Rest of the stream stores undocumented hyperlink stream.  Refer to 
        # page 128 of MS Excel binary format spec.
        self.appendLine("rows: %d - %d"%(rowFirst, rowLast))
        self.appendLine("columns: %d - %d"%(colFirst, colLast))
        msg  = "NOTE: The stream after the first 8 bytes stores undocumented hyperlink stream.  "
        msg += "Refer to page 128 of the MS Excel binary format spec."
        self.appendLine('')
        self.appendMultiLine(msg)


class PhoneticInfo(BaseRecordHandler):

    phoneticType = [
        'narrow Katakana', # 0x00
        'wide Katakana',   # 0x01
        'Hiragana',        # 0x02
        'any type'         # 0x03
    ]

    @staticmethod
    def getPhoneticType (flag):
        return getValueOrUnknown(PhoneticInfo.phoneticType, flag)

    alignType = [
        'general alignment',    # 0x00
        'left aligned',         # 0x01
        'center aligned',       # 0x02
        'distributed alignment' # 0x03
    ]

    @staticmethod
    def getAlignType (flag):
        return getValueOrUnknown(PhoneticInfo.alignType, flag)

    def parseBytes (self):
        fontIdx = self.readUnsignedInt(2)
        self.appendLine("font ID: %d"%fontIdx)
        flags = self.readUnsignedInt(1)

        # flags: 0 0 0 0 0 0 0 0
        #       | unused| B | A |

        phType    = (flags)   & 0x03
        alignType = (flags/4) & 0x03

        self.appendLine("phonetic type: %s"%PhoneticInfo.getPhoneticType(phType))
        self.appendLine("alignment: %s"%PhoneticInfo.getAlignType(alignType))

        self.readUnsignedInt(1) # unused byte
        
        # TODO: read cell ranges.

        return


class Font(BaseRecordHandler):

    fontFamilyNames = [
        'not applicable', # 0x00
        'roman',          # 0x01
        'swiss',          # 0x02
        'modern',         # 0x03
        'script',         # 0x04
        'decorative'      # 0x05
    ]

    @staticmethod
    def getFontFamily (code):
        return getValueOrUnknown(Font.fontFamilyNames, code)

    scriptNames = [
        'normal script',
        'superscript',
        'subscript'
    ]

    @staticmethod
    def getScriptName (code):
        return getValueOrUnknown(Font.scriptNames, code)


    underlineTypes = {
        0x00: 'no underline',
        0x01: 'single underline',
        0x02: 'double underline',
        0x21: 'single accounting',
        0x22: 'double accounting'
    }

    @staticmethod
    def getUnderlineStyleName (val):
        return getValueOrUnknown(Font.underlineTypes, val)

    charSetNames = {
        0x00: 'ANSI_CHARSET',
        0x01: 'DEFAULT_CHARSET',
        0x02: 'SYMBOL_CHARSET',
        0x4D: 'MAC_CHARSET',
        0x80: 'SHIFTJIS_CHARSET',
        0x81: 'HANGEUL_CHARSET',
        0x81: 'HANGUL_CHARSET',
        0x82: 'JOHAB_CHARSET',
        0x86: 'GB2312_CHARSET',
        0x88: 'CHINESEBIG5_CHARSET',
        0xA1: 'GREEK_CHARSET',
        0xA2: 'TURKISH_CHARSET',
        0xA3: 'VIETNAMESE_CHARSET',
        0xB1: 'HEBREW_CHARSET',
        0xB2: 'ARABIC_CHARSET',
        0xBA: 'BALTIC_CHARSET',
        0xCC: 'RUSSIAN_CHARSET',
        0xDD: 'THAI_CHARSET',
        0xEE: 'EASTEUROPE_CHARSET'
    }

    @staticmethod
    def getCharSetName (code):
        return getValueOrUnknown(Font.charSetNames, code)

    def parseBytes (self):
        height     = self.readUnsignedInt(2)
        flags      = self.readUnsignedInt(2)
        colorId    = self.readUnsignedInt(2)

        boldStyle  = self.readUnsignedInt(2)
        boldStyleName = '(unknown)'
        if boldStyle == 400:
            boldStyleName = 'normal'
        elif boldStyle == 700:
            boldStyleName = 'bold'

        superSub   = self.readUnsignedInt(2)
        ulStyle    = self.readUnsignedInt(1)
        fontFamily = self.readUnsignedInt(1)
        charSet    = self.readUnsignedInt(1)
        reserved   = self.readUnsignedInt(1)
        nameLen    = self.readUnsignedInt(1)
        fontName, nameLen = globals.getRichText(self.readRemainingBytes(), nameLen)
        self.appendLine("font height: %d"%height)
        self.appendLine("color ID: %d"%colorId)
        self.appendLine("bold style: %s (%d)"%(boldStyleName, boldStyle))
        self.appendLine("script type: %s"%Font.getScriptName(superSub))
        self.appendLine("underline type: %s"%Font.getUnderlineStyleName(ulStyle))
        self.appendLine("character set: %s"%Font.getCharSetName(charSet))
        self.appendLine("font family: %s"%Font.getFontFamily(fontFamily))
        self.appendLine("font name: %s (%d)"%(fontName, nameLen))


class XF(BaseRecordHandler):

    def parseBytes (self):
        fontId = self.readUnsignedInt(2)
        numId = self.readUnsignedInt(2)
        self.appendLine("font ID: %d"%fontId)
        self.appendLine("number format ID: %d"%numId)


class FeatureHeader(BaseRecordHandler):

    def parseBytes (self):
        recordType = self.readUnsignedInt(2)
        frtFlag = self.readUnsignedInt(2) # currently 0
        self.readBytes(8) # reserved (currently all 0)
        featureTypeId = self.readUnsignedInt(2)
        featureTypeText = 'unknown'
        if featureTypeId == 2:
            featureTypeText = 'enhanced protection'
        elif featureTypeId == 4:
            featureTypeText = 'smart tag'
        featureHdr = self.readUnsignedInt(1) # must be 1
        sizeHdrData = self.readSignedInt(4)
        sizeHdrDataText = 'byte size'
        if sizeHdrData == -1:
            sizeHdrDataText = 'size depends on feature type'

        self.appendLine("record type: 0x%4.4X (must match the header)"%recordType)
        self.appendLine("feature type: %d (%s)"%(featureTypeId, featureTypeText))
        self.appendLine("size of header data: %d (%s)"%(sizeHdrData, sizeHdrDataText))

        if featureTypeId == 2 and sizeHdrData == -1:
            # enhanced protection optionsss
            flags = self.readUnsignedInt(4)
            self.appendLine("enhanced protection flag: 0x%8.8X"%flags)

            optEditObj             = (flags & 0x00000001)
            optEditScenario        = (flags & 0x00000002)
            optFormatCells         = (flags & 0x00000004)
            optFormatColumns       = (flags & 0x00000008)
            optFormatRows          = (flags & 0x00000010)
            optInsertColumns       = (flags & 0x00000020)
            optInsertRows          = (flags & 0x00000040)
            optInsertLinks         = (flags & 0x00000080)
            optDeleteColumns       = (flags & 0x00000100)
            optDeleteRows          = (flags & 0x00000200)
            optSelectLockedCells   = (flags & 0x00000400)
            optSort                = (flags & 0x00000800)
            optUseAutofilter       = (flags & 0x00001000)
            optUsePivotReports     = (flags & 0x00002000)
            optSelectUnlockedCells = (flags & 0x00004000)
            self.appendLine("  edit object:             %s"%self.getEnabledDisabled(optEditObj))
            self.appendLine("  edit scenario:           %s"%self.getEnabledDisabled(optEditScenario))
            self.appendLine("  format cells:            %s"%self.getEnabledDisabled(optFormatCells))
            self.appendLine("  format columns:          %s"%self.getEnabledDisabled(optFormatColumns))
            self.appendLine("  format rows:             %s"%self.getEnabledDisabled(optFormatRows))
            self.appendLine("  insert columns:          %s"%self.getEnabledDisabled(optInsertColumns))
            self.appendLine("  insert rows:             %s"%self.getEnabledDisabled(optInsertRows))
            self.appendLine("  insert hyperlinks:       %s"%self.getEnabledDisabled(optInsertLinks))
            self.appendLine("  delete columns:          %s"%self.getEnabledDisabled(optDeleteColumns))
            self.appendLine("  delete rows:             %s"%self.getEnabledDisabled(optDeleteRows))
            self.appendLine("  select locked cells:     %s"%self.getEnabledDisabled(optSelectLockedCells))
            self.appendLine("  sort:                    %s"%self.getEnabledDisabled(optSort))
            self.appendLine("  use autofilter:          %s"%self.getEnabledDisabled(optUseAutofilter))
            self.appendLine("  use pivot table reports: %s"%self.getEnabledDisabled(optUsePivotReports))
            self.appendLine("  select unlocked cells:   %s"%self.getEnabledDisabled(optSelectUnlockedCells))

        return

# -------------------------------------------------------------------
# SX - Pivot Table

class SXViewEx9(BaseRecordHandler):

    def parseBytes (self):
        rt = self.readUnsignedInt(2)
        dummy = self.readBytes(6)
        flags = self.readUnsignedInt(4)
        autoFmtId = self.readUnsignedInt(2)

        self.appendLine("record type: %4.4Xh (always 0x0810)"%rt)
        self.appendLine("autoformat index: %d"%autoFmtId)

        nameLen = self.readSignedInt(2)
        if nameLen > 0:
            name, nameLen = globals.getRichText(self.readRemainingBytes(), nameLen)
            self.appendLine("grand total name: %s"%name)
        else:
            self.appendLine("grand total name: (none)")
        return


class SXAddlInfo(BaseRecordHandler):

    sxcNameList = {
        0x00: "sxcView",
        0x01: "sxcField",
        0x02: "sxcHierarchy",
        0x03: "sxcCache",
        0x04: "sxcCacheField",
        0x05: "sxcQsi",
        0x06: "sxcQuery",
        0x07: "sxcGrpLevel",
        0x08: "sxcGroup"
    }

    sxdNameList = {
        0x00: 'sxdId',
        0x01: 'sxdVerUpdInv',
        0x02: 'sxdVer10Info',
        0x03: 'sxdCalcMember',
        0x04: 'sxdXMLSource',
        0x05: 'sxdProperty',
        0x05: 'sxdSrcDataFile',
        0x06: 'sxdGrpLevelInfo',
        0x06: 'sxdSrcConnFile',
        0x07: 'sxdGrpInfo',
        0x07: 'sxdReconnCond',
        0x08: 'sxdMember',
        0x09: 'sxdFilterMember',
        0x0A: 'sxdCalcMemString',
        0xFF: 'sxdEnd'
    }

    def parseBytes (self):
        dummy = self.readBytes(2) # 0x0864
        dummy = self.readBytes(2) # 0x0000
        sxc = self.readBytes(1)[0]
        sxd = self.readBytes(1)[0]
        dwUserData = self.readBytes(4)
        dummy = self.readBytes(2)

        className = "(unknown)"
        if SXAddlInfo.sxcNameList.has_key(sxc):
            className = SXAddlInfo.sxcNameList[sxc]
        self.appendLine("class name: %s"%className)
        typeName = '(unknown)'
        if SXAddlInfo.sxdNameList.has_key(sxd):
            typeName = SXAddlInfo.sxdNameList[sxd]
        self.appendLine("type name: %s"%typeName)
        
        if sxd == 0x00:
            self.__parseId(sxc, dwUserData)

        elif sxd == 0x02:
            if sxc == 0x03:
                self.__parseSxDbSave10()
            elif sxc == 0x00:
                self.__parseViewFlags(dwUserData)

    def __parseViewFlags (self, dwUserData):
        flags = globals.getUnsignedInt(dwUserData)
        viewVer = (flags & 0x000000FF)
        verName = self.__getExcelVerName(viewVer)
        self.appendLine("PivotTable view version: %s"%verName)
        displayImmediateItems = (flags & 0x00000100)
        enableDataEd          = (flags & 0x00000200)
        disableFList          = (flags & 0x00000400)
        reenterOnLoadOnce     = (flags & 0x00000800)
        notViewCalcMembers    = (flags & 0x00001000)
        notVisualTotals       = (flags & 0x00002000)
        pageMultiItemLabel    = (flags & 0x00004000)
        tensorFillCv          = (flags & 0x00008000)
        hideDDData            = (flags & 0x00010000)

        self.appendLine("display immediate items: %s"%self.getYesNo(displayImmediateItems))
        self.appendLine("editing values in data area allowed: %s"%self.getYesNo(enableDataEd))
        self.appendLine("field list disabled: %s"%self.getYesNo(disableFList))
        self.appendLine("re-center on load once: %s"%self.getYesNo(reenterOnLoadOnce))
        self.appendLine("hide calculated members: %s"%self.getYesNo(notViewCalcMembers))
        self.appendLine("totals include hidden members: %s"%self.getYesNo(notVisualTotals))
        self.appendLine("(Multiple Items) instead of (All) in page field: %s"%self.getYesNo(pageMultiItemLabel))
        self.appendLine("background color from source: %s"%self.getYesNo(tensorFillCv))
        self.appendLine("hide drill-down for data field: %s"%self.getYesNo(hideDDData))

    def __parseId (self, sxc, dwUserData):
        if sxc == 0x03:
            idCache = globals.getUnsignedInt(dwUserData)
            self.appendLine("cache ID: %d"%idCache)
        elif sxc in [0x00, 0x01, 0x02, 0x05, 0x06, 0x07, 0x08]:
            lenStr = globals.getUnsignedInt(dwUserData)
            self.appendLine("length of ID string: %d"%lenStr)
            textLen = globals.getUnsignedInt(self.readBytes(2))
            data = self.bytes[self.getCurrentPos():]
            if lenStr == 0:
                self.appendLine("name (ID) string: (continued from last record)")
            elif lenStr == len(data) - 1:
                text, textLen = globals.getRichText(data, textLen)
                self.appendLine("name (ID) string: %s"%text)
            else:
                self.appendLine("name (ID) string: (first of multiple records)")


    def __parseSxDbSave10 (self):
        countGhostMax = globals.getSignedInt(self.readBytes(4))
        self.appendLine("max ghost pivot items: %d"%countGhostMax)

        # version last refreshed this cache
        lastVer = globals.getUnsignedInt(self.readBytes(1))
        verName = self.__getExcelVerName(lastVer)
        self.appendLine("last version refreshed: %s"%verName)
        
        # minimum version needed to refresh this cache
        lastVer = globals.getUnsignedInt(self.readBytes(1))
        verName = self.__getExcelVerName(lastVer)
        self.appendLine("minimum version needed to refresh: %s"%verName)

        # date last refreshed
        dateRefreshed = globals.getDouble(self.readBytes(8))
        self.appendLine("date last refreshed: %g"%dateRefreshed)


    def __getExcelVerName (self, ver):
        verName = '(unknown)'
        if ver == 0:
            verName = 'Excel 9 (2000) and earlier'
        elif ver == 1:
            verName = 'Excel 10 (XP)'
        elif ver == 2:
            verName = 'Excel 11 (2003)'
        elif ver == 3:
            verName = 'Excel 12 (2007)'
        return verName


class SXDb(BaseRecordHandler):

    def parseBytes (self):
        recCount = self.readUnsignedInt(4)
        strmId   = self.readUnsignedInt(2)
        flags    = self.readUnsignedInt(2)
        self.appendLine("number of records in database: %d"%recCount)
        self.appendLine("stream ID: %4.4Xh"%strmId)
#       self.appendLine("flags: %4.4Xh"%flags)

        saveLayout    = (flags & 0x0001)
        invalid       = (flags & 0x0002)
        refreshOnLoad = (flags & 0x0004)
        optimizeCache = (flags & 0x0008)
        backQuery     = (flags & 0x0010)
        enableRefresh = (flags & 0x0020)
        self.appendLine("save data with table layout: %s"%self.getYesNo(saveLayout))
        self.appendLine("invalid table (must be refreshed before next update): %s"%self.getYesNo(invalid))
        self.appendLine("refresh table on load: %s"%self.getYesNo(refreshOnLoad))
        self.appendLine("optimize cache for least memory use: %s"%self.getYesNo(optimizeCache))
        self.appendLine("query results obtained in the background: %s"%self.getYesNo(backQuery))
        self.appendLine("refresh is enabled: %s"%self.getYesNo(enableRefresh))

        dbBlockRecs = self.readUnsignedInt(2)
        baseFields = self.readUnsignedInt(2)
        allFields = self.readUnsignedInt(2)
        self.appendLine("number of records for each database block: %d"%dbBlockRecs)
        self.appendLine("number of base fields: %d"%baseFields)
        self.appendLine("number of all fields: %d"%allFields)

        dummy = self.readBytes(2)
        type = self.readUnsignedInt(2)
        typeName = '(unknown)'
        if type == 1:
            typeName = 'Excel worksheet'
        elif type == 2:
            typeName = 'External data'
        elif type == 4:
            typeName = 'Consolidation'
        elif type == 8:
            typeName = 'Scenario PivotTable'
        self.appendLine("type: %s (%d)"%(typeName, type))
        textLen = self.readUnsignedInt(2)
        changedBy, textLen = globals.getRichText(self.readRemainingBytes(), textLen)
        self.appendLine("changed by: %s"%changedBy)


class SXDbEx(BaseRecordHandler):

    def parseBytes (self):
        lastChanged = self.readDouble()
        sxFmlaRecs = self.readUnsignedInt(4)
        self.appendLine("last changed: %g"%lastChanged)
        self.appendLine("count of SXFORMULA records for this cache: %d"%sxFmlaRecs)


class SXField(BaseRecordHandler):

    dataTypeNames = {
        0x0000: 'spc',
        0x0480: 'str',
        0x0520: 'int[+dbl]',
        0x0560: 'dbl',
        0x05A0: 'str+int[+dbl]',
        0x05E0: 'str+dbl',
        0x0900: 'dat',
        0x0D00: 'dat+int/dbl',
        0x0D80: 'dat+str[+int/dbl]'
    }

    def parseBytes (self):
        flags = self.readUnsignedInt(2)
        origItems  = (flags & 0x0001)
        postponed  = (flags & 0x0002)
        calculated = (flags & 0x0004)
        groupChild = (flags & 0x0008)
        numGroup   = (flags & 0x0010)
        longIndex  = (flags & 0x0200)
        self.appendLine("original items: %s"%self.getYesNo(origItems))
        self.appendLine("postponed: %s"%self.getYesNo(postponed))
        self.appendLine("calculated: %s"%self.getYesNo(calculated))
        self.appendLine("group child: %s"%self.getYesNo(groupChild))
        self.appendLine("num group: %s"%self.getYesNo(numGroup))
        self.appendLine("long index: %s"%self.getYesNo(longIndex))
        dataType = (flags & 0x0DE0)
        if SXField.dataTypeNames.has_key(dataType):
            self.appendLine("data type: %s (%4.4Xh)"%(SXField.dataTypeNames[dataType], dataType))
        else:
            self.appendLine("data type: unknown (%4.4Xh)"%dataType)

        grpSubField = self.readUnsignedInt(2)
        grpBaseField = self.readUnsignedInt(2)
        itemCount = self.readUnsignedInt(2)
        grpItemCount = self.readUnsignedInt(2)
        baseItemCount = self.readUnsignedInt(2)
        srcItemCount = self.readUnsignedInt(2)
        self.appendLine("group sub-field: %d"%grpSubField)
        self.appendLine("group base-field: %d"%grpBaseField)
        self.appendLine("item count: %d"%itemCount)
        self.appendLine("group item count: %d"%grpItemCount)
        self.appendLine("base item count: %d"%baseItemCount)
        self.appendLine("source item count: %d"%srcItemCount)

        # field name
        textLen = self.readUnsignedInt(2)
        name, textLen = globals.getRichText(self.readRemainingBytes(), textLen)
        self.appendLine("field name: %s"%name)


class SXStreamID(BaseRecordHandler):

    def parseBytes (self):
        if self.size != 2:
            return

        strmId = globals.getSignedInt(self.bytes)
        self.strmData.appendPivotCacheId(strmId)
        self.appendLine("pivot cache stream ID in SX DB storage: %d"%strmId)


class SXView(BaseRecordHandler):

    def parseBytes (self):
        rowFirst = self.readUnsignedInt(2)
        rowLast  = self.readUnsignedInt(2)
        self.appendLine("row range: %d - %d"%(rowFirst, rowLast))

        colFirst = self.readUnsignedInt(2)
        colLast  = self.readUnsignedInt(2)
        self.appendLine("col range: %d - %d"%(colFirst,colLast))

        rowHeadFirst = self.readUnsignedInt(2)
        rowDataFirst = self.readUnsignedInt(2)
        colDataFirst = self.readUnsignedInt(2)
        self.appendLine("heading row: %d"%rowHeadFirst)
        self.appendLine("data row: %d"%rowDataFirst)
        self.appendLine("data col: %d"%colDataFirst)

        cacheIndex = self.readUnsignedInt(2)
        self.appendLine("cache index: %d"%cacheIndex)

        self.readBytes(2)

        dataFieldAxis = self.readUnsignedInt(2)
        self.appendLine("default data field axis: %d"%dataFieldAxis)

        dataFieldPos = self.readUnsignedInt(2)
        self.appendLine("default data field pos: %d"%dataFieldPos)

        numFields = self.readUnsignedInt(2)
        numRowFields = self.readUnsignedInt(2)
        numColFields = self.readUnsignedInt(2)
        numPageFields = self.readUnsignedInt(2)
        numDataFields = self.readUnsignedInt(2)
        numDataRows = self.readUnsignedInt(2)
        numDataCols = self.readUnsignedInt(2)
        self.appendLine("field count: %d"%numFields)
        self.appendLine("row field count: %d"%numRowFields)
        self.appendLine("col field count: %d"%numColFields)
        self.appendLine("page field count: %d"%numPageFields)
        self.appendLine("data field count: %d"%numDataFields)
        self.appendLine("data row count: %d"%numDataRows)
        self.appendLine("data col count: %d"%numDataCols)

        # option flags (TODO: display these later.)
        flags = self.readUnsignedInt(2)

        # autoformat index
        autoFmtIndex = self.readUnsignedInt(2)
        self.appendLine("autoformat index: %d"%autoFmtIndex)

        nameLenTable = self.readUnsignedInt(2)
        nameLenDataField = self.readUnsignedInt(2)
        text, nameLenTable = globals.getRichText(self.readBytes(nameLenTable+1), nameLenTable)
        self.appendLine("PivotTable name: %s"%text)
        text, nameLenDataField = globals.getRichText(self.readBytes(nameLenDataField+1), nameLenDataField)
        self.appendLine("data field name: %s"%text)


class SXViewSource(BaseRecordHandler):

    def parseBytes (self):
        if self.size != 2:
            return

        src = globals.getSignedInt(self.bytes)
        srcType = 'unknown'
        if src == 0x01:
            srcType = 'Excel list or database'
        elif src == 0x02:
            srcType = 'External data source (Microsoft Query)'
        elif src == 0x04:
            srcType = 'Multiple consolidation ranges'
        elif src == 0x10:
            srcType = 'Scenario Manager summary report'

        self.appendLine("data source type: %s"%srcType)


class SXViewFields(BaseRecordHandler):

    def parseBytes (self):
        axis          = globals.getSignedInt(self.readBytes(2))
        subtotalCount = globals.getSignedInt(self.readBytes(2))
        subtotalType  = globals.getSignedInt(self.readBytes(2))
        itemCount     = globals.getSignedInt(self.readBytes(2))
        nameLen       = globals.getSignedInt(self.readBytes(2))
        
        axisType = 'unknown'
        if axis == 0:
            axisType = 'no axis'
        elif axis == 1:
            axisType = 'row'
        elif axis == 2:
            axisType = 'column'
        elif axis == 4:
            axisType = 'page'
        elif axis == 8:
            axisType = 'data'

        subtotalTypeName = 'unknown'
        if subtotalType == 0x0000:
            subtotalTypeName = 'None'
        elif subtotalType == 0x0001:
            subtotalTypeName = 'Default'
        elif subtotalType == 0x0002:
            subtotalTypeName = 'Sum'
        elif subtotalType == 0x0004:
            subtotalTypeName = 'CountA'
        elif subtotalType == 0x0008:
            subtotalTypeName = 'Average'
        elif subtotalType == 0x0010:
            subtotalTypeName = 'Max'
        elif subtotalType == 0x0020:
            subtotalTypeName = 'Min'
        elif subtotalType == 0x0040:
            subtotalTypeName = 'Product'
        elif subtotalType == 0x0080:
            subtotalTypeName = 'Count'
        elif subtotalType == 0x0100:
            subtotalTypeName = 'Stdev'
        elif subtotalType == 0x0200:
            subtotalTypeName = 'StdevP'
        elif subtotalType == 0x0400:
            subtotalTypeName = 'Var'
        elif subtotalType == 0x0800:
            subtotalTypeName = 'VarP'

        self.appendLine("axis type: %s"%axisType)
        self.appendLine("number of subtotals: %d"%subtotalCount)
        self.appendLine("subtotal type: %s"%subtotalTypeName)
        self.appendLine("number of items: %d"%itemCount)

        if nameLen == -1:
            self.appendLine("name: null (use name in the cache)")
        else:
            name, nameLen = globals.getRichText(self.readRemainingBytes(), nameLen)
            self.appendLine("name: %s"%name)


class SXViewFieldsEx(BaseRecordHandler):

    def parseBytes (self):
        grbit1 = self.readUnsignedInt(2)
        grbit2 = self.readUnsignedInt(1)
        numItemAutoShow = self.readUnsignedInt(1)
        dataFieldSort = self.readSignedInt(2)
        dataFieldAutoShow = self.readSignedInt(2)
        numFmt = self.readUnsignedInt(2)

        # custom name length: it can be up to 254 characters.  255 (0xFF) if
        # no custom name is used.
        nameLen = self.readUnsignedInt(1)

        dummy = self.readBytes(10)
        
        self.appendLine("auto show item: %d"%numItemAutoShow)
        self.appendLine("sort field index: %d"%dataFieldSort)
        self.appendLine("auto show field index: %d"%dataFieldAutoShow)
        self.appendLine("number format: %d"%numFmt)

        if nameLen == 0xFF:
            self.appendLine("custome name: none")
        else:
            name = globals.getTextBytes(self.readRemainingBytes())
            self.appendLine("custome name: %s"%name)

        return


class SXDataItem(BaseRecordHandler):

    functionType = {
        0x00: 'sum',
        0x01: 'count',
        0x02: 'average',
        0x03: 'max',
        0x04: 'min',
        0x05: 'product',
        0x06: 'count nums',
        0x07: 'stddev',
        0x08: 'stddevp',
        0x09: 'var',
        0x0A: 'varp'
    }

    displayFormat = {
        0x00: 'normal',
        0x01: 'difference from',
        0x02: 'percentage of',
        0x03: 'perdentage difference from',
        0x04: 'running total in',
        0x05: 'percentage of row',
        0x06: 'percentage of column',
        0x07: 'percentage of total',
        0x08: 'index'
    }

    def parseBytes (self):
        isxvdData = self.readUnsignedInt(2)
        funcIndex = self.readUnsignedInt(2)

        # data display format
        df = self.readUnsignedInt(2)

        # index to the SXVD/SXVI records used by the data display format
        sxvdIndex = self.readUnsignedInt(2)
        sxviIndex = self.readUnsignedInt(2)

        # index to the format table for this item
        fmtIndex = self.readUnsignedInt(2)

        # name
        nameLen = self.readSignedInt(2)
        name, nameLen = globals.getRichText(self.readRemainingBytes(), nameLen)

        self.appendLine("field that this data item is based on: %d"%isxvdData)
        funcName = '(unknown)'
        if SXDataItem.functionType.has_key(funcIndex):
            funcName = SXDataItem.functionType[funcIndex]
        self.appendLine("aggregate function: %s"%funcName)
        dfName = '(unknown)'
        if SXDataItem.displayFormat.has_key(df):
            dfName = SXDataItem.displayFormat[df]
        self.appendLine("data display format: %s"%dfName)
        self.appendLine("SXVD record index: %d"%sxvdIndex)
        self.appendLine("SXVI record index: %d"%sxviIndex)
        self.appendLine("format table index: %d"%fmtIndex)

        if nameLen == -1:
            self.appendLine("name: null (use name in the cache)")
        else:
            self.appendLine("name: %s"%name)

        return


class SXViewItem(BaseRecordHandler):

    itemTypes = {
        0xFE: 'Page',
        0xFF: 'Null',
        0x00: 'Data',
        0x01: 'Default',
        0x02: 'SUM',
        0x03: 'COUNTA',
        0x04: 'COUNT',
        0x05: 'AVERAGE',
        0x06: 'MAX',
        0x07: 'MIN',
        0x08: 'PRODUCT',
        0x09: 'STDEV',
        0x0A: 'STDEVP',
        0x0B: 'VAR',
        0x0C: 'VARP',
        0x0D: 'Grand total',
        0x0E: 'blank'
    }

    def parseBytes (self):
        itemType = self.readSignedInt(2)
        grbit    = self.readSignedInt(2)
        iCache   = self.readSignedInt(2)
        nameLen  = self.readSignedInt(2)
        
        itemTypeName = 'unknown'
        if SXViewItem.itemTypes.has_key(itemType):
            itemTypeName = SXViewItem.itemTypes[itemType]

        flags = ''
        if (grbit & 0x0001):
            flags += 'hidden, '
        if (grbit & 0x0002):
            flags += 'detail hidden, '
        if (grbit & 0x0008):
            flags += 'formula, '
        if (grbit & 0x0010):
            flags += 'missing, '

        if len(flags) > 0:
            # strip the trailing ', '
            flags = flags[:-2]
        else:
            flags = '(none)'

        self.appendLine("item type: %s"%itemTypeName)
        self.appendLine("flags: %s"%flags)
        self.appendLine("pivot cache index: %d"%iCache)
        if nameLen == -1:
            self.appendLine("name: null (use name in the cache)")
        else:
            name, nameLen = globals.getRichText(self.readRemainingBytes(), nameLen)
            self.appendLine("name: %s"%name)


class PivotQueryTableEx(BaseRecordHandler):
    """QSISXTAG: Pivot Table and Query Table Extensions (802h)"""
    excelVersionList = [
        'Excel 2000',
        'Excel XP',
        'Excel 2003',
        'Excel 2007'
    ]

    class TableType:
        QueryTable = 0
        PivotTable = 1

    def getExcelVersion (self, lastExcelVer):
        s = '(unknown)'
        if lastExcelVer < len(PivotQueryTableEx.excelVersionList):
            s = PivotQueryTableEx.excelVersionList[lastExcelVer]
        return s

    def parseBytes (self):
        recordType = self.readUnsignedInt(2)
        self.appendLine("record type (always 0802h): %4.4Xh"%recordType)
        dummyFlags = self.readUnsignedInt(2)
        self.appendLine("flags (must be zero): %4.4Xh"%dummyFlags)
        tableType = self.readUnsignedInt(2)
        s = '(unknown)'
        if tableType == PivotQueryTableEx.TableType.QueryTable:
            s = 'query table'
        elif tableType == PivotQueryTableEx.TableType.PivotTable:
            s = 'pivot table'
        self.appendLine("this record is for: %s"%s)

        # general flags
        flags = self.readUnsignedInt(2)
        enableRefresh = (flags & 0x0001)
        invalid       = (flags & 0x0002)
        tensorEx      = (flags & 0x0004)
        s = '(unknown)'
        if enableRefresh:
            s = 'ignore'
        else:
            s = 'check'
        self.appendLine("check for SXDB or QSI for table refresh: %s"%s)
        self.appendLine("PivotTable cache is invalid: %s"%self.getYesNo(invalid))
        self.appendLine("This is an OLAP PivotTable report: %s"%self.getYesNo(tensorEx))

        # feature-specific options
        featureOptions = self.readUnsignedInt(4)
        if tableType == PivotQueryTableEx.TableType.QueryTable:
            # query table
            preserveFormat = (featureOptions & 0x00000001)
            autoFit        = (featureOptions & 0x00000002)
            self.appendLine("keep formatting applied by the user: %s"%self.getYesNo(preserveFormat))
            self.appendLine("auto-fit columns after refresh: %s"%self.getYesNo(autoFit))
        elif tableType == PivotQueryTableEx.TableType.PivotTable:
            # pivot table
            noStencil         = (featureOptions & 0x00000001)
            hideTotAnnotation = (featureOptions & 0x00000002)
            includeEmptyRow   = (featureOptions & 0x00000008)
            includeEmptyCol   = (featureOptions & 0x00000010)
            self.appendLine("no large drop zones if no data fields: %s"%self.getTrueFalse(noStencil))
            self.appendLine("no asterisk for the total in OLAP table: %s"%self.getTrueFalse(hideTotAnnotation))
            self.appendLine("retrieve and show empty rows from OLAP source: %s"%self.getTrueFalse(includeEmptyRow))
            self.appendLine("retrieve and show empty columns from OLAP source: %s"%self.getTrueFalse(includeEmptyCol))

        self.appendLine("table last refreshed by: %s"%
            self.getExcelVersion(self.readUnsignedInt(1)))

        self.appendLine("minimal version that can refresh: %s"%
            self.getExcelVersion(self.readUnsignedInt(1)))

        offsetBytes = self.readUnsignedInt(1)
        self.appendLine("offset from first FRT byte to first cchName byte: %d"%offsetBytes)

        self.appendLine("first version that created the table: %s"%
            self.getExcelVersion(self.readUnsignedInt(1)))

        textLen = self.readUnsignedInt(2)
        name, textLen = globals.getRichText(self.readRemainingBytes(), textLen)
        self.appendLine("table name: %s"%name)
        return


class SXDouble(BaseRecordHandler):
    def parseBytes (self):
        val = self.readDouble()
        self.appendLine("value: %g"%val)


class SXBoolean(BaseRecordHandler):
    def parseBytes (self):
        pass

class SXError(BaseRecordHandler):
    def parseBytes (self):
        pass


class SXInteger(BaseRecordHandler):
    def parseBytes (self):
        pass


class SXString(BaseRecordHandler):
    def parseBytes (self):
        textLen = self.readUnsignedInt(2)
        text, textLen = globals.getRichText(self.readRemainingBytes(), textLen)
        self.appendLine("value: %s"%text)

# -------------------------------------------------------------------
# CT - Change Tracking

class CTCellContent(BaseRecordHandler):
    
    EXC_CHTR_TYPE_MASK       = 0x0007
    EXC_CHTR_TYPE_FORMATMASK = 0xFF00
    EXC_CHTR_TYPE_EMPTY      = 0x0000
    EXC_CHTR_TYPE_RK         = 0x0001
    EXC_CHTR_TYPE_DOUBLE     = 0x0002
    EXC_CHTR_TYPE_STRING     = 0x0003
    EXC_CHTR_TYPE_BOOL       = 0x0004
    EXC_CHTR_TYPE_FORMULA    = 0x0005

    def parseBytes (self):
        size = globals.getSignedInt(self.readBytes(4))
        id = globals.getSignedInt(self.readBytes(4))
        opcode = globals.getSignedInt(self.readBytes(2))
        accept = globals.getSignedInt(self.readBytes(2))
        tabCreateId = globals.getSignedInt(self.readBytes(2))
        valueType = globals.getSignedInt(self.readBytes(2))
        self.appendLine("header: (size=%d; index=%d; opcode=0x%2.2X; accept=%d)"%(size, id, opcode, accept))
        self.appendLine("sheet creation id: %d"%tabCreateId)

        oldType = (valueType/(2*2*2) & CTCellContent.EXC_CHTR_TYPE_MASK)
        newType = (valueType & CTCellContent.EXC_CHTR_TYPE_MASK)
        self.appendLine("value type: (old=%4.4Xh; new=%4.4Xh)"%(oldType, newType))
        self.readBytes(2) # ignore next 2 bytes.

        row = globals.getSignedInt(self.readBytes(2))
        col = globals.getSignedInt(self.readBytes(2))
        cell = formula.CellAddress(col, row)
        self.appendLine("cell position: %s"%cell.getName())

        oldSize = globals.getSignedInt(self.readBytes(2))
        self.readBytes(4) # ignore 4 bytes.

        fmtType = (valueType & CTCellContent.EXC_CHTR_TYPE_FORMATMASK)
        if fmtType == 0x1100:
            self.readBytes(16)
        elif fmtType == 0x1300:
            self.readBytes(8)

        self.readCell(oldType, "old cell type")
        self.readCell(newType, "new cell type")

    def readCell (self, cellType, cellName):

        cellTypeText = 'unknown'

        if cellType == CTCellContent.EXC_CHTR_TYPE_FORMULA:
            cellTypeText, formulaBytes, formulaText = self.readFormula()
            self.appendLine("%s: %s"%(cellName, cellTypeText))
            self.appendLine("formula bytes: %s"%globals.getRawBytes(formulaBytes, True, False))
            self.appendLine("tokens: %s"%formulaText)
            return

        if cellType == CTCellContent.EXC_CHTR_TYPE_EMPTY:
            cellTypeText = 'empty'
        elif cellType == CTCellContent.EXC_CHTR_TYPE_RK:
            cellTypeText = self.readRK()
        elif cellType == CTCellContent.EXC_CHTR_TYPE_DOUBLE:
            cellTypeText = self.readDouble()
        elif cellType == CTCellContent.EXC_CHTR_TYPE_STRING:
            cellTypeText = self.readString()
        elif cellType == CTCellContent.EXC_CHTR_TYPE_BOOL:
            cellTypeText = self.readBool()
        elif cellType == CTCellContent.EXC_CHTR_TYPE_FORMULA:
            cellTypeText, formulaText = self.readFormula()

        self.appendLine("%s: %s"%(cellName, cellTypeText))

    def readRK (self):
        valRK = globals.getSignedInt(self.readBytes(4))
        return 'RK value'

    def readDouble (self):
        val = globals.getDouble(self.readBytes(4))
        return "value %f"%val

    def readString (self):
        size = globals.getSignedInt(self.readBytes(2))
        pos = self.getCurrentPos()
        name, byteLen = globals.getRichText(self.bytes[pos:], size)
        self.setCurrentPos(pos + byteLen)
        return "string '%s'"%name

    def readBool (self):
        bool = globals.getSignedInt(self.readBytes(2))
        return "bool (%d)"%bool

    def readFormula (self):
        size = globals.getSignedInt(self.readBytes(2))
        fmlaBytes = self.readBytes(size)
        o = formula.FormulaParser(self.header, fmlaBytes, False)
        o.parse()
        return "formula", fmlaBytes, o.getText()









# -------------------------------------------------------------------
# CH - Chart

class CHChart(BaseRecordHandler):

    def parseBytes (self):
        x = globals.getSignedInt(self.bytes[0:4])
        y = globals.getSignedInt(self.bytes[4:8])
        w = globals.getSignedInt(self.bytes[8:12])
        h = globals.getSignedInt(self.bytes[12:16])
        self.appendLine("position: (x, y) = (%d, %d)"%(x, y))
        self.appendLine("size: (width, height) = (%d, %d)"%(w, h))
        
        
class CHSeries(BaseRecordHandler):

    DATE     = 0
    NUMERIC  = 1
    SEQUENCE = 2
    TEXT     = 3

    seriesTypes = ['date', 'numeric', 'sequence', 'text']

    @staticmethod
    def getSeriesType (idx):
        r = 'unknown'
        if idx < len(CHSeries.seriesTypes):
            r = CHSeries.seriesTypes[idx]
        return r

    def parseBytes (self):
        catType     = self.readUnsignedInt(2)
        valType     = self.readUnsignedInt(2)
        catCount    = self.readUnsignedInt(2)
        valCount    = self.readUnsignedInt(2)
        bubbleType  = self.readUnsignedInt(2)
        bubbleCount = self.readUnsignedInt(2)

        self.appendLine("category type: %s (count: %d)"%
            (CHSeries.getSeriesType(catType), catCount))
        self.appendLine("value type: %s (count: %d)"%
            (CHSeries.getSeriesType(valType), valCount))
        self.appendLine("bubble type: %s (count: %d)"%
            (CHSeries.getSeriesType(bubbleType), bubbleCount))


class CHAxis(BaseRecordHandler):

    axisTypeList = ['x-axis', 'y-axis', 'z-axis']

    def parseBytes (self):
        axisType = self.readUnsignedInt(2)
        x = self.readSignedInt(4)
        y = self.readSignedInt(4)
        w = self.readSignedInt(4)
        h = self.readSignedInt(4)
        if axisType < len(CHAxis.axisTypeList):
            self.appendLine("axis type: %s (%d)"%(CHAxis.axisTypeList[axisType], axisType))
        else:
            self.appendLine("axis type: unknown")
        self.appendLine("area: (x, y, w, h) = (%d, %d, %d, %d) [no longer used]"%(x, y, w, h))


class CHProperties(BaseRecordHandler):

    def parseBytes (self):
        flags = globals.getSignedInt(self.bytes[0:2])
        emptyFlags = globals.getSignedInt(self.bytes[2:4])

        manualSeries   = "false"
        showVisCells   = "false"
        noResize       = "false"
        manualPlotArea = "false"

        if (flags & 0x0001):
            manualSeries = "true"
        if (flags & 0x0002):
            showVisCells = "true"
        if (flags & 0x0004):
            noResize = "true"
        if (flags & 0x0008):
            manualPlotArea = "true"

        self.appendLine("manual series: %s"%manualSeries)
        self.appendLine("show only visible cells: %s"%showVisCells)
        self.appendLine("no resize: %s"%noResize)
        self.appendLine("manual plot area: %s"%manualPlotArea)

        emptyValues = "skip"
        if emptyFlags == 1:
            emptyValues = "plot as zero"
        elif emptyFlags == 2:
            emptyValues = "interpolate empty values"

        self.appendLine("empty value treatment: %s"%emptyValues)


class CHLabelRange(BaseRecordHandler):

    def parseBytes (self):
        axisCross = self.readUnsignedInt(2)
        freqLabel = self.readUnsignedInt(2)
        freqTick  = self.readUnsignedInt(2)
        self.appendLine("axis crossing: %d"%axisCross)
        self.appendLine("label frequency: %d"%freqLabel)
        self.appendLine("tick frequency: %d"%freqTick)

        flags     = self.readUnsignedInt(2)
        betweenCateg = (flags & 0x0001)
        maxCross     = (flags & 0x0002)
        reversed     = (flags & 0x0004)
        self.appendLineBoolean("axis between categories", betweenCateg)
        self.appendLineBoolean("other axis crosses at maximum", maxCross)
        self.appendLineBoolean("axis reversed", reversed)


class CHLegend(BaseRecordHandler):
    
    dockModeMap = {0: 'bottom', 1: 'corner', 2: 'top', 3: 'right', 4: 'left', 7: 'not docked'}
    spacingMap = ['close', 'medium', 'open']

    def getDockModeText (self, val):
        if CHLegend.dockModeMap.has_key(val):
            return CHLegend.dockModeMap[val]
        else:
            return '(unknown)'

    def getSpacingText (self, val):
        if val < len(CHLegend.spacingMap):
            return CHLegend.spacingMap[val]
        else:
            return '(unknown)'

    def parseBytes (self):
        x = self.readSignedInt(4)
        y = self.readSignedInt(4)
        w = self.readSignedInt(4)
        h = self.readSignedInt(4)
        dockMode = self.readUnsignedInt(1)
        spacing  = self.readUnsignedInt(1)
        flags    = self.readUnsignedInt(2)

        docked     = (flags & 0x0001)
        autoSeries = (flags & 0x0002)
        autoPosX   = (flags & 0x0004)
        autoPosY   = (flags & 0x0008)
        stacked    = (flags & 0x0010)
        dataTable  = (flags & 0x0020)

        self.appendLine("legend position: (x, y) = (%d, %d)"%(x,y))
        self.appendLine("legend size: width = %d, height = %d"%(w,h))
        self.appendLine("dock mode: %s"%self.getDockModeText(dockMode))
        self.appendLine("spacing: %s"%self.getSpacingText(spacing))
        self.appendLineBoolean("docked", docked)
        self.appendLineBoolean("auto series", autoSeries)
        self.appendLineBoolean("auto position x", autoPosX)
        self.appendLineBoolean("auto position y", autoPosY)
        self.appendLineBoolean("stacked", stacked)
        self.appendLineBoolean("data table", dataTable)

        self.appendLine("")
        self.appendMultiLine("NOTE: Position and size are in units of 1/4000 of chart's width or height.")


class CHValueRange(BaseRecordHandler):

    def parseBytes (self):
        minVal = globals.getDouble(self.readBytes(8))
        maxVal = globals.getDouble(self.readBytes(8))
        majorStep = globals.getDouble(self.readBytes(8))
        minorStep = globals.getDouble(self.readBytes(8))
        cross = globals.getDouble(self.readBytes(8))
        flags = globals.getSignedInt(self.readBytes(2))

        autoMin   = (flags & 0x0001)
        autoMax   = (flags & 0x0002)
        autoMajor = (flags & 0x0004)
        autoMinor = (flags & 0x0008)
        autoCross = (flags & 0x0010)
        logScale  = (flags & 0x0020)
        reversed  = (flags & 0x0040)
        maxCross  = (flags & 0x0080)
        bit8      = (flags & 0x0100)

        self.appendLine("min: %g (auto min: %s)"%(minVal, self.getYesNo(autoMin)))
        self.appendLine("max: %g (auto max: %s)"%(maxVal, self.getYesNo(autoMax)))
        self.appendLine("major step: %g (auto major: %s)"%
            (majorStep, self.getYesNo(autoMajor)))
        self.appendLine("minor step: %g (auto minor: %s)"%
            (minorStep, self.getYesNo(autoMinor)))
        self.appendLine("cross: %g (auto cross: %s) (max cross: %s)"%
            (cross, self.getYesNo(autoCross), self.getYesNo(maxCross)))
        self.appendLine("biff5 or above: %s"%self.getYesNo(bit8))


class CHBar(BaseRecordHandler):

    def parseBytes (self):
        overlap = globals.getSignedInt(self.readBytes(2))
        gap     = globals.getSignedInt(self.readBytes(2))
        flags   = globals.getUnsignedInt(self.readBytes(2))

        horizontal = (flags & 0x0001)
        stacked    = (flags & 0x0002)
        percent    = (flags & 0x0004)
        shadow     = (flags & 0x0008)

        self.appendLine("overlap width: %d"%overlap)
        self.appendLine("gap: %d"%gap)
        self.appendLine("horizontal: %s"%self.getYesNo(horizontal))
        self.appendLine("stacked: %s"%self.getYesNo(stacked))
        self.appendLine("percent: %s"%self.getYesNo(percent))
        self.appendLine("shadow: %s"%self.getYesNo(shadow))


class CHLine(BaseRecordHandler):

    def parseBytes (self):
        flags   = globals.getUnsignedInt(self.readBytes(2))
        stacked = (flags & 0x0001)
        percent = (flags & 0x0002)
        shadow  = (flags & 0x0004)

        self.appendLine("stacked: %s"%self.getYesNo(stacked))
        self.appendLine("percent: %s"%self.getYesNo(percent))
        self.appendLine("shadow: %s"%self.getYesNo(shadow))


class CHSourceLink(BaseRecordHandler):

    destTypes = ['title', 'values', 'category', 'bubbles']
    linkTypes = ['default', 'directly', 'worksheet']

    DEST_TITLE    = 0;
    DEST_VALUES   = 1;
    DEST_CATEGORY = 2;
    DEST_BUBBLES  = 3;

    LINK_DEFAULT   = 0;
    LINK_DIRECTLY  = 1;
    LINK_WORKSHEET = 2;

    def parseBytes (self):
        destType = self.readUnsignedInt(1)
        linkType = self.readUnsignedInt(1)
        flags    = self.readUnsignedInt(2)
        numFmt   = self.readUnsignedInt(2)
        
        destName = 'unknown'
        if destType < len(CHSourceLink.destTypes):
            destName = CHSourceLink.destTypes[destType]

        linkName = 'unknown'
        if linkType < len(CHSourceLink.linkTypes):
            linkName = CHSourceLink.linkTypes[linkType]

        self.appendLine("destination type: %s"%destName)
        self.appendLine("link type: %s"%linkName)

        if linkType == CHSourceLink.LINK_WORKSHEET:
            # external reference.  Read the formula tokens.
            lenToken = self.readUnsignedInt(2)
            tokens = self.readBytes(lenToken)
            self.appendLine("formula tokens: %s"%globals.getRawBytes(tokens,True,False))


class MSODrawing(BaseRecordHandler):

    class RecordHeader:
        def __init__ (self):
            self.recVer = None
            self.recInstance = None
            self.recType = None
            self.recLen = None

    class FDG:
        def __init__ (self):
            self.shapeCount = None
            self.lastShapeID = -1

    class FOPT:
        def __init__ (self):
            self.properties = None

    class FRITContainer:
        def __init__ (self):
            self.frits = [] # array of FRIT objects

    class FRIT:
        def __init__ (self):
            self.lastGroupID = None
            self.secondLastGroupID = None

    class FSP:
        def __init__ (self):
            self.spid = None
            self.flag = None

    class FSPGR:
        def __init__ (self):
            self.left   = None
            self.top    = None
            self.right  = None
            self.bottom = None

    class FConnectorRule:
        def __init__ (self):
            self.ruleID = None
            self.spIDA = None
            self.spIDB = None
            self.spIDC = None
            self.conSiteIDA = None
            self.conSiteIDB = None

    class SpgrContainer:
        def __init__ (self):
            self.fileblock = 0x0000

    class SpContainer:
        def __init__ (self):
            pass


    def readRecordHeader (self):
        rh = MSODrawing.RecordHeader()
        mixed = self.readUnsignedInt(2)
        rh.recVer = (mixed & 0x000F)
        rh.recInstance = (mixed & 0xFFF0) / 16
        rh.recType = self.readUnsignedInt(2)
        rh.recLen  = self.readUnsignedInt(4)
        return rh

    class Type:
        dgContainer     = 0xF002
        spgrContainer   = 0xF003
        spContainer     = 0xF004
        solverContainer = 0xF005
        FDG             = 0xF008
        FSPGR           = 0xF009
        FSP             = 0xF00A
        FOPT            = 0xF00B
        FConnectorRule  = 0xF012

    containerTypeNames = {
        Type.dgContainer:      'OfficeArtDgContainer',
        Type.spContainer:      'OfficeArtSpContainer',
        Type.spgrContainer:    'OfficeArtSpgrContainer',
        Type.solverContainer:  'OfficeArtSolverContainer',
        Type.FDG:              'OfficeArtFDG',
        Type.FOPT:             'OfficeArtFOPT',
        Type.FSP:              'OfficeArtFSP',
        Type.FSPGR:            'OfficeArtFSPGR',
        Type.FConnectorRule:   'OfficeArtFConnectorRule'
    }

    @staticmethod
    def getRecTypeName (rh):
        if MSODrawing.containerTypeNames.has_key(rh.recType):
            return MSODrawing.containerTypeNames[rh.recType]
        return 'unknown'

    def parseBytes (self):
        while not self.isEndOfRecord():
            rh = self.readRecordHeader()
            self.printRecordHeader(rh)
            # if rh.recType == MSODrawing.Type.dgContainer:
            if rh.recVer == 0xF:
                # container
                pass
            elif rh.recType == MSODrawing.Type.FDG:
                fdg = self.readFDG()
            elif rh.recType == MSODrawing.Type.FOPT:
                fopt = self.readFOPT(rh)
            elif rh.recType == MSODrawing.Type.FSPGR:
                fspgr = self.readFSPGR()
            elif rh.recType == MSODrawing.Type.FSP:
                fspgr = self.readFSP()
            elif rh.recType == MSODrawing.Type.FConnectorRule:
                fcon = self.readFConnectorRule()
            else:
                # unknown object
                self.readBytes(rh.recLen)

    def readFConnectorRule (self):
        fcon = MSODrawing.FConnectorRule()
        fcon.ruleID = self.readUnsignedInt(4)
        fcon.spIDA = self.readUnsignedInt(4)
        fcon.spIDB = self.readUnsignedInt(4)
        fcon.spIDC = self.readUnsignedInt(4)
        fcon.conSiteIDA = self.readUnsignedInt(4)
        fcon.conSiteIDB = self.readUnsignedInt(4)
        return fcon

    def readFDG (self):
        fdg = MSODrawing.FDG()
        fdg.shapeCount  = self.readUnsignedInt(4)
        fdg.lastShapeID = self.readUnsignedInt(4)
        return fdg

    def readFOPT (self, rh):
        fopt = MSODrawing.FOPT()
        fopt.properties = self.readBytes(rh.recLen)
        return fopt

    def readFSP (self):
        fsp = MSODrawing.FSP()
        fsp.spid = self.readUnsignedInt(4)
        fsp.flag = self.readUnsignedInt(4)
        return fsp

    def readFSPGR (self):
        fspgr = MSODrawing.FSPGR()
        fspgr.left   = self.readSignedInt(4)
        fspgr.top    = self.readSignedInt(4)
        fspgr.right  = self.readSignedInt(4)
        fspgr.bottom = self.readSignedInt(4)
        return fspgr

    def readFRIT (self):
        frit = MSODrawing.FRIT()
        frit.lastGroupID = self.readUnsignedInt(2)
        frit.secondLastGroupID = self.readUnsignedInt(2)
        return frit

#   def readFRITContainer (self):
#       frits = MSODrawing.FRITContainer()
#       frits.rh = self.readRecordHeader()
#       fritCount = frits.rh.recInstance
#       for i in xrange(0, fritCount):
#           frits.frits.append(self.readFRIT())
#       return frits
#
#   def readSpgrContainer (self):
#       spgr = MSODrawing.SpgrContainer()
#       spgr.rh = self.readRecordHeader()
#       spgr.fileblock = self.readBytes(spgr.rh.recLen)
#       return spgr
#
#   def readSpContainer (self):
#       sp = MSODrawing.SpContainer()
#       sp.rh = self.readRecordHeader()
#       return sp

    def printRecordHeader (self, rh, level=0):
        unitIndent = ' '*2
        indent = unitIndent*level
        self.appendLine(indent + "record header:")
        self.appendLine(indent + unitIndent + "recVer: 0x%1.1X"%rh.recVer)
        self.appendLine(indent + unitIndent + "recInstance: 0x%3.3X"%rh.recInstance)
        self.appendLine(indent + unitIndent + "recType: 0x%4.4X (%s)"%(rh.recType, MSODrawing.getRecTypeName(rh)))
        self.appendLine(indent + unitIndent + "recLen: %d"%rh.recLen)

#   def readDgContainer (self):
#
#       # drawingData (16 bytes)
#       self.appendLine("DgContainer content:")
#       self.appendLine("== Drawing Data (%d) =="%self.getCurrentPos())
#       fdg = self.readFDG()
#       self.printRecordHeader(fdg.rh, 1)
#       self.appendLine("  ID of this drawing: %d"%fdg.rh.recInstance)
#       self.appendLine("  ID of last drawing: %d"%fdg.lastShapeID)
#       self.appendLine("  shape count: %d"%fdg.shapeCount)
#
#       # regroupItems (variable)
#       self.appendLine("== Regroup Items (%d) =="%self.getCurrentPos())
#       frits = self.readFRITContainer()
#       self.printRecordHeader(frits.rh, 1)
#       self.appendLine("  FRIT record count: %d"%frits.rh.recInstance)
#       self.appendLine("  FRIT record bytes: %d"%frits.rh.recLen)
#
#       # groupShape (variable)
#       self.appendLine("== Group Shape (%d) =="%self.getCurrentPos())
#       spgr = self.readSpgrContainer()
#       self.printRecordHeader(spgr.rh, 1)
#       # self.appendLine("  fileblock bytes: %s"%globals.getRawBytes(spgr.fileblock, True, False))
#
#       # shape (variable)
#       self.appendLine("== Shape (%d) =="%self.getCurrentPos())
#       sp = self.readSpContainer()
#       self.printRecordHeader(sp.rh, 1)
#
#       return
#
