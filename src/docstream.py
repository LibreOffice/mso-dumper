#!/usr/bin/env python

import ole
import struct
from docdirstream import DOCDirStream
import docrecord

class DOCFile:
    """Represents the whole word file - feed will all bytes."""
    def __init__ (self, chars, params):
        self.chars = chars
        self.size = len(self.chars)
        self.params = params

        self.header = ole.Header(self.chars, self.params)
        self.pos = self.header.parse()

    def __getDirectoryObj(self):
        obj = self.header.getDirectory()
        obj.parseDirEntries()
        return obj

    def getDirectoryNames(self):
        return self.__getDirectoryObj().getDirectoryNames()

    def getDirectoryStreamByName(self, name):
        obj = self.__getDirectoryObj()
        bytes = obj.getRawStreamByName(name)
        if name == "WordDocument":
            return WordDocumentStream(bytes, self.params, doc=self)
        if name == "1Table":
            return TableStream(bytes, self.params, name, doc=self)
        else:
            return DOCDirStream(bytes, self.params, name, doc=self)

class TableStream(DOCDirStream):
    def __init__(self, bytes, params, name, doc):
        DOCDirStream.__init__(self, bytes, params, name, doc = doc)

    def dump(self):
        print '<stream name="%s" size="%s"/>' % (self.name, self.size)

class WordDocumentStream(DOCDirStream):
    def __init__(self, bytes, params, doc):
        DOCDirStream.__init__(self, bytes, params, "WordDocument", doc = doc)

    def dump(self):
        print '<stream name="WordDocument" size="%d">' % self.size
        self.dumpFib()
        print '</stream>'

    def dumpFib(self):
        print '<fib>'
        self.dumpFibBase("base")
        self.printAndSet("csw", self.getuInt16())
        self.pos += 2
        self.dumpFibRgW97("fibRgW")
        self.printAndSet("cslw", self.getuInt16())
        self.pos += 2
        self.dumpFibRgLw97("fibRgLw")
        self.printAndSet("cbRgFcLcb", self.getuInt16())
        self.pos += 2
        self.dumpFibRgFcLcb("fibRgFcLcbBlob")
        self.printAndSet("cswNew", self.getuInt16())
        self.pos += 2
        print '</fib>'

    def dumpFibBase(self, name):
        print '<%s type="FibBase" size="32 bytes">' % name

        self.printAndSet("wIndent", self.getuInt16())
        self.pos += 2

        self.printAndSet("nFib", self.getuInt16())
        self.pos += 2

        self.printAndSet("unused", self.getuInt16())
        self.pos += 2

        self.printAndSet("lid", self.getuInt16())
        self.pos += 2

        self.printAndSet("pnNext", self.getuInt16())
        self.pos += 2

        buf = self.getuInt16()
        self.pos += 2
        self.printAndSet("fDot", self.getBit(buf, 0))
        self.printAndSet("fGlsy", self.getBit(buf, 1))
        self.printAndSet("fComplex", self.getBit(buf, 2))
        self.printAndSet("fHasPic", self.getBit(buf, 3))

        self.printAndSet("cQuickSaves", ((buf & (2**4-1 << 4)) >> 4), hexdump=False)

        self.printAndSet("fEncrypted", self.getBit(buf, 8))
        self.printAndSet("fWhichTblStm", self.getBit(buf, 9))
        self.printAndSet("fReadOnlyRecommended", self.getBit(buf, 10))
        self.printAndSet("fWriteReservation", self.getBit(buf, 11))

        self.printAndSet("fExtChar", self.getBit(buf, 12))
        self.printAndSet("fLoadOverride", self.getBit(buf, 13))
        self.printAndSet("fFarEast", self.getBit(buf, 14))
        self.printAndSet("fObfuscated", self.getBit(buf, 15))

        self.printAndSet("nFibBack", self.getuInt16())
        self.pos += 2

        self.printAndSet("lKey", self.getuInt32())
        self.pos += 4

        self.printAndSet("envr", self.getuInt8())
        self.pos += 1

        buf = self.getuInt8()
        self.pos += 1

        self.printAndSet("fMac", self.getBit(buf, 0))
        self.printAndSet("fEmptySpecial", self.getBit(buf, 1))
        self.printAndSet("fLoadOverridePage", self.getBit(buf, 2))
        self.printAndSet("reserved1", self.getBit(buf, 3))
        self.printAndSet("reserved2", self.getBit(buf, 4))
        self.printAndSet("fSpare0",  (buf & (2**3-1)))

        self.printAndSet("reserved3", self.getuInt16())
        self.pos += 2
        self.printAndSet("reserved4", self.getuInt16())
        self.pos += 2
        self.printAndSet("reserved5", self.getuInt32())
        self.pos += 4
        self.printAndSet("reserved6", self.getuInt32())
        self.pos += 4

        print '</%s>' % name

    def dumpFibRgW97(self, name):
        print '<%s type="FibRgW97" size="28 bytes">' % name

        for i in range(13):
            self.printAndSet("reserved%d" % (i + 1), self.getuInt16())
            self.pos += 2
        self.printAndSet("lidFE", self.getuInt16())
        self.pos += 2

        print '</%s>' % name

    def dumpFibRgLw97(self, name):
        print '<%s type="FibRgLw97" size="88 bytes">' % name

        fields = [
                "cbMac",
                "reserved1",
                "reserved2",
                "ccpText",
                "ccpFtn",
                "ccpHdd",
                "reserved3",
                "ccpAtn",
                "ccpEdn",
                "ccpTxbx",
                "ccpHdrTxbx",
                "reserved4",
                "reserved5",
                "reserved6",
                "reserved7",
                "reserved8",
                "reserved9",
                "reserved10",
                "reserved11",
                "reserved12",
                "reserved13",
                "reserved14",
                ]
        for i in fields:
            self.printAndSet(i, self.getuInt32())
            self.pos += 4

        print '</%s>' % name

    def dumpFibRgFcLcb(self, name):
        if self.nFib == 0x00c1:
            self.dumpFibRgFcLcb97(name)
        elif self.nFib == 0x0101:
            self.dumpFibRgFcLcb2002(name)
        else:
            print """<todo what="dumpFibRgFcLcb() doesn't know how to handle nFib = %s">""" % hex(self.nFib)

    def __dumpFibRgFcLcb97(self):
        # should be 186
        fields = [
            ["fcStshfOrig"],
            ["lcbStshfOrig"],
            ["fcStshf"],
            ["lcbStshf"],
            ["fcPlcffndRef"],
            ["lcbPlcffndRef"],
            ["fcPlcffndTxt"],
            ["lcbPlcffndTxt"],
            ["fcPlcfandRef"],
            ["lcbPlcfandRef"],
            ["fcPlcfandTxt"],
            ["lcbPlcfandTxt"],
            ["fcPlcfSed"],
            ["lcbPlcfSed"],
            ["fcPlcPad"],
            ["lcbPlcPad"],
            ["fcPlcfPhe"],
            ["lcbPlcfPhe"],
            ["fcSttbfGlsy"],
            ["lcbSttbfGlsy"],
            ["fcPlcfGlsy"],
            ["lcbPlcfGlsy"],
            ["fcPlcfHdd"],
            ["lcbPlcfHdd"],
            ["fcPlcfBteChpx"],
            ["lcbPlcfBteChpx", self.handleLcbPlcfBteChpx],
            ["fcPlcfBtePapx"],
            ["lcbPlcfBtePapx", self.handleLcbPlcfBtePapx],
            ["fcPlcfSea"],
            ["lcbPlcfSea"],
            ["fcSttbfFfn"],
            ["lcbSttbfFfn", self.handleLcbSttbfFfn],
            ["fcPlcfFldMom"],
            ["lcbPlcfFldMom"],
            ["fcPlcfFldHdr"],
            ["lcbPlcfFldHdr"],
            ["fcPlcfFldFtn"],
            ["lcbPlcfFldFtn"],
            ["fcPlcfFldAtn"],
            ["lcbPlcfFldAtn"],
            ["fcPlcfFldMcr"],
            ["lcbPlcfFldMcr"],
            ["fcSttbfBkmk"],
            ["lcbSttbfBkmk"],
            ["fcPlcfBkf"],
            ["lcbPlcfBkf"],
            ["fcPlcfBkl"],
            ["lcbPlcfBkl"],
            ["fcCmds"],
            ["lcbCmds"],
            ["fcUnused1"],
            ["lcbUnused1"],
            ["fcSttbfMcr"],
            ["lcbSttbfMcr"],
            ["fcPrDrvr"],
            ["lcbPrDrvr"],
            ["fcPrEnvPort"],
            ["lcbPrEnvPort"],
            ["fcPrEnvLand"],
            ["lcbPrEnvLand"],
            ["fcWss"],
            ["lcbWss"],
            ["fcDop"],
            ["lcbDop"],
            ["fcSttbfAssoc"],
            ["lcbSttbfAssoc"],
            ["fcClx"],
            ["lcbClx", self.handleLcbClx],
            ["fcPlcfPgdFtn"],
            ["lcbPlcfPgdFtn"],
            ["fcAutosaveSource"],
            ["lcbAutosaveSource"],
            ["fcGrpXstAtnOwners"],
            ["lcbGrpXstAtnOwners"],
            ["fcSttbfAtnBkmk"],
            ["lcbSttbfAtnBkmk"],
            ["fcUnused2"],
            ["lcbUnused2"],
            ["fcUnused3"],
            ["lcbUnused3"],
            ["fcPlcSpaMom"],
            ["lcbPlcSpaMom"],
            ["fcPlcSpaHdr"],
            ["lcbPlcSpaHdr"],
            ["fcPlcfAtnBkf"],
            ["lcbPlcfAtnBkf"],
            ["fcPlcfAtnBkl"],
            ["lcbPlcfAtnBkl"],
            ["fcPms"],
            ["lcbPms"],
            ["fcFormFldSttbs"],
            ["lcbFormFldSttbs"],
            ["fcPlcfendRef"],
            ["lcbPlcfendRef"],
            ["fcPlcfendTxt"],
            ["lcbPlcfendTxt"],
            ["fcPlcfFldEdn"],
            ["lcbPlcfFldEdn"],
            ["fcUnused4"],
            ["lcbUnused4"],
            ["fcDggInfo"],
            ["lcbDggInfo"],
            ["fcSttbfRMark"],
            ["lcbSttbfRMark"],
            ["fcSttbfCaption"],
            ["lcbSttbfCaption"],
            ["fcSttbfAutoCaption"],
            ["lcbSttbfAutoCaption"],
            ["fcPlcfWkb"],
            ["lcbPlcfWkb"],
            ["fcPlcfSpl"],
            ["lcbPlcfSpl"],
            ["fcPlcftxbxTxt"],
            ["lcbPlcftxbxTxt"],
            ["fcPlcfFldTxbx"],
            ["lcbPlcfFldTxbx"],
            ["fcPlcfHdrtxbxTxt"],
            ["lcbPlcfHdrtxbxTxt"],
            ["fcPlcffldHdrTxbx"],
            ["lcbPlcffldHdrTxbx"],
            ["fcStwUser"],
            ["lcbStwUser"],
            ["fcSttbTtmbd"],
            ["lcbSttbTtmbd"],
            ["fcCookieData"],
            ["lcbCookieData"],
            ["fcPgdMotherOldOld"],
            ["lcbPgdMotherOldOld"],
            ["fcBkdMotherOldOld"],
            ["lcbBkdMotherOldOld"],
            ["fcPgdFtnOldOld"],
            ["lcbPgdFtnOldOld"],
            ["fcBkdFtnOldOld"],
            ["lcbBkdFtnOldOld"],
            ["fcPgdEdnOldOld"],
            ["lcbPgdEdnOldOld"],
            ["fcBkdEdnOldOld"],
            ["lcbBkdEdnOldOld"],
            ["fcSttbfIntlFld"],
            ["lcbSttbfIntlFld"],
            ["fcRouteSlip"],
            ["lcbRouteSlip"],
            ["fcSttbSavedBy"],
            ["lcbSttbSavedBy"],
            ["fcSttbFnm"],
            ["lcbSttbFnm"],
            ["fcPlfLst"],
            ["lcbPlfLst"],
            ["fcPlfLfo"],
            ["lcbPlfLfo"],
            ["fcPlcfTxbxBkd"],
            ["lcbPlcfTxbxBkd"],
            ["fcPlcfTxbxHdrBkd"],
            ["lcbPlcfTxbxHdrBkd"],
            ["fcDocUndoWord9"],
            ["lcbDocUndoWord9"],
            ["fcRgbUse"],
            ["lcbRgbUse"],
            ["fcUsp"],
            ["lcbUsp"],
            ["fcUskf"],
            ["lcbUskf"],
            ["fcPlcupcRgbUse"],
            ["lcbPlcupcRgbUse"],
            ["fcPlcupcUsp"],
            ["lcbPlcupcUsp"],
            ["fcSttbGlsyStyle"],
            ["lcbSttbGlsyStyle"],
            ["fcPlgosl"],
            ["lcbPlgosl"],
            ["fcPlcocx"],
            ["lcbPlcocx"],
            ["fcPlcfBteLvc"],
            ["lcbPlcfBteLvc"],
            ["dwLowDateTime"],
            ["dwHighDateTime"],
            ["fcPlcfLvcPre10"],
            ["lcbPlcfLvcPre10"],
            ["fcPlcfAsumy"],
            ["lcbPlcfAsumy"],
            ["fcPlcfGram"],
            ["lcbPlcfGram"],
            ["fcSttbListNames"],
            ["lcbSttbListNames"],
            ["fcSttbfUssr"],
            ["lcbSttbfUssr"],
                ]
        for i in fields:
            self.printAndSet(i[0], self.getuInt32(), end = len(i) == 1)
            self.pos += 4
            if len(i) > 1:
                i[1]()
                print '</%s>' % i[0]

    def handleLcbClx(self):
        offset = self.fcClx
        size = self.lcbClx
        clx = docrecord.Clx(self.doc.getDirectoryStreamByName("1Table").bytes, self, offset, size)
        clx.dump()

    def handleLcbPlcfBteChpx(self):
        offset = self.fcPlcfBteChpx
        size = self.lcbPlcfBteChpx
        plcBteChpx = docrecord.PlcBteChpx(self.doc.getDirectoryStreamByName("1Table").bytes, self, offset, size)
        plcBteChpx.dump()

    def handleLcbPlcfBtePapx(self):
        offset = self.fcPlcfBtePapx
        size = self.lcbPlcfBtePapx
        plcBtePapx = docrecord.PlcBtePapx(self.doc.getDirectoryStreamByName("1Table").bytes, self, offset, size)
        plcBtePapx.dump()

    def handleLcbSttbfFfn(self):
        offset = self.fcSttbfFfn
        size = self.lcbSttbfFfn
        sttbfFfn = docrecord.SttbfFfn(self.doc.getDirectoryStreamByName("1Table").bytes, self, offset, size)
        sttbfFfn.dump()

    def dumpFibRgFcLcb97(self, name):
        print '<%s type="FibRgFcLcb97" size="744 bytes">' % name
        self.__dumpFibRgFcLcb97()
        print '</%s>' % name

    def __dumpFibRgFcLcb2000(self):
        self.__dumpFibRgFcLcb97()
        fields = [
            "fcPlcfTch", 
            "lcbPlcfTch", 
            "fcRmdThreading", 
            "lcbRmdThreading", 
            "fcMid", 
            "lcbMid", 
            "fcSttbRgtplc", 
            "lcbSttbRgtplc", 
            "fcMsoEnvelope", 
            "lcbMsoEnvelope", 
            "fcPlcfLad", 
            "lcbPlcfLad", 
            "fcRgDofr", 
            "lcbRgDofr", 
            "fcPlcosl", 
            "lcbPlcosl", 
            "fcPlcfCookieOld", 
            "lcbPlcfCookieOld", 
            "fcPgdMotherOld", 
            "lcbPgdMotherOld", 
            "fcBkdMotherOld", 
            "lcbBkdMotherOld", 
            "fcPgdFtnOld", 
            "lcbPgdFtnOld", 
            "fcBkdFtnOld", 
            "lcbBkdFtnOld", 
            "fcPgdEdnOld", 
            "lcbPgdEdnOld", 
            "fcBkdEdnOld", 
            "lcbBkdEdnOld", 
                ]
        for i in fields:
            self.printAndSet(i, self.getuInt32())
            self.pos += 4

    def __dumpFibRgFcLcb2002(self):
        self.__dumpFibRgFcLcb2000()
        fields = [
            "fcUnused1",
            "lcbUnused1",
            "fcPlcfPgp",
            "lcbPlcfPgp",
            "fcPlcfuim",
            "lcbPlcfuim",
            "fcPlfguidUim",
            "lcbPlfguidUim",
            "fcAtrdExtra",
            "lcbAtrdExtra",
            "fcPlrsid",
            "lcbPlrsid",
            "fcSttbfBkmkFactoid",
            "lcbSttbfBkmkFactoid",
            "fcPlcfBkfFactoid",
            "lcbPlcfBkfFactoid",
            "fcPlcfcookie",
            "lcbPlcfcookie",
            "fcPlcfBklFactoid",
            "lcbPlcfBklFactoid",
            "fcFactoidData",
            "lcbFactoidData",
            "fcDocUndo",
            "lcbDocUndo",
            "fcSttbfBkmkFcc",
            "lcbSttbfBkmkFcc",
            "fcPlcfBkfFcc",
            "lcbPlcfBkfFcc",
            "fcPlcfBklFcc",
            "lcbPlcfBklFcc",
            "fcSttbfbkmkBPRepairs",
            "lcbSttbfbkmkBPRepairs",
            "fcPlcfbkfBPRepairs",
            "lcbPlcfbkfBPRepairs",
            "fcPlcfbklBPRepairs",
            "lcbPlcfbklBPRepairs",
            "fcPmsNew",
            "lcbPmsNew",
            "fcODSO",
            "lcbODSO",
            "fcPlcfpmiOldXP",
            "lcbPlcfpmiOldXP",
            "fcPlcfpmiNewXP",
            "lcbPlcfpmiNewXP",
            "fcPlcfpmiMixedXP",
            "lcbPlcfpmiMixedXP",
            "fcUnused2",
            "lcbUnused2",
            "fcPlcffactoid",
            "lcbPlcffactoid",
            "fcPlcflvcOldXP",
            "lcbPlcflvcOldXP",
            "fcPlcflvcNewXP",
            "lcbPlcflvcNewXP",
            "fcPlcflvcMixedXP",
            "lcbPlcflvcMixedXP",
                ]
        for i in fields:
            self.printAndSet(i, self.getuInt32())
            self.pos += 4

    def dumpFibRgFcLcb2002(self, name):
        print '<%s type="dumpFibRgFcLcb2002" size="744 bytes">' % name
        self.__dumpFibRgFcLcb2002()
        print '</%s>' % name

# vim:set filetype=python shiftwidth=4 softtabstop=4 expandtab:
