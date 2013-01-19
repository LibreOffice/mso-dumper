#!/usr/bin/env python
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
#

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
        self.printAndSet("csw", self.readuInt16())
        self.dumpFibRgW97("fibRgW")
        self.printAndSet("cslw", self.readuInt16())
        self.dumpFibRgLw97("fibRgLw")
        self.printAndSet("cbRgFcLcb", self.readuInt16())

        self.blobOffset = self.pos
        self.nFibNew = self.__getFibNew()
        self.dumpFibRgFcLcb("fibRgFcLcbBlob")
        self.pos = self.__getCswNewOffset()

        self.printAndSet("cswNew", self.readuInt16(), offset = True)
        if self.cswNew != 0:
            self.dumpFibRgCswNew("fibRgCswNew")
        print '</fib>'

    def __getFibNew(self):
        cswNew = self.getuInt16(pos = self.__getCswNewOffset())
        if cswNew == 0:
            return 0
        else:
            return self.getuInt16(pos = self.__getCswNewOffset() + 2)
    
    def __getCswNewOffset(self):
        return self.blobOffset + (8 * self.cbRgFcLcb)

    def dumpFibRgCswNew(self, name):
        print '<%s type="FibRgCswNew" size="%d bytes">' % (name, self.cswNew)
        self.printAndSet("nFibNew", self.readuInt16())
        if self.nFibNew == 0x0112:
            self.dumpFibRgCswNewData2007("fibRgCswNewData2007")
        elif self.nFibNew == 0x00D9:
            self.dumpFibRgCswNewData2000("fibRgCswNewData2000")
        else:
            print """<todo what="dumpFibRgCswNew() doesn't know how to handle nFibNew = %s"/>""" % hex(self.nFibNew)
        print '</%s>' % name

    def __dumpFibRgCswNewData2000(self):
        self.printAndSet("cQuickSavesNew", self.readuInt16())

    def dumpFibRgCswNewData2000(self, name):
        print '<%s type="FibRgCswNewData2000" size="%d bytes">' % (name, 8)
        self.__dumpFibRgCswNewData2000()
        print '</%s>' % name

    def dumpFibRgCswNewData2007(self, name):
        print '<%s type="FibRgCswNewData2007" size="%d bytes">' % (name, 8)
        self.__dumpFibRgCswNewData2000()
        self.printAndSet("lidThemeOther", self.readuInt16())
        self.printAndSet("lidThemeFE", self.readuInt16())
        self.printAndSet("lidThemeCS", self.readuInt16())
        print '</%s>' % name

    def dumpFibBase(self, name):
        print '<%s type="FibBase" size="32 bytes">' % name

        self.printAndSet("wIdent", self.readuInt16())
        self.printAndSet("nFib", self.readuInt16())
        self.printAndSet("unused", self.readuInt16())
        self.printAndSet("lid", self.readuInt16())
        self.printAndSet("pnNext", self.readuInt16())

        buf = self.readuInt16()
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

        self.printAndSet("nFibBack", self.readuInt16())
        self.printAndSet("lKey", self.readuInt32())
        self.printAndSet("envr", self.readuInt8())

        buf = self.readuInt8()

        self.printAndSet("fMac", self.getBit(buf, 0))
        self.printAndSet("fEmptySpecial", self.getBit(buf, 1))
        self.printAndSet("fLoadOverridePage", self.getBit(buf, 2))
        self.printAndSet("reserved1", self.getBit(buf, 3))
        self.printAndSet("reserved2", self.getBit(buf, 4))
        self.printAndSet("fSpare0",  (buf & (2**3-1)))

        self.printAndSet("reserved3", self.readuInt16())
        self.printAndSet("reserved4", self.readuInt16())
        # reserved5 in the spec, offset of first character of text according to LO ww8 import filter
        self.printAndSet("fcMin", self.readuInt32())
        self.printAndSet("reserved6", self.readuInt32())

        print '</%s>' % name

    def dumpFibRgW97(self, name):
        print '<%s type="FibRgW97" size="28 bytes">' % name

        for i in range(13):
            self.printAndSet("reserved%d" % (i + 1), self.readuInt16())
        self.printAndSet("lidFE", self.readuInt16())

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
            self.printAndSet(i, self.readuInt32())

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
            ["lcbStshf", self.handleLcbStshf],
            ["fcPlcffndRef"],
            ["lcbPlcffndRef"],
            ["fcPlcffndTxt"],
            ["lcbPlcffndTxt"],
            ["fcPlcfandRef"],
            ["lcbPlcfandRef", self.handleLcbPlcfandRef],
            ["fcPlcfandTxt"],
            ["lcbPlcfandTxt", self.handleLcbPlcfandTxt],
            ["fcPlcfSed"],
            ["lcbPlcfSed", self.handleLcbPlcfSed],
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
            ["lcbPlcfFldMom", self.handleLcbPlcfFldMom],
            ["fcPlcfFldHdr"],
            ["lcbPlcfFldHdr"],
            ["fcPlcfFldFtn"],
            ["lcbPlcfFldFtn"],
            ["fcPlcfFldAtn"],
            ["lcbPlcfFldAtn"],
            ["fcPlcfFldMcr"],
            ["lcbPlcfFldMcr"],
            ["fcSttbfBkmk"],
            ["lcbSttbfBkmk", self.handleLcbSttbfBkmk],
            ["fcPlcfBkf"],
            ["lcbPlcfBkf", self.handleLcbPlcfBkf],
            ["fcPlcfBkl"],
            ["lcbPlcfBkl", self.handleLcbPlcfBkl],
            ["fcCmds"],
            ["lcbCmds", self.handleLcbCmds],
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
            ["lcbWss", self.handleLcbWss],
            ["fcDop"],
            ["lcbDop", self.handleDop],
            ["fcSttbfAssoc"],
            ["lcbSttbfAssoc", self.handleLcbSttbfAssoc],
            ["fcClx"],
            ["lcbClx", self.handleLcbClx],
            ["fcPlcfPgdFtn"],
            ["lcbPlcfPgdFtn"],
            ["fcAutosaveSource"],
            ["lcbAutosaveSource"],
            ["fcGrpXstAtnOwners"],
            ["lcbGrpXstAtnOwners", self.handleLcbGrpXstAtnOwners],
            ["fcSttbfAtnBkmk"],
            ["lcbSttbfAtnBkmk", self.handleLcbSttbfAtnBkmk],
            ["fcUnused2"],
            ["lcbUnused2"],
            ["fcUnused3"],
            ["lcbUnused3"],
            ["fcPlcSpaMom"],
            ["lcbPlcSpaMom"],
            ["fcPlcSpaHdr"],
            ["lcbPlcSpaHdr"],
            ["fcPlcfAtnBkf"],
            ["lcbPlcfAtnBkf", self.handleLcbPlcfAtnBkf],
            ["fcPlcfAtnBkl"],
            ["lcbPlcfAtnBkl", self.handleLcbPlcfAtnBkl],
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
            ["lcbDggInfo", self.handleLcbDggInfo],
            ["fcSttbfRMark"],
            ["lcbSttbfRMark", self.handleLcbSttbfRMark],
            ["fcSttbfCaption"],
            ["lcbSttbfCaption"],
            ["fcSttbfAutoCaption"],
            ["lcbSttbfAutoCaption"],
            ["fcPlcfWkb"],
            ["lcbPlcfWkb"],
            ["fcPlcfSpl"],
            ["lcbPlcfSpl", self.handleLcbPlcfSpl],
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
            ["lcbSttbSavedBy", self.handleLcbSttbSavedBy],
            ["fcSttbFnm"],
            ["lcbSttbFnm"],
            ["fcPlfLst"],
            ["lcbPlfLst", self.handleLcbPlfLst],
            ["fcPlfLfo"],
            ["lcbPlfLfo", self.handleLcbPlfLfo],
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
            ["lcbPlcfGram", self.handleLcbPlcfGram],
            ["fcSttbListNames"],
            ["lcbSttbListNames", self.handleLcbSttbListNames],
            ["fcSttbfUssr"],
            ["lcbSttbfUssr"],
                ]

        # Parse Clx early, as it's needed by other structures.
        posOrig = self.pos
        for i in fields:
            value = self.readInt32()
            if i[0] == "fcClx":
                self.printAndSet(i[0], value, silent = True)
            if i[0] == "lcbClx":
                self.printAndSet(i[0], value, silent = True)
                i[1](silent = True)
        self.pos = posOrig

        for i in fields:
            value = self.readInt32()
            hasHandler = len(i) > 1
            # the spec says these must be ignored
            needsIgnoring = ["lcbStshfOrig", "lcbPlcfBteLvc"]
            if self.ccpHdd == 0:
                needsIgnoring.append("lcbPlcfHdd")
            # a member needs handling if it defines the size of a struct and it's non-zero
            needsHandling = i[0].startswith("lcb") and value != 0 and (not i[0] in needsIgnoring)
            self.printAndSet(i[0], value, end = ((not hasHandler) and (not needsHandling)), offset = True)
            if hasHandler or needsHandling:
                if needsHandling:
                    if hasHandler:
                        i[1]()
                    else:
                        print '<todo what="value is non-zero and unhandled"/>'
                print '</%s>' % i[0]

    def handleDop(self):
        docrecord.Dop(self).dump()

    def handleLcbClx(self, silent = False):
        offset = self.fcClx
        size = self.lcbClx
        self.clx = docrecord.Clx(self.doc.getDirectoryStreamByName("1Table").bytes, self, offset, size)
        if not silent:
            self.clx.dump()

    def handleLcbPlcfBteChpx(self):
        plcBteChpx = docrecord.PlcBteChpx(self)
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

    def handleLcbStshf(self):
        offset = self.fcStshf
        size = self.lcbStshf
        stsh = docrecord.STSH(self.doc.getDirectoryStreamByName("1Table").bytes, self, offset, size)
        stsh.dump()

    def handleLcbPlcfandTxt(self):
        offset = self.fcPlcfandTxt
        size = self.lcbPlcfandTxt
        plcfandTxt = docrecord.PlcfandTxt(self, offset, size)
        plcfandTxt.dump()

    def handleLcbPlcfandRef(self):
        offset = self.fcPlcfandRef
        size = self.lcbPlcfandRef
        plcfandRef = docrecord.PlcfandRef(self, offset, size)
        plcfandRef.dump()

    def handleLcbGrpXstAtnOwners(self):
        docrecord.GrpXstAtnOwners(self).dump()

    def handleLcbSttbfAtnBkmk(self):
        offset = self.fcSttbfAtnBkmk
        size = self.lcbSttbfAtnBkmk
        sttbfAtnBkmk = docrecord.SttbfAtnBkmk(self, offset, size)
        sttbfAtnBkmk.dump()

    def handleLcbPlcfAtnBkf(self):
        offset = self.fcPlcfAtnBkf
        size = self.lcbPlcfAtnBkf
        self.plcfAtnBkf = docrecord.PlcfBkf(self, offset, size)
        self.plcfAtnBkf.dump()

    def handleLcbPlcfAtnBkl(self):
        offset = self.fcPlcfAtnBkl
        size = self.lcbPlcfAtnBkl
        plcfBkl = docrecord.PlcfBkl(self, offset, size, start = self.plcfAtnBkf)
        plcfBkl.dump()

    def handleLcbPlcfBkf(self):
        offset = self.fcPlcfBkf
        size = self.lcbPlcfBkf
        self.plcfBkf = docrecord.PlcfBkf(self, offset, size)
        self.plcfBkf.dump()

    def handleLcbPlcfBkl(self):
        offset = self.fcPlcfBkl
        size = self.lcbPlcfBkl
        plcfBkl = docrecord.PlcfBkl(self, offset, size, start = self.plcfBkf)
        plcfBkl.dump()

    def handleLcbPlcfSed(self):
        offset = self.fcPlcfSed
        size = self.lcbPlcfSed
        plcfSed = docrecord.PlcfSed(self, offset, size)
        plcfSed.dump()

    def handleLcbCmds(self):
        offset = self.fcCmds
        size = self.lcbCmds
        tcg = docrecord.Tcg(self, offset, size)
        tcg.dump()

    def handleLcbWss(self):
        docrecord.Selsf(self).dump()

    def handleLcbSttbfAssoc(self):
        docrecord.SttbfAssoc(self).dump()

    def handleLcbSttbfRMark(self):
        docrecord.SttbfRMark(self).dump()

    def handleLcbDggInfo(self):
        docrecord.OfficeArtContent(self).dump()

    def handleLcbPlcfFldMom(self):
        docrecord.PlcFld(self).dump()

    def handleLcbPlcfSpl(self):
        docrecord.PlcfSpl(self).dump()

    def handleLcbPlcfGram(self):
        docrecord.PlcfGram(self).dump()

    def handleLcbSttbSavedBy(self):
        docrecord.SttbSavedBy(self).dump()

    def handleLcbPlfLst(self):
        docrecord.PlfLst(self).dump()

    def handleLcbPlfLfo(self):
        docrecord.PlfLfo(self).dump()

    def handleLcbSttbListNames(self):
        docrecord.SttbListNames(self).dump()

    def handleLcbSttbfBkmk(self):
        docrecord.SttbfBkmk(self).dump()

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
            self.printAndSet(i, self.readuInt32())

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
            self.printAndSet(i, self.readuInt32())

    def dumpFibRgFcLcb2002(self, name):
        print '<%s type="FibRgFcLcb2002" size="744 bytes">' % name
        self.__dumpFibRgFcLcb2002()
        print '</%s>' % name

    def __findText(self, plcPcd, cp):
        """Find the largest i such that plcPcd.aCp[i] <= cp."""
        for i in range(len(plcPcd.aCp)):
            if plcPcd.aCp[i] <= cp:
                index = i
        return index

    def retrieveText(self, start, end, logicalLength = False):
        plcPcd = self.clx.pcdt.plcPcd
        idx = self.__findText(plcPcd, start)
        return plcPcd.aPcd[idx].fc.getTransformedValue(start, end, logicalPositions = False, logicalLength = logicalLength)

# vim:set filetype=python shiftwidth=4 softtabstop=4 expandtab:
