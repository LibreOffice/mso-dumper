#!/usr/bin/env python
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
#

import struct
import globals
from docdirstream import DOCDirStream
import docsprm

class FcCompressed(DOCDirStream):
    """The FcCompressed structure specifies the location of text in the WordDocument Stream."""
    def __init__(self, bytes, mainStream, offset, size):
        DOCDirStream.__init__(self, bytes, mainStream=mainStream)
        self.pos = offset
        self.size = size

    def dump(self):
        print '<fcCompressed type="FcCompressed" offset="%d" size="%d bytes">' % (self.pos, self.size)
        buf = self.readuInt32()
        self.printAndSet("fc", buf & ((2**32-1) >> 2)) # bits 0..29
        self.printAndSet("fCompressed", self.getBit(buf, 30))
        self.printAndSet("r1", self.getBit(buf, 31))
        print '</fcCompressed>'

    def getTransformedValue(self, start, end):
            if self.fCompressed:
                offset = self.fc/2
                return globals.encodeName(self.mainStream.bytes[offset:offset+end-start])
            else:
                offset = self.fc
                return globals.encodeName(self.mainStream.bytes[offset:offset+end*2-start].decode('utf-16'), lowOnly = True)

    @staticmethod
    def getFCTransformedValue(bytes, start, end):
        return globals.encodeName(bytes[start:end])

class Pcd(DOCDirStream):
    """The Pcd structure specifies the location of text in the WordDocument Stream and additional properties for this text."""
    def __init__(self, bytes, mainStream, offset, size):
        DOCDirStream.__init__(self, bytes, mainStream=mainStream)
        self.pos = offset
        self.size = size

    def dump(self):
        print '<pcd type="Pcd" offset="%d" size="%d bytes">' % (self.pos, self.size)
        buf = self.readuInt16()
        self.printAndSet("fNoParaLast", self.getBit(buf, 0))
        self.printAndSet("fR1", self.getBit(buf, 1))
        self.printAndSet("fDirty", self.getBit(buf, 2))
        self.printAndSet("fR2", buf & (2**13-1))
        self.fc = FcCompressed(self.bytes, self.mainStream, self.pos, 4)
        self.fc.dump()
        self.pos += 4
        print '</pcd>'

class PLC:
    """The PLC structure is an array of character positions followed by an array of data elements."""
    def __init__(self, totalSize, structSize):
        self.totalSize = totalSize
        self.structSize = structSize

    def getElements(self):
        return (self.totalSize - 4) / (4 + self.structSize) # defined by 2.2.2

    def getOffset(self, pos, i):
        return self.getPLCOffset(pos, self.getElements(), self.structSize, i)

    @staticmethod
    def getPLCOffset(pos, elements, structSize, i):
        return pos + (4 * (elements + 1)) + (structSize * i)

class BKC(DOCDirStream):
    """The BKC structure contains information about how a bookmark interacts with tables."""
    def __init__(self, bkc):
        self.bkc = bkc

    def dump(self):
        print '<bkc type="BKC">'
        self.printAndSet("itcFirst", self.bkc & 0x007f) # 1..7th bits
        self.printAndSet("fPub", self.getBit(self.bkc, 8))
        self.printAndSet("itcLim", (self.bkc & 0x3f00) >> 8) # 9..14th bits
        self.printAndSet("fNative", self.getBit(self.bkc, 15))
        self.printAndSet("fCol", self.getBit(self.bkc, 16))
        print '</bkc>'

class FBKF(DOCDirStream):
    """The FBKF structure contains information about a bookmark."""
    def __init__(self, plcfAtnBkf, offset):
        DOCDirStream.__init__(self, plcfAtnBkf.bytes)
        self.pos = offset

    def dump(self):
        print '<aFBKF type="FBKF" offset="%d">' % self.pos
        self.printAndSet("ibkl", self.readuInt16())
        BKC(self.readuInt16()).dump()
        print '</aFBKF>'

class PlcfBkf(DOCDirStream, PLC):
    """A PLCFBKF is a PLC whose data elements are FBKF structures."""
    def __init__(self, mainStream, offset, size):
        DOCDirStream.__init__(self, mainStream.doc.getDirectoryStreamByName("1Table").bytes, mainStream = mainStream)
        PLC.__init__(self, size, 4) # 4 is defined by 2.8.10
        self.pos = offset
        self.size = size

    def dump(self):
        print '<plcfBkf type="PlcfAtnBkf" offset="%d" size="%d bytes">' % (self.pos, self.size)
        offset = self.mainStream.fcMin
        pos = self.pos
        for i in range(self.getElements()):
            # aCp
            start = offset + self.getuInt32(pos = pos)
            print '<aCP index="%d" bookmarkStart="%d">' % (i, start)
            print '<transformed value="%s"/>' % FcCompressed.getFCTransformedValue(self.mainStream.bytes, start, start + 1)
            pos += 4

            # aFBKF
            aFBKF = FBKF(self, self.getOffset(self.pos, i))
            aFBKF.dump()
            print '</aCP>'
        print '</plcfBkf>'

class PlcfBkf(DOCDirStream, PLC):
    """A PLCFBKF is a PLC whose data elements are FBKF structures."""
    def __init__(self, mainStream, offset, size):
        DOCDirStream.__init__(self, mainStream.doc.getDirectoryStreamByName("1Table").bytes, mainStream = mainStream)
        PLC.__init__(self, size, 4) # 4 is defined by 2.8.10
        self.pos = offset
        self.size = size

    def dump(self):
        print '<plcfBkf type="PlcfBkf" offset="%d" size="%d bytes">' % (self.pos, self.size)
        offset = self.mainStream.fcMin
        pos = self.pos
        for i in range(self.getElements()):
            # aCp
            start = offset + self.getuInt32(pos = pos)
            print '<aCP index="%d" bookmarkStart="%d">' % (i, start)
            print '<transformed value="%s"/>' % FcCompressed.getFCTransformedValue(self.mainStream.bytes, start, start + 1)
            pos += 4

            # aFBKF
            aFBKF = FBKF(self, self.getOffset(self.pos, i))
            aFBKF.dump()
            print '</aCP>'
        print '</plcfBkf>'

class PlcfBkl(DOCDirStream, PLC):
    """The Plcfbkl structure is a PLC that contains only CPs and no additional data."""
    def __init__(self, mainStream, offset, size):
        DOCDirStream.__init__(self, mainStream.doc.getDirectoryStreamByName("1Table").bytes, mainStream = mainStream)
        PLC.__init__(self, size, 0) # 0 is defined by 2.8.12
        self.pos = offset
        self.size = size

    def dump(self):
        print '<plcfBkl type="PlcfBkl" offset="%d" size="%d bytes">' % (self.pos, self.size)
        offset = self.mainStream.fcMin
        pos = self.pos
        for i in range(self.getElements()):
            # aCp
            end = offset + self.getuInt32(pos = pos)
            print '<aCP index="%d" bookmarkEnd="%d">' % (i, end)
            print '<transformed value="%s"/>' % FcCompressed.getFCTransformedValue(self.mainStream.bytes, end, end + 1)
            pos += 4
            print '</aCP>'
        print '</plcfBkl>'

class PlcPcd(DOCDirStream, PLC):
    """The PlcPcd structure is a PLC whose data elements are Pcds (8 bytes each)."""
    def __init__(self, bytes, mainStream, offset, size):
        DOCDirStream.__init__(self, bytes, mainStream=mainStream)
        PLC.__init__(self, size, 8) # 8 is defined by 2.8.35
        self.pos = offset
        self.size = size

    def dump(self):
        print '<plcPcd type="PlcPcd" offset="%d" size="%d bytes">' % (self.pos, self.size)
        pos = self.pos
        for i in range(self.getElements()):
            # aCp
            start = self.getuInt32(pos = pos)
            end = self.getuInt32(pos = pos + 4)
            print '<aCP index="%d" start="%d" end="%d">' % (i, start, end)
            pos += 4

            # aPcd
            aPcd = Pcd(self.bytes, self.mainStream, self.getOffset(self.pos, i), 8)
            aPcd.dump()

            print '<transformed value="%s"/>' % aPcd.fc.getTransformedValue(start, end)
            print '</aCP>'
        print '</plcPcd>'

class Sprm(DOCDirStream):
    """The Sprm structure specifies a modification to a property of a character, paragraph, table, or section."""
    def __init__(self, bytes, offset):
        DOCDirStream.__init__(self, bytes)
        self.pos = offset
        self.operandSizeMap = {
                0: 1,
                1: 1,
                2: 2,
                3: 4,
                4: 2,
                5: 2,
                7: 3,
                }

        self.sprm = self.readuInt16()

        self.ispmd = (self.sprm & 0x1ff)        # 1-9th bits
        self.fSpec = (self.sprm & 0x200)  >> 9  # 10th bit
        self.sgc   = (self.sprm & 0x1c00) >> 10 # 11-13th bits
        self.spra  = (self.sprm & 0xe000) >> 13 # 14-16th bits

        if self.getOperandSize() == 1:
            self.operand = self.getuInt8()
        elif self.getOperandSize() == 2:
            self.operand = self.getuInt16()
        elif self.getOperandSize() == 3:
            self.operand = self.getuInt32() & 0x0fff
        elif self.getOperandSize() == 4:
            self.operand = self.getuInt32()
        elif self.getOperandSize() == 7:
            self.operand = self.getuInt64() & 0x0fffffff
        else:
            self.operand = "todo"

    def dump(self):
        sgcmap = {
                1: 'paragraph',
                2: 'character',
                3: 'picture',
                4: 'section',
                5: 'table'
                }
        nameMap = {
                1: docsprm.parMap,
                2: docsprm.chrMap,
                5: docsprm.tblMap,
                }
        print '<sprm value="%s" name="%s" ispmd="%s" fSpec="%s" sgc="%s" spra="%s" operandSize="%s" operand="%s"/>' % (
                hex(self.sprm), nameMap[self.sgc][self.sprm], hex(self.ispmd), hex(self.fSpec), sgcmap[self.sgc], hex(self.spra), self.getOperandSize(), hex(self.operand)
                )

    def getOperandSize(self):
        if self.spra == 6: # variable
            if self.sprm == 0xd634:
                return 7
            raise Exception()
        return self.operandSizeMap[self.spra]

class Prl(DOCDirStream):
    """The Prl structure is a Sprm that is followed by an operand."""
    def __init__(self, bytes, offset):
        DOCDirStream.__init__(self, bytes)
        self.pos = offset

    def dump(self):
        print '<prl type="Prl" offset="%d">' % self.pos
        self.sprm = Sprm(self.bytes, self.pos)
        self.pos += 2
        self.sprm.dump()
        print '</prl>'

    def getSize(self):
        return 2 + self.sprm.getOperandSize()

class GrpPrlAndIstd(DOCDirStream):
    """The GrpPrlAndIstd structure specifies the style and properties that are applied to a paragraph, a table row, or a table cell."""
    def __init__(self, bytes, offset, size):
        DOCDirStream.__init__(self, bytes)
        self.pos = offset
        self.size = size

    def dump(self):
        print '<grpPrlAndIstd type="GrpPrlAndIstd" offset="%d" size="%d bytes">' % (self.pos, self.size)
        pos = self.pos
        self.printAndSet("istd", self.getuInt16())
        pos += 2
        while (self.size - (pos - self.pos)) > 0:
            prl = Prl(self.bytes, pos)
            prl.dump()
            pos += prl.getSize()
        print '</grpPrlAndIstd>'

class Chpx(DOCDirStream):
    """The Chpx structure specifies a set of properties for text."""
    def __init__(self, bytes, mainStream, offset):
        DOCDirStream.__init__(self, bytes)
        self.pos = offset

    def dump(self):
        print '<chpx type="Chpx" offset="%d">' % self.pos
        self.printAndSet("cb", self.readuInt8())
        pos = self.pos
        while (self.cb - (pos - self.pos)) > 0:
            prl = Prl(self.bytes, pos)
            prl.dump()
            pos += prl.getSize()
        print '</chpx>'
    
class PapxInFkp(DOCDirStream):
    """The PapxInFkp structure specifies a set of text properties."""
    def __init__(self, bytes, mainStream, offset):
        DOCDirStream.__init__(self, bytes)
        self.pos = offset

    def dump(self):
        print '<papxInFkp type="PapxInFkp" offset="%d">' % self.pos
        self.printAndSet("cb", self.readuInt8())
        if self.cb == 0:
            self.printAndSet("cb_", self.readuInt8())
            grpPrlAndIstd = GrpPrlAndIstd(self.bytes, self.pos, 2 * self.cb_)
            grpPrlAndIstd.dump()
        else:
            print '<todo what="PapxInFkp::dump() first byte is not 0"/>'
        print '</papxInFkp>'
    
class BxPap(DOCDirStream):
    """The BxPap structure specifies the offset of a PapxInFkp in PapxFkp."""
    size = 13 # in bytes, see 2.9.23
    def __init__(self, bytes, mainStream, offset, parentoffset):
        DOCDirStream.__init__(self, bytes)
        self.pos = offset
        self.parentpos = parentoffset

    def dump(self):
        print '<bxPap type="BxPap" offset="%d" size="%d bytes">' % (self.pos, self.size)
        self.printAndSet("bOffset", self.readuInt8())
        papxInFkp = PapxInFkp(self.bytes, self.mainStream, self.parentpos + self.bOffset*2)
        papxInFkp.dump()
        print '</bxPap>'

class ChpxFkp(DOCDirStream):
    """The ChpxFkp structure maps text to its character properties."""
    def __init__(self, bytes, mainStream, offset, size):
        DOCDirStream.__init__(self, mainStream.bytes)
        self.pos = offset
        self.size = size

    def dump(self):
        print '<chpxFkp type="ChpxFkp" offset="%d" size="%d bytes">' % (self.pos, self.size)
        self.crun = self.getuInt8(pos = self.pos + self.size - 1)
        pos = self.pos
        for i in range(self.crun):
            # rgfc
            start = self.getuInt32(pos = pos)
            end = self.getuInt32(pos = pos + 4)
            print '<rgfc index="%d" start="%d" end="%d">' % (i, start, end)
            print '<transformed value="%s"/>' % FcCompressed.getFCTransformedValue(self.bytes, start, end)
            pos += 4

            # rgbx
            offset = PLC.getPLCOffset(self.pos, self.crun, 1, i)
            chpxOffset = self.getuInt8(pos = offset) * 2
            chpx = Chpx(self.bytes, self.mainStream, self.pos + chpxOffset)
            chpx.dump()
            print '</rgfc>'

        self.printAndSet("crun", self.crun)
        print '</chpxFkp>'

class PapxFkp(DOCDirStream):
    """The PapxFkp structure maps paragraphs, table rows, and table cells to their properties."""
    def __init__(self, bytes, mainStream, offset, size):
        DOCDirStream.__init__(self, mainStream.bytes)
        self.pos = offset
        self.size = size

    def dump(self):
        print '<papxFkp type="PapxFkp" offset="%d" size="%d bytes">' % (self.pos, self.size)
        self.cpara = self.getuInt8(pos = self.pos + self.size - 1)
        pos = self.pos
        for i in range(self.cpara):
            # rgfc
            start = self.getuInt32(pos = pos)
            end = self.getuInt32(pos = pos + 4)
            print '<rgfc index="%d" start="%d" end="%d">' % (i, start, end)
            print '<transformed value="%s"/>' % FcCompressed.getFCTransformedValue(self.bytes, start, end)
            pos += 4

            # rgbx
            offset = PLC.getPLCOffset(self.pos, self.cpara, BxPap.size, i)
            bxPap = BxPap(self.bytes, self.mainStream, offset, self.pos)
            bxPap.dump()
            print '</rgfc>'

        self.printAndSet("cpara", self.cpara)
        print '</papxFkp>'

class PnFkpChpx(DOCDirStream):
    """The PnFkpChpx structure specifies the location in the WordDocument Stream of a ChpxFkp structure."""
    def __init__(self, bytes, mainStream, offset, size, name):
        DOCDirStream.__init__(self, bytes, mainStream=mainStream)
        self.pos = offset
        self.size = size
        self.name = name

    def dump(self):
        print '<%s type="PnFkpChpx" offset="%d" size="%d bytes">' % (self.name, self.pos, self.size)
        buf = self.readuInt32()
        self.printAndSet("pn", buf & (2**22-1))
        chpxFkp = ChpxFkp(self.bytes, self.mainStream, self.pn*512, 512)
        chpxFkp.dump()
        print '</%s>' % self.name

class LPXCharBuffer9(DOCDirStream):
    """The LPXCharBuffer9 structure is a length-prefixed buffer for up to 9 Unicode characters."""
    def __init__(self, parent, name):
        DOCDirStream.__init__(self, parent.bytes)
        self.pos = parent.pos
        self.name = name

    def dump(self):
        print '<%s type="LPXCharBuffer9" offset="%d" size="20 bytes">' % (self.name, self.pos)
        self.printAndSet("cch", self.readuInt16())
        self.printAndSet("xcharArray", self.bytes[self.pos:self.pos+(self.cch*2)].decode('utf-16'), hexdump = False)
        print '</%s>' % self.name

class ATRDPre10(DOCDirStream):
    """The ATRDPre10 structure contains information about a comment in the document."""
    def __init__(self, aPlcfandRef, offset):
        DOCDirStream.__init__(self, aPlcfandRef.bytes)
        self.pos = offset

    def dump(self):
        print '<aATRDPre10 type="ATRDPre10" offset="%d" size="30 bytes">' % self.pos
        xstUsrInitl = LPXCharBuffer9(self, "xstUsrInitl")
        xstUsrInitl.dump()
        self.pos += 20
        self.printAndSet("ibst", self.readuInt16())
        self.printAndSet("bitsNotUsed", self.readuInt16())
        self.printAndSet("grfNotUsed", self.readuInt16())
        self.printAndSet("ITagBkmk", self.readInt32())
        print '</aATRDPre10>'

class PnFkpPapx(DOCDirStream):
    """The PnFkpPapx structure specifies the offset of a PapxFkp in the WordDocument Stream."""
    def __init__(self, bytes, mainStream, offset, size, name):
        DOCDirStream.__init__(self, bytes, mainStream=mainStream)
        self.pos = offset
        self.size = size
        self.name = name

    def dump(self):
        print '<%s type="PnFkpPapx" offset="%d" size="%d bytes">' % (self.name, self.pos, self.size)
        buf = self.readuInt32()
        self.printAndSet("pn", buf & (2**22-1))
        papxFkp = PapxFkp(self.bytes, self.mainStream, self.pn*512, 512)
        papxFkp.dump()
        print '</%s>' % self.name

class PlcBteChpx(DOCDirStream, PLC):
    """The PlcBteChpx structure is a PLC that maps the offsets of text in the WordDocument stream to the character properties of that text."""
    def __init__(self, bytes, mainStream, offset, size):
        DOCDirStream.__init__(self, bytes, mainStream=mainStream)
        PLC.__init__(self, size, 4)
        self.pos = offset
        self.size = size

    def dump(self):
        print '<plcBteChpx type="PlcBteChpx" offset="%d" size="%d bytes">' % (self.pos, self.size)
        pos = self.pos
        for i in range(self.getElements()):
            # aFC
            start = self.getuInt32(pos = pos)
            end = self.getuInt32(pos = pos + 4)
            print '<aFC index="%d" start="%d" end="%d">' % (i, start, end)
            pos += 4

            # aPnBteChpx
            aPnBteChpx = PnFkpChpx(self.bytes, self.mainStream, self.getOffset(self.pos, i), 4, "aPnBteChpx")
            aPnBteChpx.dump()
            print '</aFC>'
        print '</plcBteChpx>'

class PlcfandTxt(DOCDirStream, PLC):
    """The PlcfandTxt structure is a PLC that contains only CPs and no additional data."""
    def __init__(self, mainStream, offset, size):
        DOCDirStream.__init__(self, mainStream.doc.getDirectoryStreamByName("1Table").bytes, mainStream=mainStream)
        PLC.__init__(self, size, 0)
        self.pos = offset
        self.size = size

    def dump(self):
        print '<plcfandTxt type="PlcfandTxt" offset="%d" size="%d bytes">' % (self.pos, self.size)
        offset = self.mainStream.fcMin + self.mainStream.ccpText + self.mainStream.ccpFtn + self.mainStream.ccpHdd # TODO do this in a better way when headers are handled
        pos = self.pos
        for i in range(self.getElements() - 1):
            start = self.getuInt32(pos = pos)
            end = self.getuInt32(pos = pos + 4)
            print '<aCP index="%d" start="%d" end="%d">' % (i, start, end)
            print '<transformed value="%s"/>' % FcCompressed.getFCTransformedValue(self.mainStream.bytes, offset+start, offset+end)
            pos += 4
            print '</aCP>'
        print '</plcfandTxt>'

class PlcfandRef(DOCDirStream, PLC):
    """The PlcfandRef structure is a PLC whose data elements are ATRDPre10 structures."""
    def __init__(self, mainStream, offset, size):
        DOCDirStream.__init__(self, mainStream.doc.getDirectoryStreamByName("1Table").bytes, mainStream=mainStream)
        PLC.__init__(self, size, 30)
        self.pos = offset
        self.size = size

    def dump(self):
        print '<plcfandRef type="PlcfandRef" offset="%d" size="%d bytes">' % (self.pos, self.size)
        offset = self.mainStream.fcMin
        pos = self.pos
        for i in range(self.getElements()):
            start = offset + self.getuInt32(pos = pos)
            print '<aCP index="%d" commentEndOffset="%d">' % (i, start)
            print '<transformed value="%s"/>' % FcCompressed.getFCTransformedValue(self.mainStream.bytes, start, start + 1)
            pos += 4

            # aATRDPre10
            aATRDPre10 = ATRDPre10(self, self.getOffset(self.pos, i))
            aATRDPre10.dump()
            print '</aCP>'
        print '</plcfandRef>'

class PlcBtePapx(DOCDirStream, PLC):
    """The PlcBtePapx structure is a PLC that specifies paragraph, table row, or table cell properties."""
    def __init__(self, bytes, mainStream, offset, size):
        DOCDirStream.__init__(self, bytes, mainStream=mainStream)
        PLC.__init__(self, size, 4)
        self.pos = offset
        self.size = size

    def dump(self):
        print '<plcBtePapx type="PlcBtePapx" offset="%d" size="%d bytes">' % (self.pos, self.size)
        pos = self.pos
        for i in range(self.getElements()):
            # aFC
            start = self.getuInt32(pos = pos)
            end = self.getuInt32(pos = pos + 4)
            print '<aFC index="%d" start="%d" end="%d">' % (i, start, end)
            pos += 4

            # aPnBtePapx
            aPnBtePapx = PnFkpPapx(self.bytes, self.mainStream, self.getOffset(self.pos, i), 4, "aPnBtePapx")
            aPnBtePapx.dump()
            print '</aFC>'
        print '</plcBtePapx>'

class Pcdt(DOCDirStream):
    """The Pcdt structure contains a PlcPcd structure and specifies its size."""
    def __init__(self, bytes, mainStream, offset, size):
        DOCDirStream.__init__(self, bytes, mainStream=mainStream)
        self.pos = offset
        self.size = size

    def dump(self):
        print '<pcdt type="Pcdt" offset="%d" size="%d bytes">' % (self.pos, self.size)
        self.printAndSet("clxt", self.readuInt8())
        self.printAndSet("lcb", self.readuInt32())
        PlcPcd(self.bytes, self.mainStream, self.pos, self.lcb).dump()
        print '</pcdt>'

class Clx(DOCDirStream):
    def __init__(self, bytes, mainStream, offset, size):
        DOCDirStream.__init__(self, bytes, mainStream=mainStream)
        self.pos = offset
        self.size = size

    def dump(self):
        print '<clx type="Clx" offset="%d" size="%d bytes">' % (self.pos, self.size)
        firstByte = self.getuInt8()
        if firstByte == 0x02:
            print '<info what="Array of Prc, 0 elements"/>'
            Pcdt(self.bytes, self.mainStream, self.pos, self.size).dump()
        else:
            print '<todo what="Clx::dump() first byte is not 0x02"/>'
        print '</clx>'

class Copts60(DOCDirStream):
    """The Copts60 structure specifies compatibility options."""
    def __init__(self, dop):
        DOCDirStream.__init__(self, dop.bytes)
        self.pos = dop.pos

    def dump(self):
        print '<copts60 type="Copts60" offset="%s" size="2 bytes">' % self.pos
        # Copts60 first byte
        buf = self.readuInt8()
        self.printAndSet("fNoTabForInd", self.getBit(buf, 0))
        self.printAndSet("fNoSpaceRaiseLower", self.getBit(buf, 1))
        self.printAndSet("fSuppressSpBfAfterPgBrk", self.getBit(buf, 2))
        self.printAndSet("fWrapTrailSpaces", self.getBit(buf, 3))
        self.printAndSet("fMapPrintTextColor", self.getBit(buf, 4))
        self.printAndSet("fNoColumnBalance", self.getBit(buf, 5))
        self.printAndSet("fConvMailMergeEsc", self.getBit(buf, 6))
        self.printAndSet("fSuppressTopSpacing", self.getBit(buf, 7))

        # Copts60 second byte
        buf = self.readuInt8()
        self.printAndSet("fOrigWordTableRules", self.getBit(buf, 0))
        self.printAndSet("unused14", self.getBit(buf, 1))
        self.printAndSet("fShowBreaksInFrames", self.getBit(buf, 2))
        self.printAndSet("fSwapBordersFacingPgs", self.getBit(buf, 3))
        self.printAndSet("fLeaveBackslashAlone", self.getBit(buf, 4))
        self.printAndSet("fExpShRtn", self.getBit(buf, 5))
        self.printAndSet("fDntULTrlSpc", self.getBit(buf, 6))
        self.printAndSet("fDntBlnSbDbWid", self.getBit(buf, 7))
        print '</copts60>'

class DopBase(DOCDirStream):
    """The DopBase structure contains document and compatibility settings."""
    def __init__(self, dop):
        DOCDirStream.__init__(self, dop.bytes)
        self.pos = dop.pos
        self.dop = dop

    def dump(self):
        print '<dopBase offset="%d" size="%d bytes">' % (self.pos, 84)
        buf = self.readuInt8()
        self.printAndSet("fFacingPages", self.getBit(buf, 0))
        self.printAndSet("unused1", self.getBit(buf, 1))
        self.printAndSet("fPMHMainDoc", self.getBit(buf, 2))
        self.printAndSet("unused2", (buf & 0x18) >> 3) # 4..5th bits
        self.printAndSet("fpc", (buf & 0x60) >> 5) # 6..7th bits
        self.printAndSet("unused3", self.getBit(buf, 7))

        self.printAndSet("unused4", self.readuInt8())

        buf = self.readuInt16()
        self.printAndSet("rncFtn", buf & 0x03) # 1..2nd bits
        self.printAndSet("nFtn", (buf & 0xfffc) >> 2) # 3..16th bits

        buf = self.readuInt8()
        self.printAndSet("unused5", self.getBit(buf, 0))
        self.printAndSet("unused6", self.getBit(buf, 1))
        self.printAndSet("unused7", self.getBit(buf, 2))
        self.printAndSet("unused8", self.getBit(buf, 3))
        self.printAndSet("unused9", self.getBit(buf, 4))
        self.printAndSet("unused10", self.getBit(buf, 5))
        self.printAndSet("fSplAllDone", self.getBit(buf, 6))
        self.printAndSet("fSplAllClean", self.getBit(buf, 7))

        buf = self.readuInt8()
        self.printAndSet("fSplHideErrors", self.getBit(buf, 0))
        self.printAndSet("fGramHideErrors", self.getBit(buf, 1))
        self.printAndSet("fLabelDoc", self.getBit(buf, 2))
        self.printAndSet("fHyphCapitals", self.getBit(buf, 3))
        self.printAndSet("fAutoHyphen", self.getBit(buf, 4))
        self.printAndSet("fFormNoFields", self.getBit(buf, 5))
        self.printAndSet("fLinkStyles", self.getBit(buf, 6))
        self.printAndSet("fRevMarking", self.getBit(buf, 7))

        buf = self.readuInt8()
        self.printAndSet("unused11", self.getBit(buf, 0))
        self.printAndSet("fExactCWords", self.getBit(buf, 1))
        self.printAndSet("fPagHidden", self.getBit(buf, 2))
        self.printAndSet("fPagResults", self.getBit(buf, 3))
        self.printAndSet("fLockAtn", self.getBit(buf, 4))
        self.printAndSet("fMirrorMargins", self.getBit(buf, 5))
        self.printAndSet("fWord97Compat", self.getBit(buf, 6))
        self.printAndSet("unused12", self.getBit(buf, 7))

        buf = self.readuInt8()
        self.printAndSet("unused13", self.getBit(buf, 0))
        self.printAndSet("fProtEnabled", self.getBit(buf, 1))
        self.printAndSet("fDispFormFldSel", self.getBit(buf, 2))
        self.printAndSet("fRMView", self.getBit(buf, 3))
        self.printAndSet("fRMPrint", self.getBit(buf, 4))
        self.printAndSet("fLockVbaProj", self.getBit(buf, 5))
        self.printAndSet("fLockRev", self.getBit(buf, 6))
        self.printAndSet("fEmbedFonts", self.getBit(buf, 7))

        copts60 = Copts60(self)
        copts60.dump()
        self.pos += 2

        self.printAndSet("dxaTab", self.readuInt16())
        self.printAndSet("cpgWebOpt", self.readuInt16())
        self.printAndSet("dxaHotZ", self.readuInt16())
        self.printAndSet("cConsecHypLim", self.readuInt16())
        self.printAndSet("wSpare2", self.readuInt16())
        self.printAndSet("dttmCreated", self.readuInt32()) # TODO extract DTTM here and below
        self.printAndSet("dttmRevised", self.readuInt32())
        self.printAndSet("dttmLastPrint", self.readuInt32())
        self.printAndSet("nRevision", self.readInt16())
        self.printAndSet("tmEdited", self.readInt32())
        self.printAndSet("cWords", self.readInt32())
        self.printAndSet("cCh", self.readInt32())
        self.printAndSet("cPg", self.readInt16())
        self.printAndSet("cParas", self.readInt32())

        buf = self.readuInt16()
        self.printAndSet("rncEdn", buf & 0x0003) # 1..2nd bits
        self.printAndSet("nEdn", (buf & 0xfffc) >> 2) # 3..16th bits

        buf = self.readuInt16()
        self.printAndSet("epc", buf & 0x0003) # 1..2nd bits
        self.printAndSet("unused14", (buf & 0x003c) >> 2) # 3..6th bits
        self.printAndSet("unused15", (buf & 0x03c0) >> 6) # 7..10th bits
        self.printAndSet("fPrintFormData", self.getBit(buf, 10))
        self.printAndSet("fSaveFormData", self.getBit(buf, 11))
        self.printAndSet("fShadeFormData", self.getBit(buf, 12))
        self.printAndSet("fShadeMergeFields", self.getBit(buf, 13))
        self.printAndSet("reserved2", self.getBit(buf, 14))
        self.printAndSet("fIncludeSubdocsInStats", self.getBit(buf, 15))

        self.printAndSet("cLines", self.readInt32())
        self.printAndSet("cWordsWithSubdocs", self.readInt32())
        self.printAndSet("cChWithSubdocs", self.readInt32())
        self.printAndSet("cPgWithSubdocs", self.readInt16())
        self.printAndSet("cParasWithSubdocs", self.readInt32())
        self.printAndSet("cLinesWithSubdocs", self.readInt32())
        self.printAndSet("lKeyProtDoc", self.readInt32())

        buf = self.readuInt16()
        self.printAndSet("wvkoSaved", buf & 0x0007) # 1..3rd bits
        self.printAndSet("pctWwdSaved", (buf & 0x0ff8) >> 3) # 4..12th bits
        self.printAndSet("zkSaved", (buf & 0x3000) >> 12) # 13..14th bits
        self.printAndSet("unused16", self.getBit(buf, 14))
        self.printAndSet("iGutterPos", self.getBit(buf, 15))
        print '</dopBase>'

class Copts80(DOCDirStream):
    """The Copts80 structure specifies compatibility options."""
    def __init__(self, dop):
        DOCDirStream.__init__(self, dop.bytes)
        self.pos = dop.pos

    def dump(self):
        print '<copts80 type="Copts80" offset="%d" size="4 bytes">' % self.pos
        Copts60(self).dump()
        self.pos += 2

        buf = self.readuInt8()
        self.printAndSet("fSuppressTopSpacingMac5", self.getBit(buf, 0))
        self.printAndSet("fTruncDxaExpand", self.getBit(buf, 1))
        self.printAndSet("fPrintBodyBeforeHdr", self.getBit(buf, 2))
        self.printAndSet("fNoExtLeading", self.getBit(buf, 3))
        self.printAndSet("fDontMakeSpaceForUL", self.getBit(buf, 4))
        self.printAndSet("fMWSmallCaps", self.getBit(buf, 5))
        self.printAndSet("f2ptExtLeadingOnly", self.getBit(buf, 6))
        self.printAndSet("fTruncFontHeight", self.getBit(buf, 7))

        buf = self.readuInt8()
        self.printAndSet("fSubOnSize", self.getBit(buf, 0))
        self.printAndSet("fLineWrapLikeWord6", self.getBit(buf, 1))
        self.printAndSet("fWW6BorderRules", self.getBit(buf, 2))
        self.printAndSet("fExactOnTop", self.getBit(buf, 3))
        self.printAndSet("fExtraAfter", self.getBit(buf, 4))
        self.printAndSet("fWPSpace", self.getBit(buf, 5))
        self.printAndSet("fWPJust", self.getBit(buf, 6))
        self.printAndSet("fPrintMet", self.getBit(buf, 7))
        print '</copts80>'

class Copts(DOCDirStream):
    """A structure that specifies compatibility options."""
    def __init__(self, dop):
        DOCDirStream.__init__(self, dop.bytes)
        self.pos = dop.pos

    def dump(self):
        print '<copts type="Copts" offset="%d" size="32 bytes">' % self.pos
        Copts80(self).dump()
        self.pos += 4

        buf = self.readuInt8()
        self.printAndSet("fSpLayoutLikeWW8", self.getBit(buf, 0))
        self.printAndSet("fFtnLayoutLikeWW8", self.getBit(buf, 1))
        self.printAndSet("fDontUseHTMLParagraphAutoSpacing", self.getBit(buf, 2))
        self.printAndSet("fDontAdjustLineHeightInTable", self.getBit(buf, 3))
        self.printAndSet("fForgetLastTabAlign", self.getBit(buf, 4))
        self.printAndSet("fUseAutospaceForFullWidthAlpha", self.getBit(buf, 5))
        self.printAndSet("fAlignTablesRowByRow", self.getBit(buf, 6))
        self.printAndSet("fLayoutRawTableWidth", self.getBit(buf, 7))

        buf = self.readuInt8()
        self.printAndSet("fLayoutTableRowsApart", self.getBit(buf, 0))
        self.printAndSet("fUseWord97LineBreakingRules", self.getBit(buf, 1))
        self.printAndSet("fDontBreakWrappedTables", self.getBit(buf, 2))
        self.printAndSet("fDontSnapToGridInCell", self.getBit(buf, 3))
        self.printAndSet("fDontAllowFieldEndSelect", self.getBit(buf, 4))
        self.printAndSet("fApplyBreakingRules", self.getBit(buf, 5))
        self.printAndSet("fDontWrapTextWithPunct", self.getBit(buf, 6))
        self.printAndSet("fDontUseAsianBreakRules", self.getBit(buf, 7))

        buf = self.readuInt8()
        self.printAndSet("fUseWord2002TableStyleRules", self.getBit(buf, 0))
        self.printAndSet("fGrowAutoFit", self.getBit(buf, 1))
        self.printAndSet("fUseNormalStyleForList", self.getBit(buf, 2))
        self.printAndSet("fDontUseIndentAsNumberingTabStop", self.getBit(buf, 3))
        self.printAndSet("fFELineBreak11", self.getBit(buf, 4))
        self.printAndSet("fAllowSpaceOfSameStyleInTable", self.getBit(buf, 5))
        self.printAndSet("fWW11IndentRules", self.getBit(buf, 6))
        self.printAndSet("fDontAutofitConstrainedTables", self.getBit(buf, 7))

        buf = self.readuInt8()
        self.printAndSet("fAutofitLikeWW11", self.getBit(buf, 0))
        self.printAndSet("fUnderlineTabInNumList", self.getBit(buf, 1))
        self.printAndSet("fHangulWidthLikeWW11", self.getBit(buf, 2))
        self.printAndSet("fSplitPgBreakAndParaMark", self.getBit(buf, 3))
        self.printAndSet("fDontVertAlignCellWithSp", self.getBit(buf, 4))
        self.printAndSet("fDontBreakConstrainedForcedTables", self.getBit(buf, 5))
        self.printAndSet("fDontVertAlignInTxbx", self.getBit(buf, 6))
        self.printAndSet("fWord11KerningPairs", self.getBit(buf, 7))

        buf = self.readuInt32()
        self.printAndSet("fCachedColBalance", self.getBit(buf, 0))
        self.printAndSet("empty1", (buf & 0xfffffffe) >> 1) # 2..32th bits
        self.printAndSet("empty2", self.readuInt32())
        self.printAndSet("empty3", self.readuInt32())
        self.printAndSet("empty4", self.readuInt32())
        self.printAndSet("empty5", self.readuInt32())
        self.printAndSet("empty6", self.readuInt32())
        print '</copts>'

class Dop95(DOCDirStream):
    """The Dop95 structure contains document and compatibility settings."""
    def __init__(self, dop):
        DOCDirStream.__init__(self, dop.bytes)
        self.pos = dop.pos
        self.dop = dop

    def dump(self):
        print '<dop95 type="Dop95" offset="%d" size="88 bytes">' % self.pos
        dopBase = DopBase(self)
        dopBase.dump()
        assert dopBase.pos == self.pos + 84
        self.pos += 84
        Copts80(self).dump()
        self.pos += 4
        print '</dop95>'

class DopTypography(DOCDirStream):
    """The DopTypography structure contains East Asian language typography settings."""
    def __init__(self, dop):
        DOCDirStream.__init__(self, dop.bytes)
        self.pos = dop.pos

    def dump(self):
        print '<dopTypography type="DopTypography" offset="%d" size="310 bytes">' % self.pos
        buf = self.readuInt16()
        self.printAndSet("fKerningPunct", self.getBit(buf, 0))
        self.printAndSet("iJustification", (buf & 0x0006) >> 1) # 2..3rd bits
        self.printAndSet("iLevelOfKinsoku", (buf & 0x0018) >> 1) # 4..5th bits
        self.printAndSet("f2on1", self.getBit(buf, 5))
        self.printAndSet("unused", self.getBit(buf, 6))
        self.printAndSet("iCustomKsu", (buf & 0x0380) >> 7) # 8..10th bits
        self.printAndSet("fJapaneseUseLevel2", self.getBit(buf, 10))
        self.printAndSet("reserved", (buf & 0xf800) >> 11) # 12..16th bits

        self.printAndSet("cchFollowingPunct", self.readInt16())
        self.printAndSet("cchLeadingPunct", self.readInt16())
        if self.cchFollowingPunct != 0:
            print '<todo what="DopTypography::dump(): cchFollowingPunct != 0 not handled"/>'
        else:
            self.pos += 202
        if self.cchLeadingPunct != 0:
            print '<todo what="DopTypography::dump(): cchLeadingPunct != 0 not handled"/>'
        else:
            self.pos += 102
        print '</dopTypography>'

class Dogrid(DOCDirStream):
    """The Dogrid structure specifies parameters for the drawn object properties of the document."""
    def __init__(self, dop):
        DOCDirStream.__init__(self, dop.bytes)
        self.pos = dop.pos

    def dump(self):
        print '<dogrid type="Dogrid" offset="%d" size="10 bytes">' % self.pos
        self.printAndSet("xaGrid", self.readuInt16())
        self.printAndSet("yaGrid", self.readuInt16())
        self.printAndSet("dxaGrid", self.readuInt16())
        self.printAndSet("dyaGrid", self.readuInt16())

        buf = self.readuInt8()
        self.printAndSet("dyGridDisplay", (buf & 0x7f)) # 1..7th bits
        self.printAndSet("unused", self.getBit(buf, 7))

        buf = self.readuInt8()
        self.printAndSet("dxGridDisplay", (buf & 0x7f)) # 1..7th bits
        self.printAndSet("fFollowMargins", self.getBit(buf, 7))
        print '</dogrid>'

class Asumyi(DOCDirStream):
    """The Asumyi structure specifies AutoSummary state information"""
    def __init__(self, dop):
        DOCDirStream.__init__(self, dop.bytes)
        self.pos = dop.pos

    def dump(self):
        print '<asumyi type="Asumyi" offset="%d" size="12 bytes">' % self.pos
        buf = self.readuInt16()
        self.printAndSet("fValid", self.getBit(buf, 0))
        self.printAndSet("fView", self.getBit(buf, 1))
        self.printAndSet("iViewBy", (buf & 0x0c) >> 2) # 3..4th bits
        self.printAndSet("fUpdateProps", self.getBit(buf, 4))
        self.printAndSet("reserved", (buf & 0xffe0) >> 5) # 6..16th bits

        self.printAndSet("wDlgLevel", self.readuInt16())
        self.printAndSet("lHighestLevel", self.readuInt32())
        self.printAndSet("lCurrentLevel", self.readuInt32())
        print '</asumyi>'

class Dop97(DOCDirStream):
    """The Dop97 structure contains document and compatibility settings."""
    def __init__(self, dop):
        DOCDirStream.__init__(self, dop.bytes)
        self.pos = dop.pos
        self.dop = dop

    def dump(self):
        print '<dop97 type="Dop97" offset="%d" size="500 bytes">' % self.pos
        dop95 = Dop95(self)
        dop95.dump()
        assert dop95.pos == self.pos + 88
        self.pos += 88

        self.printAndSet("adt", self.readuInt16())
        dopTypography = DopTypography(self)
        dopTypography.dump()
        assert dopTypography.pos == self.pos + 310
        self.pos += 310
        dogrid = Dogrid(self)
        dogrid.dump()
        assert dogrid.pos == self.pos + 10
        self.pos += 10

        buf = self.readuInt8()
        self.printAndSet("unused1", self.getBit(buf, 0))
        self.printAndSet("lvlDop", (buf & 0x1e) >> 1) # 2..5th bits
        self.printAndSet("fGramAllDone", self.getBit(buf, 5))
        self.printAndSet("fGramAllClean", self.getBit(buf, 6))
        self.printAndSet("fSubsetFonts", self.getBit(buf, 7))

        buf = self.readuInt8()
        self.printAndSet("unused2", self.getBit(buf, 0))
        self.printAndSet("fHtmlDoc", self.getBit(buf, 1))
        self.printAndSet("fDiskLvcInvalid", self.getBit(buf, 2))
        self.printAndSet("fSnapBorder", self.getBit(buf, 3))
        self.printAndSet("fIncludeHeader", self.getBit(buf, 4))
        self.printAndSet("fIncludeFooter", self.getBit(buf, 5))
        self.printAndSet("unused3", self.getBit(buf, 6))
        self.printAndSet("unused4", self.getBit(buf, 7))

        self.printAndSet("unused5", self.readuInt16())
        Asumyi(self).dump()
        self.pos += 12
        self.printAndSet("cChWS", self.readuInt32())
        self.printAndSet("cChWSWithSubdocs", self.readuInt32())
        self.printAndSet("grfDocEvents", self.readuInt32())

        buf = self.readuInt32()
        self.printAndSet("fVirusPrompted", self.getBit(buf, 0))
        self.printAndSet("fVirusLoadSafe", self.getBit(buf, 1))
        self.printAndSet("KeyVirusSession30", (buf & 0xfffffffc) >> 2)

        self.printAndSet("space1", self.readuInt32())
        self.printAndSet("space2", self.readuInt32())
        self.printAndSet("space3", self.readuInt32())
        self.printAndSet("space4", self.readuInt32())
        self.printAndSet("space5", self.readuInt32())
        self.printAndSet("space6", self.readuInt32())
        self.printAndSet("space7", self.readuInt32())
        self.printAndSet("space8", self.readuInt16())

        self.printAndSet("cpMaxListCacheMainDoc", self.readuInt32())
        self.printAndSet("ilfoLastBulletMain", self.readuInt16())
        self.printAndSet("ilfoLastNumberMain", self.readuInt16())
        self.printAndSet("cDBC", self.readuInt32())
        self.printAndSet("cDBCWithSubdocs", self.readuInt32())
        self.printAndSet("reserved3a", self.readuInt32())
        self.printAndSet("nfcFtnRef", self.readuInt16())
        self.printAndSet("nfcEdnRef", self.readuInt16())
        self.printAndSet("hpsZoomFontPag", self.readuInt16())
        self.printAndSet("dywDispPag", self.readuInt16())
        print '</dop97>'

class Dop2000(DOCDirStream):
    """The Dop2000 structure contains document and compatibility settings."""
    def __init__(self, dop):
        DOCDirStream.__init__(self, dop.bytes)
        self.pos = dop.pos
        self.dop = dop

    def dump(self):
        print '<dop2000 type="Dop2000" offset="%d" size="544 bytes">' % self.pos
        dop97 = Dop97(self)
        dop97.dump()
        assert dop97.pos == self.pos + 500
        self.pos += 500

        self.printAndSet("ilvlLastBulletMain", self.readuInt8())
        self.printAndSet("ilvlLastNumberMain", self.readuInt8())
        self.printAndSet("istdClickParaType", self.readuInt16())

        buf = self.readuInt8()
        self.printAndSet("fLADAllDone", self.getBit(buf, 0))
        self.printAndSet("fEnvelopeVis", self.getBit(buf, 1))
        self.printAndSet("fMaybeTentativeListInDoc", self.getBit(buf, 2))
        self.printAndSet("fMaybeFitText", self.getBit(buf, 3))
        self.printAndSet("empty1", (buf & 0xf0) >> 4) # 5..8th bits

        buf = self.readuInt8()
        self.printAndSet("fFCCAllDone", self.getBit(buf, 0))
        self.printAndSet("fRelyOnCSS_WebOpt", self.getBit(buf, 1))
        self.printAndSet("fRelyOnVML_WebOpt", self.getBit(buf, 2))
        self.printAndSet("fAllowPNG_WebOpt", self.getBit(buf, 3))
        self.printAndSet("screenSize_WebOpt", (buf & 0xf0) >> 4) # 5..8th bits

        buf = self.readuInt16()
        self.printAndSet("fOrganizeInFolder_WebOpt", self.getBit(buf, 0))
        self.printAndSet("fUseLongFileNames_WebOpt", self.getBit(buf, 1))
        self.printAndSet("iPixelsPerInch_WebOpt", (buf & 0x0ffc) >> 2) # 3..12th bits
        self.printAndSet("fWebOptionsInit", self.getBit(buf, 12))
        self.printAndSet("fMaybeFEL", self.getBit(buf, 12))
        self.printAndSet("fCharLineUnits", self.getBit(buf, 12))
        self.printAndSet("unused1", self.getBit(buf, 12))

        copts = Copts(self)
        copts.dump()
        assert copts.pos == self.pos + 32
        self.pos += 32

        self.printAndSet("verCompatPre10", self.readuInt16())
        buf = self.readuInt8()
        self.printAndSet("fNoMargPgvwSaved", self.getBit(buf, 0))
        self.printAndSet("unused2", self.getBit(buf, 1))
        self.printAndSet("unused3", self.getBit(buf, 2))
        self.printAndSet("unused4", self.getBit(buf, 3))
        self.printAndSet("fBulletProofed", self.getBit(buf, 4))
        self.printAndSet("empty2", self.getBit(buf, 5))
        self.printAndSet("fSaveUim", self.getBit(buf, 6))
        self.printAndSet("fFilterPrivacy", self.getBit(buf, 7))

        buf = self.readuInt8()
        self.printAndSet("empty3", self.getBit(buf, 0))
        self.printAndSet("fSeenRepairs", self.getBit(buf, 1))
        self.printAndSet("fHasXML", self.getBit(buf, 2))
        self.printAndSet("unused5", self.getBit(buf, 3))
        self.printAndSet("fValidateXML", self.getBit(buf, 4))
        self.printAndSet("fSaveInvalidXML", self.getBit(buf, 5))
        self.printAndSet("fShowXMLErrors", self.getBit(buf, 6))
        self.printAndSet("fAlwaysMergeEmptyNamespace", self.getBit(buf, 7))
        print '</dop2000>'

class Dop2002(DOCDirStream):
    """The Dop2002 structure contains document and compatibility settings."""
    def __init__(self, dop):
        DOCDirStream.__init__(self, dop.bytes)
        self.pos = dop.pos
        self.dop = dop

    def dump(self):
        print '<dop2002 type="Dop2002" offset="%d" size="594 bytes">' % self.pos
        dop2000 = Dop2000(self)
        dop2000.dump()
        assert dop2000.pos == self.pos + 544
        self.pos += 544

        self.printAndSet("unused", self.readuInt32())

        buf = self.readuInt8()
        self.printAndSet("fDoNotEmbedSystemFont", self.getBit(buf, 0))
        self.printAndSet("fWordCompat", self.getBit(buf, 1))
        self.printAndSet("fLiveRecover", self.getBit(buf, 2))
        self.printAndSet("fEmbedFactoids", self.getBit(buf, 3))
        self.printAndSet("fFactoidXML", self.getBit(buf, 4))
        self.printAndSet("fFactoidAllDone", self.getBit(buf, 5))
        self.printAndSet("fFolioPrint", self.getBit(buf, 6))
        self.printAndSet("fReverseFolio", self.getBit(buf, 7))

        buf = self.readuInt8()
        self.printAndSet("iTextLineEnding", (buf & 0x7)) # 1..3rd bits
        self.printAndSet("fHideFcc", self.getBit(buf, 3))
        self.printAndSet("fAcetateShowMarkup", self.getBit(buf, 4))
        self.printAndSet("fAcetateShowAtn", self.getBit(buf, 5))
        self.printAndSet("fAcetateShowInsDel", self.getBit(buf, 6))
        self.printAndSet("fAcetateShowProps", self.getBit(buf, 7))

        self.printAndSet("istdTableDflt", self.readuInt16())
        self.printAndSet("verCompat", self.readuInt16())
        self.printAndSet("grfFmtFilter", self.readuInt16())
        self.printAndSet("iFolioPages", self.readuInt16())
        self.printAndSet("cpgText", self.readuInt32())
        self.printAndSet("cpMinRMText", self.readuInt32())
        self.printAndSet("cpMinRMFtn", self.readuInt32())
        self.printAndSet("cpMinRMHdd", self.readuInt32())
        self.printAndSet("cpMinRMAtn", self.readuInt32())
        self.printAndSet("cpMinRMEdn", self.readuInt32())
        self.printAndSet("cpMinRmTxbx", self.readuInt32())
        self.printAndSet("cpMinRmHdrTxbx", self.readuInt32())
        self.printAndSet("rsidRoot", self.readuInt32())
        print '</dop2002>'

class Dop2003(DOCDirStream):
    """The Dop2003 structure contains document and compatibility settings."""
    def __init__(self, dop):
        DOCDirStream.__init__(self, dop.bytes)
        self.pos = dop.pos
        self.dop = dop

    def dump(self):
        print '<dop2003 type="Dop2003" offset="%d" size="616 bytes">' % self.pos
        dop2002 = Dop2002(self)
        dop2002.dump()
        assert dop2002.pos == self.pos + 594
        self.pos += 594

        buf = self.readuInt8()
        self.printAndSet("fTreatLockAtnAsReadOnly", self.getBit(buf, 0))
        self.printAndSet("fStyleLock", self.getBit(buf, 1))
        self.printAndSet("fAutoFmtOverride", self.getBit(buf, 2))
        self.printAndSet("fRemoveWordML", self.getBit(buf, 3))
        self.printAndSet("fApplyCustomXForm", self.getBit(buf, 4))
        self.printAndSet("fStyleLockEnforced", self.getBit(buf, 5))
        self.printAndSet("fFakeLockAtn", self.getBit(buf, 6))
        self.printAndSet("fIgnoreMixedContent", self.getBit(buf, 7))

        buf = self.readuInt8()
        self.printAndSet("fShowPlaceholderText", self.getBit(buf, 0))
        self.printAndSet("unused", self.getBit(buf, 1))
        self.printAndSet("fWord97Doc", self.getBit(buf, 2))
        self.printAndSet("fStyleLockTheme", self.getBit(buf, 3))
        self.printAndSet("fStyleLockQFSet", self.getBit(buf, 4))
        self.printAndSet("empty1", (buf & 0xe0) >> 5) # 6..8th bits
        self.printAndSet("empty1_", self.readuInt16())

        buf = self.readuInt8()
        self.printAndSet("fReadingModeInkLockDown", self.getBit(buf, 0))
        self.printAndSet("fAcetateShowInkAtn", self.getBit(buf, 1))
        self.printAndSet("fFilterDttm", self.getBit(buf, 2))
        self.printAndSet("fEnforceDocProt", self.getBit(buf, 3))
        self.printAndSet("iDocProtCur", (buf & 0x70) >> 4) # 5..7th bits
        self.printAndSet("fDispBkSpSaved", self.getBit(buf, 7))

        self.printAndSet("empty2", self.readuInt8())
        self.printAndSet("dxaPageLock", self.readuInt32())
        self.printAndSet("dyaPageLock", self.readuInt32())
        self.printAndSet("pctFontLock", self.readuInt32())
        self.printAndSet("grfitbid", self.readuInt8())
        self.printAndSet("empty3", self.readuInt8())
        self.printAndSet("ilfoMacAtCleanup", self.readuInt16())
        print '</dop2003>'

class DopMth(DOCDirStream):
    """The DopMth structure specifies document-wide math settings."""
    def __init__(self, dop):
        DOCDirStream.__init__(self, dop.bytes)
        self.pos = dop.pos

    def dump(self):
        print '<dopMth type="DopMth" offset="%d" size="34 bytes">' % self.pos
        buf = self.readuInt32()
        self.printAndSet("mthbrk", (buf & 0x03)) # 1..2nd bits
        self.printAndSet("mthbrkSub", (buf & 0xc) >> 2) # 3..4th bits
        self.printAndSet("mthbpjc", (buf & 0x70) >> 4) # 5..7th bits
        self.printAndSet("reserved1", self.getBit(buf, 7))
        self.printAndSet("fMathSmallFrac", self.getBit(buf, 8))
        self.printAndSet("fMathIntLimUndOvr", self.getBit(buf, 9))
        self.printAndSet("fMathNaryLimUndOvr", self.getBit(buf, 10))
        self.printAndSet("fMathWrapAlignLeft", self.getBit(buf, 11))
        self.printAndSet("fMathUseDispDefaults", self.getBit(buf, 12))
        self.printAndSet("reserved2", (buf & 0xffffe000) >> 13) # 14..32th bits

        self.printAndSet("ftcMath", self.readuInt16())
        self.printAndSet("dxaLeftMargin", self.readuInt32())
        self.printAndSet("dxaRightMargin", self.readuInt32())
        self.printAndSet("empty1", self.readuInt32())
        self.printAndSet("empty2", self.readuInt32())
        self.printAndSet("empty3", self.readuInt32())
        self.printAndSet("empty4", self.readuInt32())
        self.printAndSet("dxaIndentWrapped", self.readuInt32())
        print '</dopMth>'

class Dop2007(DOCDirStream):
    """The Dop2007 structure contains document and compatibility settings."""
    def __init__(self, dop):
        DOCDirStream.__init__(self, dop.bytes)
        self.pos = dop.pos
        self.dop = dop

    def dump(self):
        print '<dop2007 type="Dop2007" offset="%d">' % self.pos
        dop2003 = Dop2003(self)
        dop2003.dump()
        assert dop2003.pos == self.pos + 616
        self.pos += 616

        self.printAndSet("reserved1", self.readuInt32())

        buf = self.readuInt16()
        self.printAndSet("fRMTrackFormatting", self.getBit(buf, 0))
        self.printAndSet("fRMTrackMoves", self.getBit(buf, 1))
        self.printAndSet("reserved2", self.getBit(buf, 2))
        self.printAndSet("empty1", self.getBit(buf, 3))
        self.printAndSet("empty2", self.getBit(buf, 4))
        self.printAndSet("ssm", (buf & 0x01e0) >> 5) # 6..9th bits
        self.printAndSet("fReadingModeInkLockDownActualPage", self.getBit(buf, 9))
        self.printAndSet("fAutoCompressPictures", self.getBit(buf, 10))
        self.printAndSet("reserved3", (buf & 0xf800) >> 11) # 12..16th bits
        self.printAndSet("reserved3_", self.readuInt16())

        self.printAndSet("empty3", self.readuInt32())
        self.printAndSet("empty4", self.readuInt32())
        self.printAndSet("empty5", self.readuInt32())
        self.printAndSet("empty6", self.readuInt32())
        DopMth(self).dump()
        self.pos += 34
        print '</dop2007>'

class Dop(DOCDirStream):
    """The Dop structure contains the document and compatibility settings for the document."""
    def __init__(self, fib):
        DOCDirStream.__init__(self, fib.doc.getDirectoryStreamByName("1Table").bytes)
        self.pos = fib.fcDop
        self.size = fib.lcbDop
        self.fib = fib

    def dump(self):
        print '<dop type="Dop" offset="%s" size="%d bytes">' % (self.pos, self.size)
        if self.fib.nFibNew == 0x0112:
            Dop2007(self).dump()
        else:
            print """<todo what="Dop.dump() doesn't know how to handle nFibNew = %s">""" % hex(self.nFibNew)
        print '</dop>'

class FFID(DOCDirStream):
    """The FFID structure specifies the font family and character pitch for a font."""
    def __init__(self, bytes, offset):
        DOCDirStream.__init__(self, bytes)
        self.pos = offset

    def dump(self):
        self.ffid = self.readuInt8()

        self.prq =       (self.ffid & 0x3)       # first two bits
        self.fTrueType = (self.ffid & 0x4) >> 2  # 3rd bit
        self.unused1 =   (self.ffid & 0x8) >> 3  # 4th bit
        self.ff =        (self.ffid & 0x70) >> 4 # 5-7th bits
        self.unused2 =   (self.ffid & 0x80) >> 7 # 8th bit

        print '<ffid value="%s" prq="%s" fTrueType="%s" ff="%s"/>' % (hex(self.ffid), hex(self.prq), self.fTrueType, hex(self.ff))

class PANOSE(DOCDirStream):
    """The PANOSE structure defines the PANOSE font classification values for a TrueType font."""
    def __init__(self, bytes, offset):
        DOCDirStream.__init__(self, bytes)
        self.pos = offset

    def dump(self):
        print '<panose type="PANOSE" offset="%s" size="10 bytes">' % self.pos
        for i in ["bFamilyType", "bSerifStyle", "bWeight", "bProportion", "bContrast", "bStrokeVariation", "bArmStyle", "bLetterform", "bMidline", "bHeight"]:
            self.printAndSet(i, self.readuInt8())
        print '</panose>'

class FontSignature(DOCDirStream):
    """Contains information identifying the code pages and Unicode subranges for which a given font provides glyphs."""
    def __init__(self, bytes, offset):
        DOCDirStream.__init__(self, bytes)
        self.pos = offset

    def dump(self):
        fsUsb1 = self.readuInt32()
        fsUsb2 = self.readuInt32()
        fsUsb3 = self.readuInt32()
        fsUsb4 = self.readuInt32()
        fsCsb1 = self.readuInt32()
        fsCsb2 = self.readInt32()
        print '<fontSignature fsUsb1="%s" fsUsb2="%s" fsUsb3="%s" fsUsb4="%s" fsCsb1="%s" fsCsb2="%s"/>' % (
                hex(fsUsb1), hex(fsUsb2), hex(fsUsb3), hex(fsUsb4), hex(fsCsb1), hex(fsCsb2)
                )

class FFN(DOCDirStream):
    """The FFN structure specifies information about a font that is used in the document."""
    def __init__(self, bytes, mainStream, offset, size):
        DOCDirStream.__init__(self, bytes, mainStream=mainStream)
        self.pos = offset
        self.size = size

    def dump(self):
        print '<ffn type="FFN" offset="%d" size="%d bytes">' % (self.pos, self.size)
        FFID(self.bytes, self.pos).dump()
        self.pos += 1
        self.printAndSet("wWeight", self.readInt16(), hexdump = False)
        self.printAndSet("chs", self.readuInt8(), hexdump = False)
        self.printAndSet("ixchSzAlt", self.readuInt8())
        PANOSE(self.bytes, self.pos).dump()
        self.pos += 10
        FontSignature(self.bytes, self.pos).dump()
        self.pos += 24
        print '<xszFfn value="%s"/>' % self.getString()
        print '</ffn>'

class SttbfFfn(DOCDirStream):
    """The SttbfFfn structure is an STTB whose strings are FFN records that specify details of system fonts."""
    def __init__(self, bytes, mainStream, offset, size):
        DOCDirStream.__init__(self, bytes, mainStream=mainStream)
        self.pos = offset
        self.size = size

    def dump(self):
        print '<sttbfFfn type="SttbfFfn" offset="%d" size="%d bytes">' % (self.pos, self.size)
        self.printAndSet("cData", self.readuInt16())
        self.printAndSet("cbExtra", self.readuInt16())
        for i in range(self.cData):
            cchData = self.readuInt8()
            print '<cchData index="%d" offset="%d" size="%d bytes">' % (i, self.pos, cchData)
            FFN(self.bytes, self.mainStream, self.pos, cchData).dump()
            self.pos += cchData
            print '</cchData>'
        print '</sttbfFfn>'

class ATNBE(DOCDirStream):
    """The ATNBE structure contains information about an annotation bookmark in the document."""
    size = 10 # in bytes, see 2.9.4
    def __init__(self, sttbfAtnBkmk):
        DOCDirStream.__init__(self, sttbfAtnBkmk.bytes)
        self.pos = sttbfAtnBkmk.pos

    def dump(self):
        print '<atnbe type="ATNBE">'
        self.printAndSet("bmc", self.readuInt16())
        self.printAndSet("ITag", self.readuInt32())
        self.printAndSet("ITagOld", self.readuInt32())
        print '</atnbe>'

class SttbfAtnBkmk(DOCDirStream):
    """The SttbfAtnBkmk structure is an STTB whose strings are all of zero length."""
    def __init__(self, mainStream, offset, size):
        DOCDirStream.__init__(self, mainStream.doc.getDirectoryStreamByName("1Table").bytes)
        self.pos = offset
        self.size = size

    def dump(self):
        print '<sttbfAtnBkmk type="SttbfAtnBkmk" offset="%d" size="%d bytes">' % (self.pos, self.size)
        self.printAndSet("fExtended", self.readuInt16())
        self.printAndSet("cData", self.readuInt16())
        self.printAndSet("cbExtra", self.readuInt16())
        for i in range(self.cData):
            cchData = self.readuInt16()
            print '<cchData index="%d" offset="%d" size="%d bytes"/>' % (i, self.pos, cchData)
            print '<extraData index="%d" offset="%d" size="%d bytes">' % (i, self.pos, ATNBE.size)
            atnbe = ATNBE(self)
            atnbe.dump()
            self.pos += ATNBE.size
            print '</extraData>'
        print '</sttbfAtnBkmk>'

class Stshif(DOCDirStream):
    """The Stshif structure specifies general stylesheet information."""
    def __init__(self, bytes, mainStream, offset):
        DOCDirStream.__init__(self, bytes, mainStream=mainStream)
        self.pos = offset
        self.size = 18

    def dump(self):
        print '<stshif type="Stshif" offset="%d" size="%d bytes">' % (self.pos, self.size)
        self.printAndSet("cstd", self.readuInt16())
        self.printAndSet("cbSTDBaseInFile", self.readuInt16())
        buf = self.readuInt16()
        self.printAndSet("fStdStylenamesWritten", buf & 1) # first bit
        self.printAndSet("fReserved", (buf & 0xfe) >> 1) # 2..16th bits
        self.printAndSet("stiMaxWhenSaved", self.readuInt16())
        self.printAndSet("istdMaxFixedWhenSaved", self.readuInt16())
        self.printAndSet("nVerBuiltInNamesWhenSaved", self.readuInt16())
        self.printAndSet("ftcAsci", self.readuInt16())
        self.printAndSet("ftcFE", self.readuInt16())
        self.printAndSet("ftcOther", self.readuInt16())
        print '</stshif>'

class LSD(DOCDirStream):
    """The LSD structure specifies the properties to be used for latent application-defined styles (see StshiLsd) when they are created."""
    def __init__(self, bytes, offset):
        DOCDirStream.__init__(self, bytes)
        self.pos = offset

    def dump(self):
        buf = self.readuInt16()
        self.printAndSet("fLocked", self.getBit(buf, 1))
        self.printAndSet("fSemiHidden", self.getBit(buf, 2))
        self.printAndSet("fUnhideWhenUsed", self.getBit(buf, 3))
        self.printAndSet("fQFormat", self.getBit(buf, 4))
        self.printAndSet("iPriority", (buf & 0xfff0) >> 4) # 5-16th bits
        self.printAndSet("fReserved", self.readuInt16())

class StshiLsd(DOCDirStream):
    """The StshiLsd structure specifies latent style data for application-defined styles."""
    def __init__(self, bytes, stshi, offset):
        DOCDirStream.__init__(self, bytes)
        self.stshi = stshi
        self.pos = offset
    
    def dump(self):
        print '<stshiLsd type="StshiLsd" offset="%d">' % (self.pos)
        self.printAndSet("cbLSD", self.readuInt16())
        for i in range(self.stshi.stshif.stiMaxWhenSaved):
            print '<mpstiilsd index="%d" type="LSD">' % i
            LSD(self.bytes, self.pos).dump()
            print '</mpstiilsd>'
            self.pos += self.cbLSD
        print '</stshiLsd>'

class STSHI(DOCDirStream):
    """The STSHI structure specifies general stylesheet and related information."""
    def __init__(self, bytes, mainStream, offset, size):
        DOCDirStream.__init__(self, bytes, mainStream=mainStream)
        self.pos = offset
        self.size = size

    def dump(self):
        print '<stshi type="STSHI" offset="%d" size="%d bytes">' % (self.pos, self.size)
        self.stshif = Stshif(self.bytes, self.mainStream, self.pos)
        self.stshif.dump()
        self.pos += self.stshif.size
        self.printAndSet("ftcBi", self.readuInt16())
        stshiLsd = StshiLsd(self.bytes, self, self.pos)
        stshiLsd.dump()
        print '</stshi>'

class LPStshi(DOCDirStream):
    """The LPStshi structure specifies general stylesheet information."""
    def __init__(self, bytes, mainStream, offset):
        DOCDirStream.__init__(self, bytes, mainStream=mainStream)
        self.pos = offset

    def dump(self):
        print '<lpstshi type="LPStshi" offset="%d">' % self.pos
        self.printAndSet("cbStshi", self.readuInt16(), hexdump = False)
        self.stshi = STSHI(self.bytes, self.mainStream, self.pos, self.cbStshi)
        self.stshi.dump()
        self.pos += self.cbStshi
        print '</lpstshi>'

class StdfBase(DOCDirStream):
    """The Stdf structure specifies general information about the style."""
    def __init__(self, bytes, mainStream, offset):
        DOCDirStream.__init__(self, bytes, mainStream=mainStream)
        self.pos = offset
        self.size = 10

    def dump(self):
        print '<stdfBase type="StdfBase" offset="%d" size="%d bytes">' % (self.pos, self.size)
        buf = self.readuInt16()
        self.printAndSet("sti", buf & 0x0fff) # 1..12th bits
        self.printAndSet("fScratch", self.getBit(buf, 13))
        self.printAndSet("fInvalHeight", self.getBit(buf, 14))
        self.printAndSet("fHasUpe", self.getBit(buf, 15))
        self.printAndSet("fMassCopy", self.getBit(buf, 16))
        buf = self.readuInt16()
        self.stk = buf & 0x000f # 1..4th bits
        stkmap = {
                1: "paragraph",
                2: "character",
                3: "table",
                4: "numbering"
                }
        print '<stk value="%d" name="%s"/>' % (self.stk, stkmap[self.stk])
        self.printAndSet("istdBase", (buf & 0xfff0) >> 4) # 5..16th bits
        buf = self.readuInt16()
        self.printAndSet("cupx", buf & 0x000f) # 1..4th bits
        self.printAndSet("istdNext", (buf & 0xfff0) >> 4) # 5..16th bits
        self.printAndSet("bchUpe", self.readuInt16(), hexdump = False)
        self.printAndSet("grfstd", self.readuInt16()) # TODO dedicated GRFSTD class
        print '</stdfBase>'

class StdfPost2000(DOCDirStream):
    """The StdfPost2000 structure specifies general information about a style."""
    def __init__(self, stdf):
        DOCDirStream.__init__(self, stdf.bytes, mainStream = stdf.mainStream)
        self.pos = stdf.pos
        self.size = 8

    def dump(self):
        print '<stdfPost2000 type="StdfPost2000" offset="%d" size="%d bytes">' % (self.pos, self.size)
        buf = self.readuInt16()
        self.printAndSet("istdLink", buf & 0xfff) # 1..12th bits
        self.printAndSet("fHasOriginalStyle", self.getBit(buf, 13)) # 13th bit
        self.printAndSet("fSpare", (buf & 0xe000) >> 13) # 14..16th bits
        self.printAndSet("rsid", self.readuInt32())
        buf = self.readuInt16()
        self.printAndSet("iftcHtml", buf & 0x7) # 1..3rd bits
        self.printAndSet("unused", self.getBit(buf, 4))
        self.printAndSet("iPriority", (buf & 0xfff0) >> 4) # 5..16th bits
        print '</stdfPost2000>'

class Stdf(DOCDirStream):
    """The Stdf structure specifies general information about the style."""
    def __init__(self, std):
        DOCDirStream.__init__(self, std.bytes, mainStream = std.mainStream)
        self.std = std
        self.pos = std.pos

    def dump(self):
        print '<stdf type="Stdf" offset="%d">' % self.pos
        self.stdfBase = StdfBase(self.bytes, self.mainStream, self.pos)
        self.stdfBase.dump()
        self.pos += self.stdfBase.size
        stsh = self.std.lpstd.stsh # root of the stylesheet table
        cbSTDBaseInFile = stsh.lpstshi.stshi.stshif.cbSTDBaseInFile
        print '<stdfPost2000OrNone cbSTDBaseInFile="%s">' % hex(cbSTDBaseInFile)
        if cbSTDBaseInFile == 0x0012:
            stdfPost2000 = StdfPost2000(self)
            stdfPost2000.dump()
            self.pos = stdfPost2000.pos
        print '</stdfPost2000OrNone>'
        print '</stdf>'

class Xst(DOCDirStream):
    """The Xst structure is a string. The string is prepended by its length and is not null-terminated."""
    def __init__(self, parent):
        DOCDirStream.__init__(self, parent.bytes)
        self.pos = parent.pos

    def dump(self):
        print '<xst type="Xst" offset="%d">' % self.pos
        self.printAndSet("cch", self.readuInt16())
        print '<rgtchar value="%s"/>' % globals.encodeName(self.bytes[self.pos:self.pos+2*self.cch].decode('utf-16'), lowOnly = True)
        self.pos += 2*self.cch
        print '</xst>'

class Xstz(DOCDirStream):
    """The Xstz structure is a string. The string is prepended by its length and is null-terminated."""
    def __init__(self, parent):
        DOCDirStream.__init__(self, parent.bytes)
        self.pos = parent.pos

    def dump(self):
        print '<xstz type="Xstz" offset="%d">' % self.pos
        xst = Xst(self)
        xst.dump()
        self.pos = xst.pos
        self.printAndSet("chTerm", self.readuInt16())
        print '</xstz>'

class UpxPapx(DOCDirStream):
    """The UpxPapx structure specifies the paragraph formatting properties that differ from the parent"""
    def __init__(self, lPUpxPapx):
        DOCDirStream.__init__(self, lPUpxPapx.bytes)
        self.lPUpxPapx = lPUpxPapx
        self.pos = lPUpxPapx.pos

    def dump(self):
        print '<upxPapx type="UpxPapx" offset="%d">' % self.pos
        self.printAndSet("istd", self.readuInt16())
        size = self.lPUpxPapx.cbUpx - 2
        pos = 0
        print '<grpprlPapx offset="%d" size="%d bytes">' % (self.pos, size)
        while size - pos > 0:
            prl = Prl(self.bytes, self.pos + pos)
            prl.dump()
            pos += prl.getSize()
        print '</grpprlPapx>'
        print '</upxPapx>'

class UpxChpx(DOCDirStream):
    """The UpxChpx structure specifies the character formatting properties that differ from the parent"""
    def __init__(self, lPUpxChpx):
        DOCDirStream.__init__(self, lPUpxChpx.bytes)
        self.lPUpxChpx = lPUpxChpx
        self.pos = lPUpxChpx.pos

    def dump(self):
        print '<upxChpx type="UpxChpx" offset="%d">' % self.pos
        size = self.lPUpxChpx.cbUpx
        pos = 0
        print '<grpprlChpx offset="%d" size="%d bytes">' % (self.pos, size)
        while size - pos > 0:
            prl = Prl(self.bytes, self.pos + pos)
            prl.dump()
            pos += prl.getSize()
        print '</grpprlChpx>'
        print '</upxChpx>'

class UpxTapx(DOCDirStream):
    """The UpxTapx structure specifies the table formatting properties that differ from the parent"""
    def __init__(self, lPUpxTapx):
        DOCDirStream.__init__(self, lPUpxTapx.bytes)
        self.lPUpxTapx = lPUpxTapx
        self.pos = lPUpxTapx.pos

    def dump(self):
        print '<upxTapx type="UpxTapx" offset="%d">' % self.pos
        size = self.lPUpxTapx.cbUpx
        pos = 0
        print '<grpprlTapx offset="%d" size="%d bytes">' % (self.pos, size)
        while size - pos > 0:
            prl = Prl(self.bytes, self.pos + pos)
            prl.dump()
            pos += prl.getSize()
        print '</grpprlTapx>'
        print '</upxTapx>'

class UPXPadding:
    """The UPXPadding structure specifies the padding that is used to pad the UpxPapx/Chpx/Tapx structures if any of them are an odd number of bytes in length."""
    def __init__(self, parent):
        self.parent = parent
        self.pos = parent.pos

    def pad(self):
        if self.parent.cbUpx % 2 == 1:
            self.pos += 1

class LPUpxPapx(DOCDirStream):
    """The LPUpxPapx structure specifies paragraph formatting properties."""
    def __init__(self, stkParaGRLPUPX):
        DOCDirStream.__init__(self, stkParaGRLPUPX.bytes)
        self.pos = stkParaGRLPUPX.pos

    def dump(self):
        print '<lPUpxPapx type="LPUpxPapx" offset="%d">' % self.pos
        self.printAndSet("cbUpx", self.readuInt16())
        upxPapx = UpxPapx(self)
        upxPapx.dump()
        self.pos += self.cbUpx
        uPXPadding = UPXPadding(self)
        uPXPadding.pad()
        self.pos = uPXPadding.pos
        print '</lPUpxPapx>'

class LPUpxChpx(DOCDirStream):
    """The LPUpxChpx structure specifies character formatting properties."""
    def __init__(self, stkParaGRLPUPX):
        DOCDirStream.__init__(self, stkParaGRLPUPX.bytes)
        self.pos = stkParaGRLPUPX.pos

    def dump(self):
        print '<lPUpxChpx type="LPUpxChpx" offset="%d">' % self.pos
        self.printAndSet("cbUpx", self.readuInt16())
        upxChpx = UpxChpx(self)
        upxChpx.dump()
        self.pos += self.cbUpx
        uPXPadding = UPXPadding(self)
        uPXPadding.pad()
        self.pos = uPXPadding.pos
        print '</lPUpxChpx>'

class LPUpxTapx(DOCDirStream):
    """The LPUpxTapx structure specifies table formatting properties."""
    def __init__(self, stkParaGRLPUPX):
        DOCDirStream.__init__(self, stkParaGRLPUPX.bytes)
        self.pos = stkParaGRLPUPX.pos

    def dump(self):
        print '<lPUpxTapx type="LPUpxTapx" offset="%d">' % self.pos
        self.printAndSet("cbUpx", self.readuInt16())
        upxTapx = UpxTapx(self)
        upxTapx.dump()
        self.pos += self.cbUpx
        uPXPadding = UPXPadding(self)
        uPXPadding.pad()
        self.pos = uPXPadding.pos
        print '</lPUpxTapx>'

class StkCharLpUpxGrLpUpxRM(DOCDirStream):
    """The StkCharLPUpxGrLPUpxRM structure specifies revision-marking information and formatting for character styles."""
    def __init__(self, stkCharGRLPUPX):
        DOCDirStream.__init__(self, stkCharGRLPUPX.bytes)
        self.pos = stkCharGRLPUPX.pos

    def dump(self):
        print '<stkCharLpUpxGrLpUpxRM type="StkCharLpUpxGrLpUpxRM" offset="%d">' % self.pos
        self.printAndSet("cbStkCharUpxGrLpUpxRM", self.getuInt16())
        if self.cbStkCharUpxGrLpUpxRM != 0:
            print '<todo what="StkCharLpUpxGrLpUpxRM: cbStkCharUpxGrLpUpxRM != 0 not implemented"/>'
        print '</stkCharLpUpxGrLpUpxRM>'

class StkParaLpUpxGrLpUpxRM(DOCDirStream):
    """The StkParaLPUpxGrLPUpxRM structure specifies revision-marking information and formatting for paragraph styles."""
    def __init__(self, stkParaGRLPUPX):
        DOCDirStream.__init__(self, stkParaGRLPUPX.bytes)
        self.pos = stkParaGRLPUPX.pos

    def dump(self):
        print '<stkParaLpUpxGrLpUpxRM type="StkParaLpUpxGrLpUpxRM" offset="%d">' % self.pos
        self.printAndSet("cbStkParaUpxGrLpUpxRM", self.getuInt16())
        if self.cbStkParaUpxGrLpUpxRM != 0:
            print '<todo what="StkParaLpUpxGrLpUpxRM: cbStkParaUpxGrLpUpxRM != 0 not implemented"/>'
        print '</stkParaLpUpxGrLpUpxRM>'

class StkListGRLPUPX(DOCDirStream):
    """The StkListGRLPUPX structure that specifies the formatting properties for a list style."""
    def __init__(self, grLPUpxSw):
        DOCDirStream.__init__(self, grLPUpxSw.bytes)
        self.grLPUpxSw = grLPUpxSw
        self.pos = grLPUpxSw.pos

    def dump(self):
        print '<stkListGRLPUPX type="StkListGRLPUPX" offset="%d">' % self.pos
        lpUpxPapx = LPUpxPapx(self)
        lpUpxPapx.dump()
        self.pos = lpUpxPapx.pos
        print '</stkListGRLPUPX>'

class StkTableGRLPUPX(DOCDirStream):
    """The StkTableGRLPUPX structure that specifies the formatting properties for a table style."""
    def __init__(self, grLPUpxSw):
        DOCDirStream.__init__(self, grLPUpxSw.bytes)
        self.grLPUpxSw = grLPUpxSw
        self.pos = grLPUpxSw.pos

    def dump(self):
        print '<stkTableGRLPUPX type="StkTableGRLPUPX" offset="%d">' % self.pos
        lpUpxTapx = LPUpxTapx(self)
        lpUpxTapx.dump()
        self.pos = lpUpxTapx.pos
        lpUpxPapx = LPUpxPapx(self)
        lpUpxPapx.dump()
        self.pos = lpUpxPapx.pos
        lpUpxChpx = LPUpxChpx(self)
        lpUpxChpx.dump()
        self.pos = lpUpxChpx.pos
        print '</stkTableGRLPUPX>'

class StkCharGRLPUPX(DOCDirStream):
    """The StkCharGRLPUPX structure that specifies the formatting properties for a character style."""
    def __init__(self, grLPUpxSw):
        DOCDirStream.__init__(self, grLPUpxSw.bytes)
        self.grLPUpxSw = grLPUpxSw
        self.pos = grLPUpxSw.pos

    def dump(self):
        print '<stkCharGRLPUPX type="StkCharGRLPUPX" offset="%d">' % self.pos
        lpUpxChpx = LPUpxChpx(self)
        lpUpxChpx.dump()
        self.pos = lpUpxChpx.pos
        stkCharLpUpxGrLpUpxRM = StkCharLpUpxGrLpUpxRM(self)
        stkCharLpUpxGrLpUpxRM.dump()
        self.pos = stkCharLpUpxGrLpUpxRM.pos
        print '</stkCharGRLPUPX>'

class StkParaGRLPUPX(DOCDirStream):
    """The StkParaGRLPUPX structure that specifies the formatting properties for a paragraph style."""
    def __init__(self, grLPUpxSw):
        DOCDirStream.__init__(self, grLPUpxSw.bytes)
        self.grLPUpxSw = grLPUpxSw
        self.pos = grLPUpxSw.pos

    def dump(self):
        print '<stkParaGRLPUPX type="StkParaGRLPUPX" offset="%d">' % self.pos
        lPUpxPapx = LPUpxPapx(self)
        lPUpxPapx.dump()
        self.pos = lPUpxPapx.pos
        lpUpxChpx = LPUpxChpx(self)
        lpUpxChpx.dump()
        self.pos = lpUpxChpx.pos
        stkParaLpUpxGrLpUpxRM = StkParaLpUpxGrLpUpxRM(self)
        stkParaLpUpxGrLpUpxRM.dump()
        self.pos = stkParaLpUpxGrLpUpxRM.pos
        print '</stkParaGRLPUPX>'

class GrLPUpxSw(DOCDirStream):
    """The GrLPUpxSw structure is an array of variable-size structures that specify the formatting of the style."""
    def __init__(self, std):
        DOCDirStream.__init__(self, std.bytes)
        self.std = std
        self.pos = std.pos

    def dump(self):
        stkMap = {
                1: StkParaGRLPUPX,
                2: StkCharGRLPUPX,
                3: StkTableGRLPUPX,
                4: StkListGRLPUPX
                }
        child = stkMap[self.std.stdf.stdfBase.stk](self)
        child.dump()
        self.pos = child.pos

class STD(DOCDirStream):
    """The STD structure specifies a style definition."""
    def __init__(self, lpstd):
        DOCDirStream.__init__(self, lpstd.bytes, mainStream = lpstd.mainStream)
        self.lpstd = lpstd
        self.pos = lpstd.pos
        self.size = lpstd.cbStd

    def dump(self):
        print '<std type="STD" offset="%d" size="%d bytes">' % (self.pos, self.size)
        self.stdf = Stdf(self)
        self.stdf.dump()
        self.pos = self.stdf.pos
        xstzName = Xstz(self)
        xstzName.dump()
        self.pos = xstzName.pos
        grLPUpxSw = GrLPUpxSw(self)
        grLPUpxSw.dump()
        self.pos = grLPUpxSw.pos
        print '</std>'

class LPStd(DOCDirStream):
    """The LPStd structure specifies a length-prefixed style definition."""
    def __init__(self, stsh):
        DOCDirStream.__init__(self, stsh.bytes, mainStream = stsh.mainStream)
        self.stsh = stsh
        self.pos = stsh.pos

    def dump(self):
        self.printAndSet("cbStd", self.readuInt16())
        if self.cbStd > 0:
            std = STD(self)
            std.dump()
            self.pos = std.pos

class STSH(DOCDirStream):
    """The STSH structure specifies the stylesheet for a document."""
    def __init__(self, bytes, mainStream, offset, size):
        DOCDirStream.__init__(self, bytes, mainStream=mainStream)
        self.pos = offset
        self.size = size

    def dump(self):
        print '<stsh type="STSH" offset="%d" size="%d bytes">' % (self.pos, self.size)
        self.lpstshi = LPStshi(self.bytes, self.mainStream, self.pos)
        self.lpstshi.dump()
        self.pos = self.lpstshi.pos
        for i in range(self.lpstshi.stshi.stshif.cstd):
            print '<rglpstd index="%d" type="LPStd" offset="%d">' % (i, self.pos)
            lpstd = LPStd(self)
            lpstd.dump()
            self.pos = lpstd.pos
            print '</rglpstd>'
        print '</stsh>'

# vim:set filetype=python shiftwidth=4 softtabstop=4 expandtab:
