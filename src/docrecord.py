#!/usr/bin/env python

import struct
from docdirstream import DOCDirStream
import globals

class FcCompressed(DOCDirStream):
    """The FcCompressed structure specifies the location of text in the WordDocument Stream."""
    def __init__(self, bytes, mainStream, offset, size):
        DOCDirStream.__init__(self, bytes, mainStream=mainStream)
        self.pos = offset
        self.size = size

    def dump(self):
        print '<fcCompressed type="FcCompressed" offset="%d" size="%d bytes">' % (self.pos, self.size)
        buf = struct.unpack("<I", self.bytes[self.pos:self.pos+4])[0]
        self.pos += 4
        self.printAndSet("fc", buf & ((2**32-1) >> 2)) # bits 0..29
        self.printAndSet("fCompressed", self.getBit(buf, 30))
        self.printAndSet("r1", self.getBit(buf, 31))
        print '</fcCompressed>'

    def getTransformedAddress(self):
        if self.fCompressed:
            return self.fc/2
        else:
            print '<todo what="FcCompressed: fCompressed = 0 not supported"/>'

class Pcd(DOCDirStream):
    """The Pcd structure specifies the location of text in the WordDocument Stream and additional properties for this text."""
    def __init__(self, bytes, mainStream, offset, size):
        DOCDirStream.__init__(self, bytes, mainStream=mainStream)
        self.pos = offset
        self.size = size

    def dump(self):
        print '<pcd type="Pcd" offset="%d" size="%d bytes">' % (self.pos, self.size)
        buf = struct.unpack("<H", self.bytes[self.pos:self.pos+2])[0]
        self.pos += 2
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
        return pos + (4 * (self.getElements() + 1)) + (self.structSize * i)

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
            start = struct.unpack("<I", self.bytes[pos:pos+4])[0]
            end = struct.unpack("<I", self.bytes[pos+4:pos+8])[0]
            print '<aCP index="%d" start="%d" end="%d">' % (i, start, end)
            pos += 4

            # aPcd
            aPcd = Pcd(self.bytes, self.mainStream, self.getOffset(self.pos, i), 8)
            aPcd.dump()

            offset = aPcd.fc.getTransformedAddress()
            print '<transformed value="%s"/>' % globals.encodeName(self.mainStream.bytes[offset:offset+end-start])
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
                # 6: variable,
                7: 3,
                }

        self.sprm = struct.unpack("<H", self.bytes[self.pos:self.pos+2])[0]
        self.pos += 2

        self.ispmd = (self.sprm & 0x1ff)        # 1-9th bits
        self.fSpec = (self.sprm & 0x200)  >> 9  # 10th bit
        self.sgc   = (self.sprm & 0x1c00) >> 10 # 11-13th bits
        self.spra  = (self.sprm & 0xe000) >> 13 # 14-16th bits

        if self.operandSizeMap[self.spra] == 1:
            self.operand = ord(struct.unpack("<c", self.bytes[self.pos:self.pos+1])[0])
        elif self.operandSizeMap[self.spra] == 2:
            self.operand = struct.unpack("<H", self.bytes[self.pos:self.pos+2])[0]
        elif self.operandSizeMap[self.spra] == 4:
            self.operand = struct.unpack("<I", self.bytes[self.pos:self.pos+4])[0] # TODO generalize this
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
        print '<sprm value="%s" ispmd="%s" fSpec="%s" sgc="%s" spra="%s" operandSize="%s" operand="%s"/>' % (
                hex(self.sprm), hex(self.ispmd), hex(self.fSpec), sgcmap[self.sgc], hex(self.spra), self.getOperandSize(), hex(self.operand)
                )

    def getOperandSize(self):
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
        self.printAndSet("istd", struct.unpack("<H", self.bytes[self.pos:self.pos+2])[0])
        pos += 2
        while (self.size - (pos - self.pos)) > 0:
            prl = Prl(self.bytes, pos)
            prl.dump()
            pos += prl.getSize()
        print '</grpPrlAndIstd>'

class PapxInFkp(DOCDirStream):
    """The PapxInFkp structure specifies a set of text properties."""
    def __init__(self, bytes, mainStream, offset):
        DOCDirStream.__init__(self, bytes)
        self.pos = offset

    def dump(self):
        print '<papxInFkp type="PapxInFkp" offset="%d">' % self.pos
        self.printAndSet("cb", ord(struct.unpack("<c", self.bytes[self.pos:self.pos+1])[0]))
        self.pos += 1
        if self.cb == 0:
            self.printAndSet("cb_", ord(struct.unpack("<c", self.bytes[self.pos:self.pos+1])[0]))
            self.pos += 1
            grpPrlAndIstd = GrpPrlAndIstd(self.bytes, self.pos, 2 * self.cb_)
            grpPrlAndIstd.dump()
        else:
            print '<todo what="PapxInFkp::dump() first byte is not 0"/>'
        print '</papxInFkp>'
    
class BxPap(DOCDirStream):
    """The BxPap structure specifies the offset of a PapxInFkp in PapxFkp."""
    def __init__(self, bytes, mainStream, offset, size, parentoffset):
        DOCDirStream.__init__(self, bytes)
        self.pos = offset
        self.size = size
        self.parentpos = parentoffset

    def dump(self):
        print '<bxPap type="BxPap" offset="%d" size="%d bytes">' % (self.pos, self.size)
        self.printAndSet("bOffset", ord(struct.unpack("<c", self.bytes[self.pos:self.pos+1])[0]))
        papxInFkp = PapxInFkp(self.bytes, self.mainStream, self.parentpos + self.bOffset*2)
        papxInFkp.dump()
        print '</bxPap>'

class PapxFkp(DOCDirStream):
    """The PapxFkp structure maps paragraphs, table rows, and table cells to their properties."""
    def __init__(self, bytes, mainStream, offset, size):
        DOCDirStream.__init__(self, mainStream.bytes)
        self.pos = offset
        self.size = size

    def dump(self):
        print '<papxFkp type="PapxFkp" offset="%d" size="%d bytes">' % (self.pos, self.size)
        self.cpara = ord(struct.unpack("<c", self.bytes[self.pos+self.size-1:self.pos+self.size-1+1])[0])
        pos = self.pos
        for i in range(self.cpara):
            # rgfc
            start = struct.unpack("<I", self.bytes[pos:pos+4])[0]
            end = struct.unpack("<I", self.bytes[pos+4:pos+8])[0]
            print '<rgfc index="%d" start="%d" end="%d">' % (i, start, end)
            print '<transformed value="%s"/>' % globals.encodeName(self.bytes[start:end])
            pos += 4

            # rgbx
            offset = self.pos + ( 4 * ( self.cpara + 1 ) ) + ( 13 * i ) # TODO, 13 is hardwired here
            bxPap = BxPap(self.bytes, self.mainStream, offset, 13, self.pos) # TODO 13 hardwired
            bxPap.dump()
            print '</rgfc>'

        self.printAndSet("cpara", self.cpara)
        print '</papxFkp>'

class PnFkpPapx(DOCDirStream):
    """The PnFkpPapx structure specifies the offset of a PapxFkp in the WordDocument Stream."""
    def __init__(self, bytes, mainStream, offset, size, name):
        DOCDirStream.__init__(self, bytes, mainStream=mainStream)
        self.pos = offset
        self.size = size
        self.name = name

    def dump(self):
        print '<%s type="PnFkpPapx" offset="%d" size="%d bytes">' % (self.name, self.pos, self.size)
        buf = struct.unpack("<I", self.bytes[self.pos:self.pos+4])[0]
        self.pos += 4
        self.printAndSet("pn", buf & (2**22-1))
        papxFkp = PapxFkp(self.bytes, self.mainStream, self.pn*512, 512)
        papxFkp.dump()
        print '</%s>' % self.name

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
            start = struct.unpack("<I", self.bytes[pos:pos+4])[0]
            end = struct.unpack("<I", self.bytes[pos+4:pos+8])[0]
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
        self.printAndSet("clxt", ord(struct.unpack("<c", self.bytes[self.pos:self.pos+1])[0]))
        self.pos += 1
        self.printAndSet("lcb", struct.unpack("<I", self.bytes[self.pos:self.pos+4])[0])
        self.pos += 4
        PlcPcd(self.bytes, self.mainStream, self.pos, self.lcb).dump()
        print '</pcdt>'

class Clx(DOCDirStream):
    def __init__(self, bytes, mainStream, offset, size):
        DOCDirStream.__init__(self, bytes, mainStream=mainStream)
        self.pos = offset
        self.size = size

    def dump(self):
        print '<clx type="Clx" offset="%d" size="%d bytes">' % (self.pos, self.size)
        firstByte = ord(struct.unpack("<c", self.bytes[self.pos:self.pos+1])[0])
        if firstByte == 0x02:
            print '<info what="Array of Prc, 0 elements"/>'
            Pcdt(self.bytes, self.mainStream, self.pos, self.size).dump()
        else:
            print '<todo what="Clx::dump() first byte is not 0x02"/>'
        print '</clx>'

# vim:set filetype=python shiftwidth=4 softtabstop=4 expandtab:
