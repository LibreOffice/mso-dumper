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
            print "TODO FcCompressed: fCompressed = 0 not supported"

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

class PlcPcd(DOCDirStream):
    """The PlcPcd structure is a PLC whose data elements are Pcds (8 bytes each)."""
    def __init__(self, bytes, mainStream, offset, size):
        DOCDirStream.__init__(self, bytes, mainStream=mainStream)
        self.pos = offset
        self.size = size

    def dump(self):
        print '<plcPcd type="PlcPcd" offset="%d" size="%d bytes">' % (self.pos, self.size)
        elements = (self.size - 4) / (4 + 8) # 8 is defined by 2.8.35, the rest is defined by 2.2.2
        pos = self.pos
        for i in range(elements):
            # aCp
            start = struct.unpack("<I", self.bytes[pos:pos+4])[0]
            end = struct.unpack("<I", self.bytes[pos+4:pos+8])[0]
            print '<aCP index="%d" start="%d" end="%d">' % (i, start, end)
            pos += 4

            # aPcd
            offset = self.pos + ( 4 * ( elements + 1 ) ) + ( 8 * i ) # 8 as defined by 2.8.35
            aPcd = Pcd(self.bytes, self.mainStream, offset, 8)
            aPcd.dump()

            offset = aPcd.fc.getTransformedAddress()
            print '<transformed value="%s"/>' % globals.encodeName(self.mainStream.bytes[offset:offset+end-start])
            print '</aCP>'
        print '</plcPcd>'

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
