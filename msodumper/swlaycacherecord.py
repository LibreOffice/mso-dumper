#!/usr/bin/env python3
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
#

from .binarystream import BinaryStream


class SwLayCacheStream(BinaryStream):
    def __init__(self, bytes):
        BinaryStream.__init__(self, bytes)

    def dump(self):
        print('<stream type="SwLayCache" size="%d">' % self.size)
        posOrig = self.pos
        header = Header(self)
        header.dump()

        while posOrig + self.size > self.pos:
            record = CacheRecord(self)
            record.dump()
        print('</stream>')


class Record(BinaryStream):
    def __init__(self, parent):
        BinaryStream.__init__(self, parent.bytes)
        self.parent = parent
        self.pos = parent.pos


class Header(Record):
    def __init__(self, parent):
        Record.__init__(self, parent)

    def dump(self):
        self.printAndSet("nMajorVersion", self.readuInt16())
        self.printAndSet("nMinorVersion", self.readuInt16())
        self.parent.pos = self.pos


class CacheRecord(Record):
    def __init__(self, parent):
        Record.__init__(self, parent)

    def dump(self):
        val = self.readuInt32()
        print('<record type="' + RecordType[val & 0xff] + '">')
        self.printAndSet("cRecTyp", val & 0xff)  # 1..8th bits
        self.printAndSet("nSize", val >> 8)  # 9th..32th bits
        if self.cRecTyp == 0x70:  # SW_LAYCACHE_IO_REC_PAGES
            self.printAndSet("cFlags", self.readuInt8())
        elif self.cRecTyp == 0x46:  # SW_LAYCACHE_IO_REC_FLY
            self.printAndSet("cFlags", self.readuInt8())
            self.printAndSet("nPgNum", self.readuInt16(), hexdump=False)
            self.printAndSet("nIndex", self.readuInt32(), hexdump=False)
            self.printAndSet("nX", self.readuInt32(), hexdump=False)
            self.printAndSet("nY", self.readuInt32(), hexdump=False)
            self.printAndSet("nW", self.readuInt32(), hexdump=False)
            self.printAndSet("nH", self.readuInt32(), hexdump=False)
        elif self.cRecTyp == 0x50:  # SW_LAYCACHE_IO_REC_PARA
            self.printAndSet("cFlags", self.readuInt8())
            self.printAndSet("nIndex", self.readuInt32(), hexdump=False)
            if self.cFlags & 0x01:
                print('<todo what="CacheRecord::dump: unhandled cRecTyp == SW_LAYCACHE_IO_REC_PARA && cFalgs == 1/>' % hex(self.cRecTyp))
        else:
            print('<todo what="CacheRecord::dump: unhandled cRecTyp=%s"/>' % hex(self.cRecTyp))
            self.pos += self.nSize - 4  # 'val' is already read
        print('</record>')
        self.parent.pos = self.pos


RecordType = {
    0x70: 'SW_LAYCACHE_IO_REC_PAGES',  # 'p'
    0x46: 'SW_LAYCACHE_IO_REC_FLY',  # 'F'
    0x50: 'SW_LAYCACHE_IO_REC_PARA',  # 'P'
}

# vim:set filetype=python shiftwidth=4 softtabstop=4 expandtab:
