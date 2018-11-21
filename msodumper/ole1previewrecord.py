#!/usr/bin/env python3
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
#

from . import globals
from .binarystream import BinaryStream


class Ole1PreviewStream(BinaryStream):
    def __init__(self, bytes):
        BinaryStream.__init__(self, bytes)

    def dump(self):
        print('<stream type="Ole1Preview" size="%d">' % self.size)
        header = StandardPresentationObject(self, "Header")
        header.dump()
        self.printAndSet("PresentationDataSize", self.readuInt32(), hexdump=False, offset=True)
        self.printAndSet("Reserved1", self.readuInt16())
        self.printAndSet("Reserved2", self.readuInt16())
        self.printAndSet("Reserved3", self.readuInt16())
        self.printAndSet("Reserved4", self.readuInt16())
        print('<PresentationData offset="%s" size="%s"/>' % (self.pos, int(self.PresentationDataSize) - 8))
        print('</stream>')


class Record(BinaryStream):
    def __init__(self, parent):
        BinaryStream.__init__(self, parent.bytes)
        self.parent = parent
        self.pos = parent.pos


class LengthPrefixedAnsiString(Record):
    """Specified by [MS-OLEDS] 2.1.4, specifies a length-prefixed and
    null-terminated ANSI string."""
    def __init__(self, parent, name):
        Record.__init__(self, parent)
        self.parent = parent
        self.pos = parent.pos
        self.name = name

    def dump(self):
        print('<%s type="LengthPrefixedAnsiString">' % self.name)
        self.printAndSet("Length", self.readuInt32(), offset=True)
        bytes = []
        for dummy in range(self.Length):
            c = self.readuInt8()
            bytes.append(c)

        self.printAndSet("String", globals.encodeName("".join(map(lambda c: chr(c), bytes[:-1])), lowOnly=True).encode('utf-8'), hexdump=False, offset=True)

        print('</%s>' % self.name)
        self.parent.pos = self.pos


class StandardPresentationObject(Record):
    def __init__(self, parent, name):
        Record.__init__(self, parent)
        self.name = name

    def dump(self):
        print('<%s type="StandardPresentationObject">' % self.name)
        header = PresentationObjectHeader(self, "Header")
        header.dump()
        self.printAndSet("Width", self.readuInt32())
        self.printAndSet("Height", self.readInt32() * -1)

        print('</%s>' % self.name)
        self.parent.pos = self.pos


class PresentationObjectHeader(Record):
    def __init__(self, parent, name):
        Record.__init__(self, parent)
        self.name = name

    def dump(self):
        print('<%s type="PresentationObjectHeader">' % self.name)
        self.printAndSet("OLEVersion", self.readuInt32())
        self.printAndSet("FormatID", self.readuInt32())
        LengthPrefixedAnsiString(self, "ClassName").dump()

        print('</%s>' % self.name)
        self.parent.pos = self.pos

# vim:set filetype=python shiftwidth=4 softtabstop=4 expandtab:
