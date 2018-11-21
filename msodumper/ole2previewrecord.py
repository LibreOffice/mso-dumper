#!/usr/bin/env python3
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
#

from . import globals
from .binarystream import BinaryStream


class Ole2PreviewStream(BinaryStream):
    def __init__(self, bytes):
        BinaryStream.__init__(self, bytes)

    def dump(self):
        print('<stream type="Ole2Preview" size="%d">' % self.size)

        ansiClipboardFormat = ClipboardFormatOrAnsiString(self, "AnsiClipboardFormat")
        ansiClipboardFormat.dump()
        self.printAndSet("TargetDeviceSize", self.readuInt32())
        if self.TargetDeviceSize == 0x00000004:
            # TargetDevice is not present
            pass
        else:
            print('<todo what="TargetDeviceSize != 0x00000004"/>')
        self.printAndSet("Aspect", self.readuInt32())
        self.printAndSet("Lindex", self.readuInt32())
        self.printAndSet("Advf", self.readuInt32())
        self.printAndSet("Reserved1", self.readuInt32())
        self.printAndSet("Width", self.readuInt32(), hexdump=False)
        self.printAndSet("Height", self.readuInt32(), hexdump=False)
        self.printAndSet("Size", self.readuInt32(), hexdump=False)
        print('<Data offset="%s" size="%s"/>' % (self.pos, self.Size))

        print('</stream>')


class Record(BinaryStream):
    def __init__(self, parent):
        BinaryStream.__init__(self, parent.bytes)
        self.parent = parent
        self.pos = parent.pos


class ClipboardFormatOrAnsiString(Record):
    def __init__(self, parent, name):
        Record.__init__(self, parent)
        self.parent = parent
        self.pos = parent.pos
        self.name = name

    def dump(self):
        print('<%s type="ClipboardFormatOrAnsiString">' % self.name)

        self.printAndSet("MarkerOrLength", self.readuInt32())
        if self.MarkerOrLength == 0xffffffff:
            self.printAndSet("FormatOrAnsiLength", self.readuInt32(), dict=ClipboardFormats)
        else:
            print('<todo what="MarkerOrLength != 0xffffffff"/>')

        print('</%s>' % self.name)
        self.parent.pos = self.pos


ClipboardFormats = {
    0x00000002: "CF_BITMAP",
    0x00000003: "CF_METAFILEPICT",
    0x00000008: "CF_DIB",
    0x0000000E: "CF_ENHMETAFILE",
}

# vim:set filetype=python shiftwidth=4 softtabstop=4 expandtab:
