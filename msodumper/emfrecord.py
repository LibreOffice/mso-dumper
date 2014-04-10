#!/usr/bin/env python2
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
#

from docdirstream import DOCDirStream
import wmfrecord

# The FormatSignature enumeration defines valuesembedded data in EMF records.
FormatSignature = {
    0x464D4520: "ENHMETA_SIGNATURE",
    0x46535045: "EPS_SIGNATURE"
}


class EMFStream(DOCDirStream):
    def __init__(self, bytes):
        DOCDirStream.__init__(self, bytes)

    def dump(self):
        print '<stream type="EMF" size="%d">' % self.size
        EmrHeader(self).dump()
        print '</stream>'


class EMFRecord(DOCDirStream):
    def __init__(self, parent):
        DOCDirStream.__init__(self, parent.bytes)
        self.parent = parent
        self.pos = parent.pos


class EmrHeader(EMFRecord):
    """The EMR_HEADER record types define the starting points of EMF metafiles."""
    def __init__(self, parent):
        EMFRecord.__init__(self, parent)

    def dump(self):
        print '<emrHeader>'
        self.printAndSet("Type", self.readuInt32())
        self.printAndSet("Size", self.readuInt32(), hexdump=False)
        Header(self).dump()
        if self.Size >= 100:
            HeaderExtension1(self).dump()
        if self.Size >= 108:
            HeaderExtension2(self).dump()
        print '</emrHeader>'


class Header(EMFRecord):
    """The Header object defines the EMF metafile header."""
    def __init__(self, parent):
        EMFRecord.__init__(self, parent)

    def dump(self):
        posOrig = self.pos
        print("<header>")
        wmfrecord.RectL(self, "Bounds").dump()
        wmfrecord.RectL(self, "Frame").dump()
        self.printAndSet("RecordSignature", self.readuInt32(), dict=FormatSignature)
        self.printAndSet("Version", self.readuInt32())
        self.printAndSet("Bytes", self.readuInt32(), hexdump=False)
        self.printAndSet("Records", self.readuInt32(), hexdump=False)
        self.printAndSet("Handles", self.readuInt16(), hexdump=False)
        self.printAndSet("Reserved", self.readuInt16(), hexdump=False)
        self.printAndSet("nDescription", self.readuInt32(), hexdump=False)
        self.printAndSet("offDescription", self.readuInt32(), hexdump=False)
        self.printAndSet("nPalEntries", self.readuInt32(), hexdump=False)
        wmfrecord.SizeL(self, "Device").dump()
        wmfrecord.SizeL(self, "Millimeters").dump()
        print("</header>")
        assert posOrig == self.pos - 80
        self.parent.pos = self.pos


class HeaderExtension1(EMFRecord):
    """The HeaderExtension1 object defines the first extension to the EMF metafile header."""
    def __init__(self, parent):
        EMFRecord.__init__(self, parent)

    def dump(self):
        posOrig = self.pos
        print("<headerExtension1>")
        self.printAndSet("cbPixelFormat", self.readuInt32(), hexdump=False)
        self.printAndSet("offPixelFormat", self.readuInt32(), hexdump=False)
        self.printAndSet("bOpenGL", self.readuInt32())
        print("</headerExtension1>")
        assert posOrig == self.pos - 12
        self.parent.pos = self.pos


class HeaderExtension2(EMFRecord):
    """The HeaderExtension1 object defines the first extension to the EMF metafile header."""
    def __init__(self, parent):
        EMFRecord.__init__(self, parent)

    def dump(self):
        posOrig = self.pos
        print("<headerExtension2>")
        self.printAndSet("MicrometersX", self.readuInt32(), hexdump=False)
        self.printAndSet("MicrometersY", self.readuInt32(), hexdump=False)
        print("</headerExtension2>")
        assert posOrig == self.pos - 8
        self.parent.pos = self.pos

# vim:set filetype=python shiftwidth=4 softtabstop=4 expandtab:
