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

class OfficeArtRecordHeader(DOCDirStream):
    """The OfficeArtRecordHeader record specifies the common record header for all the OfficeArt records."""
    size = 8
    def __init__(self, officeArtDggContainer, name):
        DOCDirStream.__init__(self, officeArtDggContainer.bytes)
        self.name = name
        self.pos = officeArtDggContainer.pos
        self.officeArtDggContainer = officeArtDggContainer

    def dump(self):
        print '<%s type="OfficeArtRecordHeader" offset="%d" size="%d bytes">' % (self.name, self.pos, OfficeArtRecordHeader.size)
        buf = self.readuInt16()
        self.printAndSet("recVer", buf & 0x000f) # 1..4th bits
        self.printAndSet("recInstance", buf & 0xfff0) # 5..16th bits

        self.printAndSet("recType", self.readuInt16())
        self.printAndSet("recLen", self.readuInt32())
        assert self.pos == self.officeArtDggContainer.pos + OfficeArtRecordHeader.size
        print '</%s>' % self.name

class OfficeArtDggContainer(DOCDirStream):
    """The OfficeArtDggContainer record type specifies the container for all the OfficeArt file records that contain document-wide data."""
    def __init__(self, officeArtContent, name):
        DOCDirStream.__init__(self, officeArtContent.bytes)
        self.name = name
        self.pos = officeArtContent.pos
        self.officeArtContent = officeArtContent

    def dump(self):
        print '<%s type="OfficeArtDggContainer" offset="%d">' % (self.name, self.pos)
        OfficeArtRecordHeader(self, "rh").dump()
        # TODO OfficeArtFDGGBlock(self, "drawingGroup").dump()
        print '</%s>' % self.name

# vim:set filetype=python shiftwidth=4 softtabstop=4 expandtab:
