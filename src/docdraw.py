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
    def __init__(self, parent, name):
        DOCDirStream.__init__(self, parent.bytes)
        self.name = name
        self.pos = parent.pos
        self.parent = parent

    def dump(self):
        print '<%s type="OfficeArtRecordHeader" offset="%d" size="%d bytes">' % (self.name, self.pos, OfficeArtRecordHeader.size)
        buf = self.readuInt16()
        self.printAndSet("recVer", buf & 0x000f) # 1..4th bits
        self.printAndSet("recInstance", buf & 0xfff0) # 5..16th bits

        self.printAndSet("recType", self.readuInt16())
        self.printAndSet("recLen", self.readuInt32())
        print '</%s>' % self.name
        assert self.pos == self.parent.pos + OfficeArtRecordHeader.size
        self.parent.pos = self.pos

class OfficeArtFDGG(DOCDirStream):
    """The OfficeArtFDGG record specifies documnet-wide information about all of the drawings that have been saved in the file."""
    size = 16
    def __init__(self, officeArtFDGGBlock, name):
        DOCDirStream.__init__(self, officeArtFDGGBlock.bytes)
        self.name = name
        self.pos = officeArtFDGGBlock.pos
        self.officeArtFDGGBlock = officeArtFDGGBlock

    def dump(self):
        print '<%s type="OfficeArtFDGG" offset="%d" size="%d bytes">' % (self.name, self.pos, OfficeArtFDGG.size)
        self.printAndSet("spidMax", self.readuInt32()) # TODO dump MSOSPID
        self.printAndSet("cidcl", self.readuInt32())
        self.printAndSet("cspSaved", self.readuInt32())
        self.printAndSet("cdgSaved", self.readuInt32())
        print '</%s>' % self.name
        assert self.pos == self.officeArtFDGGBlock.pos + OfficeArtFDGG.size
        self.officeArtFDGGBlock.pos = self.pos

class OfficeArtIDCL(DOCDirStream):
    """The OfficeArtIDCL record specifies a file identifier cluster, which is used to group shape identifiers within a drawing."""
    def __init__(self, officeArtFDGGBlock):
        DOCDirStream.__init__(self, officeArtFDGGBlock.bytes)
        self.pos = officeArtFDGGBlock.pos
        self.officeArtFDGGBlock = officeArtFDGGBlock

    def dump(self):
        print '<officeArtIDCL type="OfficeArtIDCL" pos="%d">' % self.pos
        self.printAndSet("dgid", self.readuInt32()) # TODO dump MSODGID
        self.printAndSet("cspidCur", self.readuInt32())
        print '</officeArtIDCL>'
        self.officeArtFDGGBlock.pos = self.pos

class OfficeArtFDGGBlock(DOCDirStream):
    """The OfficeArtFDGGBlock record specifies document-wide information about all of the drawings that have been saved in the file."""
    def __init__(self, officeArtDggContainer, name):
        DOCDirStream.__init__(self, officeArtDggContainer.bytes)
        self.name = name
        self.pos = officeArtDggContainer.pos
        self.officeArtDggContainer = officeArtDggContainer

    def dump(self):
        print '<%s type="OfficeArtFDGGBlock" offset="%d">' % (self.name, self.pos)
        OfficeArtRecordHeader(self, "rh").dump()
        self.head = OfficeArtFDGG(self, "head")
        self.head.dump()
        for i in range(self.head.cidcl - 1):
            print '<Rgidcl index="%d">' % i
            OfficeArtIDCL(self).dump()
            print '</Rgidcl>'
        print '</%s>' % self.name
        self.officeArtDggContainer.pos = self.pos

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
        OfficeArtFDGGBlock(self, "drawingGroup").dump()
        print '</%s>' % self.name

# vim:set filetype=python shiftwidth=4 softtabstop=4 expandtab:
