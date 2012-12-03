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
        self.printAndSet("recInstance", (buf & 0xfff0) >> 4) # 5..16th bits

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

class MSOCR(DOCDirStream):
    """The MSOCR record specifies either the RGB color or the scheme color index."""
    size = 4
    def __init__(self, parent):
        DOCDirStream.__init__(self, parent.bytes)
        self.pos = parent.pos
        self.parent = parent

    def dump(self):
        print '<msocr type="MSOCR" offset="%d" size="%d">' % (self.pos, MSOCR.size)
        buf = self.readuInt32()
        self.printAndSet("red",           buf & 0x000000ff) # 1..8th bits
        self.printAndSet("green",        (buf & 0x0000ff00) >> 8) # 9..16th bits
        self.printAndSet("blue",         (buf & 0x00ff0000) >> 16) # 17..24th bits
        self.printAndSet("unused1",      (buf & 0x07000000) >> 24) # 25..27th bits
        self.printAndSet("fSchemeIndex", (buf & 0x08000000) >> 27) # 28th bits
        self.printAndSet("unused2",      (buf & 0xf0000000) >> 28) # 29..32th bits
        print '</msocr>'
        self.parent.pos = self.pos

class OfficeArtSplitMenuColorContainer(DOCDirStream):
    """The OfficeArtSplitMenuColorContainer record specifies a container for the colors that were most recently used to format shapes."""
    def __init__(self, officeArtDggContainer, name):
        DOCDirStream.__init__(self, officeArtDggContainer.bytes)
        self.name = name
        self.pos = officeArtDggContainer.pos
        self.officeArtDggContainer = officeArtDggContainer

    def dump(self):
        print '<%s type="OfficeArtSplitMenuColorContainer" offset="%d">' % (self.name, self.pos)
        OfficeArtRecordHeader(self, "rh").dump()
        self.head = OfficeArtFDGG(self, "head")
        self.head.dump()
        for i in ["fill", "line", "shadow", "3d"]:
            print '<smca type="%s">' % i
            MSOCR(self).dump()
            print '</smca>'
        print '</%s>' % self.name
        self.officeArtDggContainer.pos = self.pos

class OfficeArtDggContainer(DOCDirStream):
    """The OfficeArtDggContainer record type specifies the container for all the OfficeArt file records that contain document-wide data."""
    def __init__(self, officeArtContent, name):
        DOCDirStream.__init__(self, officeArtContent.bytes)
        self.name = name
        self.pos = officeArtContent.pos
        self.officeArtContent = officeArtContent

    def getNextType(self):
        return self.getuInt16(pos = self.pos + 2)

    def dump(self):
        print '<%s type="OfficeArtDggContainer" offset="%d">' % (self.name, self.pos)
        OfficeArtRecordHeader(self, "rh").dump()
        if self.getNextType() == 0xf006:
            OfficeArtFDGGBlock(self, "drawingGroup").dump()
        if self.getNextType() == 0xf001:
            print '<todo what="OfficeArtDggContainer: unhandled OfficeArtBStoreContainer"/>'
        if self.getNextType() == 0xf00b:
            print '<todo what="OfficeArtDggContainer: unhandled OfficeArtFOPT"/>'
        if self.getNextType() == 0xf122:
            print '<todo what="OfficeArtDggContainer: unhandled OfficeArtTertiaryFOPT"/>'
        if self.getNextType() == 0xf11a:
            print '<todo what="OfficeArtDggContainer: unhandled OfficeArtColorMRUContainer"/>'
        if self.getNextType() == 0xf11e:
            OfficeArtSplitMenuColorContainer(self, "splitColors").dump()
        print '</%s>' % self.name

# vim:set filetype=python shiftwidth=4 softtabstop=4 expandtab:
