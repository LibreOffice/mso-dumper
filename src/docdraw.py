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
    def __init__(self, parent, name = None, pos = None):
        DOCDirStream.__init__(self, parent.bytes)
        self.name = name
        if pos:
            self.pos = pos
        else:
            self.pos = parent.pos
        self.posOrig = self.pos
        self.parent = parent

    def dump(self, parseOnly = False):
        if not parseOnly:
            print '<%s type="OfficeArtRecordHeader" offset="%d" size="%d bytes">' % (self.name, self.pos, OfficeArtRecordHeader.size)
        buf = self.readuInt16()
        self.printAndSet("recVer", buf & 0x000f, silent = parseOnly) # 1..4th bits
        self.printAndSet("recInstance", (buf & 0xfff0) >> 4, silent = parseOnly) # 5..16th bits

        self.printAndSet("recType", self.readuInt16(), silent = parseOnly)
        self.printAndSet("recLen", self.readuInt32(), silent = parseOnly)
        if not parseOnly:
            print '</%s>' % self.name
        assert self.pos == self.posOrig + OfficeArtRecordHeader.size
        if not parseOnly:
            self.parent.pos = self.pos

    def parse(self):
        self.dump(True)

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
        self.printAndSet("spidMax", self.readuInt32())
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
        self.printAndSet("dgid", self.readuInt32())
        self.printAndSet("cspidCur", self.readuInt32())
        print '</officeArtIDCL>'
        self.officeArtFDGGBlock.pos = self.pos

class OfficeArtFDGGBlock(DOCDirStream):
    """The OfficeArtFDGGBlock record specifies document-wide information about all of the drawings that have been saved in the file."""
    def __init__(self, officeArtDggContainer, pos):
        DOCDirStream.__init__(self, officeArtDggContainer.bytes)
        self.pos = pos

    def dump(self):
        print '<drawingGroup type="OfficeArtFDGGBlock" offset="%d">' % self.pos
        OfficeArtRecordHeader(self, "rh").dump()
        self.head = OfficeArtFDGG(self, "head")
        self.head.dump()
        for i in range(self.head.cidcl - 1):
            print '<Rgidcl index="%d">' % i
            OfficeArtIDCL(self).dump()
            print '</Rgidcl>'
        print '</drawingGroup>'

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
    def __init__(self, officeArtDggContainer, pos):
        DOCDirStream.__init__(self, officeArtDggContainer.bytes)
        self.pos = pos

    def dump(self):
        print '<splitColors type="OfficeArtSplitMenuColorContainer" offset="%d">' % self.pos
        OfficeArtRecordHeader(self, "rh").dump()
        for i in ["fill", "line", "shadow", "3d"]:
            print '<smca type="%s">' % i
            MSOCR(self).dump()
            print '</smca>'
        print '</splitColors>'

class OfficeArtDggContainer(DOCDirStream):
    """The OfficeArtDggContainer record type specifies the container for all the OfficeArt file records that contain document-wide data."""
    def __init__(self, officeArtContent, name):
        DOCDirStream.__init__(self, officeArtContent.bytes)
        self.name = name
        self.pos = officeArtContent.pos
        self.officeArtContent = officeArtContent

    def dump(self):
        print '<%s type="OfficeArtDggContainer" offset="%d">' % (self.name, self.pos)
        self.rh = OfficeArtRecordHeader(self, "rh")
        self.rh.dump()
        pos = self.pos
        recMap = {
                0xf006: OfficeArtFDGGBlock,
                0xf11e: OfficeArtSplitMenuColorContainer,
                }
        while (self.rh.recLen - (pos - self.pos)) > 0:
            rh = OfficeArtRecordHeader(self, pos = pos)
            rh.parse()
            if rh.recType in recMap:
                child = recMap[rh.recType](self, pos)
                child.dump()
                assert child.pos == pos + OfficeArtRecordHeader.size + rh.recLen
            else:
                print '<todo what="OfficeArtDggContainer: recType = %s unhandled (size: %d bytes)"/>' % (hex(rh.recType), rh.recLen)
            pos += OfficeArtRecordHeader.size + rh.recLen
        print '</%s>' % self.name
        assert pos == self.pos + self.rh.recLen
        self.officeArtContent.pos = pos

class OfficeArtFDG(DOCDirStream):
    """The OfficeArtFDG record specifies the number of shapes, the drawing identifier, and the shape identifier of the last shape in a drawing."""
    def __init__(self, officeArtDgContainer, pos):
        DOCDirStream.__init__(self, officeArtDgContainer.bytes)
        self.pos = pos

    def dump(self):
        print '<drawingData type="OfficeArtFDG" offset="%d">' % self.pos
        OfficeArtRecordHeader(self, "rh").dump()
        self.printAndSet("csp", self.readuInt32())
        self.printAndSet("spidCur", self.readuInt32())
        print '</drawingData>'

class OfficeArtSpgrContainer(DOCDirStream):
    """The OfficeArtSpgrContainer record specifies a container for groups of shapes."""
    def __init__(self, officeArtDgContainer, pos):
        DOCDirStream.__init__(self, officeArtDgContainer.bytes)
        self.pos = pos
        self.officeArtDgContainer = officeArtDgContainer

    def dump(self):
        print '<groupShape type="OfficeArtSpgrContainer" offset="%d">' % self.pos
        self.rh = OfficeArtRecordHeader(self, "rh")
        self.rh.dump()
        pos = self.pos
        recMap = {
                }
        while (self.rh.recLen - (pos - self.pos)) > 0:
            rh = OfficeArtRecordHeader(self, pos = pos)
            rh.parse()
            if rh.recType in recMap:
                child = recMap[rh.recType](self, pos)
                child.dump()
                assert child.pos == pos + OfficeArtRecordHeader.size + rh.recLen
            else:
                print '<todo what="OfficeArtSpgrContainer: recType = %s unhandled (size: %d bytes)"/>' % (hex(rh.recType), rh.recLen)
            pos += OfficeArtRecordHeader.size + rh.recLen
        print '</groupShape>'
        assert pos == self.pos + self.rh.recLen
        self.pos = pos

class OfficeArtDgContainer(DOCDirStream):
    """The OfficeArtDgContainer record specifies the container for all the file records for the objects in a drawing."""
    def __init__(self, officeArtContent, name):
        DOCDirStream.__init__(self, officeArtContent.bytes)
        self.name = name
        self.pos = officeArtContent.pos
        self.officeArtContent = officeArtContent

    def dump(self):
        print '<%s type="OfficeArtDgContainer" offset="%d">' % (self.name, self.pos)
        self.rh = OfficeArtRecordHeader(self, "rh")
        self.rh.dump()
        pos = self.pos
        recMap = {
                0xf003: OfficeArtSpgrContainer,
                0xf008: OfficeArtFDG,
                }
        while (self.rh.recLen - (pos - self.pos)) > 0:
            rh = OfficeArtRecordHeader(self, pos = pos)
            rh.parse()
            if rh.recType in recMap:
                child = recMap[rh.recType](self, pos)
                child.dump()
                assert child.pos == pos + OfficeArtRecordHeader.size + rh.recLen
            else:
                print '<todo what="OfficeArtDgContainer: recType = %s unhandled (size: %d bytes)"/>' % (hex(rh.recType), rh.recLen)
            pos += OfficeArtRecordHeader.size + rh.recLen
        print '</%s>' % self.name
        assert pos == self.pos + self.rh.recLen
        self.officeArtContent.pos = pos

# vim:set filetype=python shiftwidth=4 softtabstop=4 expandtab:
