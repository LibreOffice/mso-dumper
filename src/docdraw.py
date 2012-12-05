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
import msodraw

class OfficeArtDggContainer(DOCDirStream):
    """The OfficeArtDggContainer record type specifies the container for all the OfficeArt file records that contain document-wide data."""
    def __init__(self, officeArtContent, name):
        DOCDirStream.__init__(self, officeArtContent.bytes)
        self.name = name
        self.pos = officeArtContent.pos
        self.officeArtContent = officeArtContent

    def dump(self):
        print '<%s type="OfficeArtDggContainer" offset="%d">' % (self.name, self.pos)
        self.rh = msodraw.RecordHeader(self)
        self.rh.dumpXml(self)
        pos = self.pos
        while (self.rh.recLen - (pos - self.pos)) > 0:
            posOrig = self.pos
            self.pos = pos
            rh = msodraw.RecordHeader(self)
            rh.dumpXml(self)
            self.pos = posOrig
            pos += msodraw.RecordHeader.size
            if rh.recType in recMap:
                if len(recMap[rh.recType]) == 2:
                    child = recMap[rh.recType][0](self, pos)
                    child.dumpXml(self, rh)
                    assert child.pos == pos + rh.recLen
                else:
                    posOrig = self.pos
                    self.pos = pos
                    child = recMap[rh.recType][0](self)
                    child.dumpXml(self, rh)
                    self.pos = posOrig
            else:
                print '<todo what="OfficeArtDggContainer: recType = %s unhandled (size: %d bytes)"/>' % (hex(rh.recType), rh.recLen)
            pos += rh.recLen
        print '</%s>' % self.name
        assert pos == self.pos + self.rh.recLen
        self.officeArtContent.pos = pos

class OfficeArtSpContainer(DOCDirStream):
    """The OfficeArtSpContainer record specifies a shape container."""
    def __init__(self, parent, pos):
        DOCDirStream.__init__(self, parent.bytes)
        self.pos = pos
        self.parent = parent

    def dumpXml(self, compat, rh):
        self.rh = rh
        print '<shape type="OfficeArtSpContainer">'
        pos = self.pos
        while (self.rh.recLen - (pos - self.pos)) > 0:
            posOrig = self.pos
            self.pos = pos
            rh = msodraw.RecordHeader(self)
            rh.dumpXml(self)
            self.pos = posOrig
            pos += msodraw.RecordHeader.size
            if rh.recType in recMap:
                if len(recMap[rh.recType]) == 2:
                    child = recMap[rh.recType][0](self, pos)
                    child.dumpXml(self, rh)
                    assert child.pos == pos + rh.recLen
                else:
                    posOrig = self.pos
                    self.pos = pos
                    child = recMap[rh.recType][0](self)
                    child.dumpXml(self, rh)
                    self.pos = posOrig
            else:
                print '<todo what="OfficeArtSpContainer: recType = %s unhandled (size: %d bytes)"/>' % (hex(rh.recType), rh.recLen)
            pos += rh.recLen
        print '</shape>'
        assert pos == self.pos + self.rh.recLen
        self.pos = pos

class OfficeArtSpgrContainer(DOCDirStream):
    """The OfficeArtSpgrContainer record specifies a container for groups of shapes."""
    def __init__(self, officeArtDgContainer, pos):
        DOCDirStream.__init__(self, officeArtDgContainer.bytes)
        self.pos = pos
        self.officeArtDgContainer = officeArtDgContainer

    def dumpXml(self, compat, rh):
        self.rh = rh
        print '<groupShape type="OfficeArtSpgrContainer" offset="%d">' % self.pos
        pos = self.pos
        while (self.rh.recLen - (pos - self.pos)) > 0:
            posOrig = self.pos
            self.pos = pos
            rh = msodraw.RecordHeader(self)
            rh.dumpXml(self)
            self.pos = posOrig
            pos += msodraw.RecordHeader.size
            if rh.recType in recMap:
                if len(recMap[rh.recType]) == 2:
                    child = recMap[rh.recType][0](self, pos)
                    child.dumpXml(self, rh)
                    assert child.pos == pos + rh.recLen
                else:
                    posOrig = self.pos
                    self.pos = pos
                    child = recMap[rh.recType][0](self)
                    child.dumpXml(self, rh)
                    self.pos = posOrig
            else:
                print '<todo what="OfficeArtSpgrContainer: recType = %s unhandled (size: %d bytes)"/>' % (hex(rh.recType), rh.recLen)
            pos += rh.recLen
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
        self.rh = msodraw.RecordHeader(self)
        self.rh.dumpXml(self)
        pos = self.pos
        while (self.rh.recLen - (pos - self.pos)) > 0:
            posOrig = self.pos
            self.pos = pos
            rh = msodraw.RecordHeader(self)
            rh.dumpXml(self)
            self.pos = posOrig
            pos += msodraw.RecordHeader.size
            if rh.recType in recMap:
                if len(recMap[rh.recType]) == 2:
                    child = recMap[rh.recType][0](self, pos)
                    child.dumpXml(self, rh)
                    assert child.pos == pos + rh.recLen
                else:
                    posOrig = self.pos
                    self.pos = pos
                    child = recMap[rh.recType][0](self)
                    child.dumpXml(self, rh)
                    self.pos = posOrig
            else:
                print '<todo what="OfficeArtDgContainer: recType = %s unhandled (size: %d bytes)"/>' % (hex(rh.recType), rh.recLen)
            pos += rh.recLen
        print '</%s>' % self.name
        assert pos == self.pos + self.rh.recLen
        self.officeArtContent.pos = pos

recMap = {
        0xf003: [OfficeArtSpgrContainer, True],
        0xf004: [OfficeArtSpContainer, True],
        0xf006: [msodraw.FDGGBlock],
        0xf008: [msodraw.FDG],
        0xf009: [msodraw.FSPGR],
        0xf00a: [msodraw.FSP],
        0xf00b: [msodraw.FOPT],
        0xf011: [msodraw.FClientData],
        0xf11e: [msodraw.SplitMenuColorContainer],
        }

# vim:set filetype=python shiftwidth=4 softtabstop=4 expandtab:
