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

class OfficeArtContainer(DOCDirStream):
    def __init__(self, parent, name, type):
        DOCDirStream.__init__(self, parent.bytes)
        self.name = name
        self.type = type
        self.pos = parent.pos
        self.parent = parent

    def dumpXml(self):
        print '<%s type="%s" offset="%d">' % (self.name, self.type, self.pos)
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
                print '<todo what="%s: recType = %s unhandled (size: %d bytes)"/>' % (self.type, hex(rh.recType), rh.recLen)
            pos += rh.recLen
        print '</%s>' % self.name
        assert pos == self.pos + self.rh.recLen
        self.parent.pos = pos

class OfficeArtDggContainer(OfficeArtContainer):
    """The OfficeArtDggContainer record type specifies the container for all the OfficeArt file records that contain document-wide data."""
    def __init__(self, officeArtContent, name):
        OfficeArtContainer.__init__(self, officeArtContent, name, "OfficeArtDggContainer")

class OfficeArtDgContainer(OfficeArtContainer):
    """The OfficeArtDgContainer record specifies the container for all the file records for the objects in a drawing."""
    def __init__(self, officeArtContent, name):
        OfficeArtContainer.__init__(self, officeArtContent, name, "OfficeArtDgContainer")

class OfficeArtContainedContainer(DOCDirStream):
    def __init__(self, parent, pos, name, type):
        DOCDirStream.__init__(self, parent.bytes)
        self.pos = pos
        self.name = name
        self.type = type
        self.parent = parent

    def dumpXml(self, compat, rh):
        self.rh = rh
        print '<%s type="OfficeArtSpContainer">' % self.name
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
                print '<todo what="%s: recType = %s unhandled (size: %d bytes)"/>' % (self.type, hex(rh.recType), rh.recLen)
            pos += rh.recLen
        print '</%s>' % self.name
        assert pos == self.pos + self.rh.recLen
        self.pos = pos

class OfficeArtSpContainer(OfficeArtContainedContainer):
    """The OfficeArtSpContainer record specifies a shape container."""
    def __init__(self, parent, pos):
        OfficeArtContainedContainer.__init__(self, parent, pos, "shape", "OfficeArtSpContainer")

class OfficeArtSpgrContainer(OfficeArtContainedContainer):
    """The OfficeArtSpgrContainer record specifies a container for groups of shapes."""
    def __init__(self, officeArtDgContainer, pos):
        OfficeArtContainedContainer.__init__(self, officeArtDgContainer, pos, "groupShape", "OfficeArtSpgrContainer")

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
