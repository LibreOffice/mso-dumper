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
    def __init__(self, parent, name, type, contained):
        DOCDirStream.__init__(self, parent.bytes)
        self.name = name
        self.type = type
        self.contained = contained
        self.pos = parent.pos
        self.parent = parent

    def dumpXml(self, recHdl, rh = None):
        recHdl.appendLine('<%s type="%s">' % (self.name, self.type))
        if self.contained:
            self.rh = rh
        else:
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
                posOrig = self.pos
                self.pos = pos
                child = recMap[rh.recType](self)
                child.dumpXml(self, rh)
                self.pos = posOrig
            else:
                recHdl.appendLine('<todo what="%s: recType = %s unhandled (size: %d bytes)"/>' % (self.type, hex(rh.recType), rh.recLen))
            pos += rh.recLen
        recHdl.appendLine('</%s>' % self.name)
        assert pos == self.pos + self.rh.recLen
        self.parent.pos = pos

class OfficeArtDggContainer(OfficeArtContainer):
    """The OfficeArtDggContainer record type specifies the container for all the OfficeArt file records that contain document-wide data."""
    def __init__(self, officeArtContent, name):
        OfficeArtContainer.__init__(self, officeArtContent, name, "OfficeArtDggContainer", False)

class OfficeArtDgContainer(OfficeArtContainer):
    """The OfficeArtDgContainer record specifies the container for all the file records for the objects in a drawing."""
    def __init__(self, officeArtContent, name):
        OfficeArtContainer.__init__(self, officeArtContent, name, "OfficeArtDgContainer", False)

class OfficeArtSpContainer(OfficeArtContainer):
    """The OfficeArtSpContainer record specifies a shape container."""
    def __init__(self, parent):
        OfficeArtContainer.__init__(self, parent, "shape", "OfficeArtSpContainer", True)

class OfficeArtSpgrContainer(OfficeArtContainer):
    """The OfficeArtSpgrContainer record specifies a container for groups of shapes."""
    def __init__(self, officeArtDgContainer):
        OfficeArtContainer.__init__(self, officeArtDgContainer, "groupShape", "OfficeArtSpgrContainer", True)

recMap = {
        0xf003: OfficeArtSpgrContainer,
        0xf004: OfficeArtSpContainer,
        0xf006: msodraw.FDGGBlock,
        0xf008: msodraw.FDG,
        0xf009: msodraw.FSPGR,
        0xf00a: msodraw.FSP,
        0xf00b: msodraw.FOPT,
        0xf011: msodraw.FClientData,
        0xf11e: msodraw.SplitMenuColorContainer,
        }

# vim:set filetype=python shiftwidth=4 softtabstop=4 expandtab:
