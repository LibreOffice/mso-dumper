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

class OfficeArtFDG(DOCDirStream):
    """The OfficeArtFDG record specifies the number of shapes, the drawing identifier, and the shape identifier of the last shape in a drawing."""
    def __init__(self, officeArtDgContainer, pos):
        DOCDirStream.__init__(self, officeArtDgContainer.bytes)
        self.pos = pos

    def dumpXml(self, compat, rh):
        self.rh = rh
        print '<drawingData type="OfficeArtFDG" offset="%d">' % self.pos
        self.printAndSet("csp", self.readuInt32())
        self.printAndSet("spidCur", self.readuInt32())
        print '</drawingData>'

class OfficeArtFSPGR(DOCDirStream):
    """The OfficeArtFSPGR record specifies the coordinate system of the group shape that the anchors of the child shape are expressed in."""
    def __init__(self, officeArtSpContainer, pos):
        DOCDirStream.__init__(self, officeArtSpContainer.bytes)
        self.pos = pos
        self.officeArtSpContainer = officeArtSpContainer

    def dumpXml(self, compat, rh):
        self.rh = rh
        print '<shapeGroup type="OfficeArtFSPGR" offset="%d">' % (self.pos)
        pos = self.pos
        self.printAndSet("xLeft", self.readuInt32())
        self.printAndSet("yTop", self.readuInt32())
        self.printAndSet("xRight", self.readuInt32())
        self.printAndSet("yBottom", self.readuInt32())
        print '</shapeGroup>'
        assert self.pos == pos + self.rh.recLen

class OfficeArtFSP(DOCDirStream):
    """The OfficeArtFSP record specifies an instance of a shape."""
    def __init__(self, officeArtSpContainer, pos):
        DOCDirStream.__init__(self, officeArtSpContainer.bytes)
        self.pos = pos
        self.officeArtSpContainer = officeArtSpContainer

    def dumpXml(self, compat, rh):
        self.rh = rh
        print '<shapeProp type="OfficeArtFSP" offset="%d">' % (self.pos)
        pos = self.pos
        self.printAndSet("spid", self.readuInt32())

        buf = self.readuInt32()
        self.printAndSet("fGroup", self.getBit(buf, 0))
        self.printAndSet("fChild", self.getBit(buf, 1))
        self.printAndSet("fPatriarch", self.getBit(buf, 2))
        self.printAndSet("fDeleted", self.getBit(buf, 3))
        self.printAndSet("fOleShape", self.getBit(buf, 4))
        self.printAndSet("fHaveMaster", self.getBit(buf, 5))
        self.printAndSet("fFlipH", self.getBit(buf, 6))
        self.printAndSet("fFlipV", self.getBit(buf, 7))
        self.printAndSet("fConnector", self.getBit(buf, 8))
        self.printAndSet("fHaveAnchor", self.getBit(buf, 9))
        self.printAndSet("fBackground", self.getBit(buf, 10))
        self.printAndSet("fHaveSpt", self.getBit(buf, 11))
        self.printAndSet("unused1", (buf & 0xfffff000) >> 12) # 13..32th bits

        print '</shapeProp>'
        assert self.pos == pos + self.rh.recLen

class OfficeArtClientData(DOCDirStream):
    def __init__(self, officeArtSpContainer, pos):
        DOCDirStream.__init__(self, officeArtSpContainer.bytes)
        self.pos = pos
        self.officeArtSpContainer = officeArtSpContainer

    def dumpXml(self, compat, rh):
        self.rh = rh
        print '<clientData type="OfficeArtClientData" offset="%d">' % self.pos
        pos = self.pos
        self.printAndSet("data", self.readuInt32())
        print '</clientData>'
        assert self.pos == pos + self.rh.recLen

class OfficeArtFOPTEOPID(DOCDirStream):
    """The OfficeArtFOPTEOPID record specifies the header for an entry in a property table."""
    def __init__(self, parent):
        DOCDirStream.__init__(self, parent.bytes)
        self.pos = parent.pos
        self.parent = parent

    def dump(self):
        buf = self.readuInt16()
        self.printAndSet("opid", buf & 0x3fff) # 1..14th bits
        self.printAndSet("fBid", self.getBit(buf, 14))
        self.printAndSet("fComplex", self.getBit(buf, 15))
        self.parent.pos = self.pos

class OfficeArtFOPTE(DOCDirStream):
    """The OfficeArtFOPTE record specifies an entry in a property table."""
    def __init__(self, parent):
        DOCDirStream.__init__(self, parent.bytes)
        self.pos = parent.pos
        self.parent = parent

    def dump(self):
        print '<opid>'
        self.opid = OfficeArtFOPTEOPID(self)
        self.opid.dump()
        print '</opid>'
        self.printAndSet("op", self.readInt32())
        self.parent.pos = self.pos

class OfficeArtRGFOPTE(DOCDirStream):
    """The OfficeArtRGFOPTE record specifies a property table."""
    def __init__(self, parent, name):
        DOCDirStream.__init__(self, parent.bytes)
        self.pos = parent.pos
        self.name = name
        self.parent = parent

    def dump(self):
        print '<%s type="OfficeArtRGFOPTE" offset="%d">' % (self.name, self.pos)
        for i in range(self.parent.rh.recInstance):
            print '<rgfopte index="%d" offset="%d">' % (i, self.pos)
            entry = OfficeArtFOPTE(self)
            entry.dump()
            if entry.opid.fComplex:
                print '<todo what="OfficeArtRGFOPTE: fComplex != 0 unhandled"/>'
            print '</rgfopte>'
        print '</%s>' % self.name
        self.parent.pos = self.pos

class OfficeArtFOPT(DOCDirStream):
    """The OfficeArtFOPT record specifies a table of OfficeArtRGFOPTE properties."""
    def __init__(self, officeArtSpContainer, pos):
        DOCDirStream.__init__(self, officeArtSpContainer.bytes)
        self.pos = pos
        self.officeArtSpContainer = officeArtSpContainer

    def dumpXml(self, compat, rh):
        self.rh = rh
        print '<shapePrimaryOptions type="OfficeArtFOPT" offset="%d">' % self.pos
        pos = self.pos
        OfficeArtRGFOPTE(self, "fopt").dump()
        print '</shapePrimaryOptions>'
        assert self.pos == pos + self.rh.recLen

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
        0xf008: [OfficeArtFDG, True],
        0xf009: [OfficeArtFSPGR, True],
        0xf00a: [OfficeArtFSP, True],
        0xf00b: [OfficeArtFOPT, True],
        0xf011: [OfficeArtClientData, True],
        0xf11e: [msodraw.SplitMenuColorContainer],
        }

# vim:set filetype=python shiftwidth=4 softtabstop=4 expandtab:
