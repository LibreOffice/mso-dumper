#!/usr/bin/env python2
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
#

from docdirstream import DOCDirStream


# The BrushStyle Enumeration specifies the different possible brush types that can be used in graphics operations.
BrushStyle = {
    0x0000: "BS_SOLID",
    0x0001: "BS_NULL",
    0x0002: "BS_HATCHED",
    0x0003: "BS_PATTERN",
    0x0004: "BS_INDEXED",
    0x0005: "BS_DIBPATTERN",
    0x0006: "BS_DIBPATTERNPT",
    0x0007: "BS_PATTERN8X8",
    0x0008: "BS_DIBPATTERN8X8",
    0x0009: "BS_MONOPATTERN"
}


# The HatchStyle Enumeration specifies the hatch pattern.
HatchStyle = {
    0x0000: "HS_HORIZONTAL",
    0x0001: "HS_VERTICAL",
    0x0002: "HS_FDIAGONAL",
    0x0003: "HS_BDIAGONAL",
    0x0004: "HS_CROSS",
    0x0005: "HS_DIAGCROSS"
}


class WMFRecord(DOCDirStream):
    def __init__(self, parent):
        DOCDirStream.__init__(self, parent.bytes)
        self.parent = parent
        self.pos = parent.pos


class RectL(WMFRecord):
    """The RectL Object defines a rectangle."""
    def __init__(self, parent, name=None):
        WMFRecord.__init__(self, parent)
        if name:
            self.name = name
        else:
            self.name = "rectL"

    def dump(self):
        print '<%s type="RectL">' % self.name
        self.printAndSet("Left", self.readInt32(), hexdump=False)
        self.printAndSet("Top", self.readInt32(), hexdump=False)
        self.printAndSet("Right", self.readInt32(), hexdump=False)
        self.printAndSet("Bottom", self.readInt32(), hexdump=False)
        print '</%s>' % self.name
        self.parent.pos = self.pos


class SizeL(WMFRecord):
    """The SizeL Object defines a rectangle."""
    def __init__(self, parent, name=None):
        WMFRecord.__init__(self, parent)
        if name:
            self.name = name
        else:
            self.name = "sizeL"

    def dump(self):
        print '<%s type="SizeL">' % self.name
        self.printAndSet("cx", self.readuInt32(), hexdump=False)
        self.printAndSet("cy", self.readuInt32(), hexdump=False)
        print '</%s>' % self.name
        self.parent.pos = self.pos


class PointL(WMFRecord):
    """The PointL Object defines the coordinates of a point."""
    def __init__(self, parent, name=None):
        WMFRecord.__init__(self, parent)
        if name:
            self.name = name
        else:
            self.name = "pointL"

    def dump(self):
        print '<%s type="PointL">' % self.name
        self.printAndSet("x", self.readInt32(), hexdump=False)
        self.printAndSet("y", self.readInt32(), hexdump=False)
        print '</%s>' % self.name
        self.parent.pos = self.pos


class PointS(WMFRecord):
    """The PointS Object defines the x- and y-coordinates of a point."""
    def __init__(self, parent, name):
        WMFRecord.__init__(self, parent)
        self.name = name

    def dump(self):
        print '<%s type="PointS">' % self.name
        self.printAndSet("x", self.readInt16(), hexdump=False)
        self.printAndSet("y", self.readInt16(), hexdump=False)
        print '</%s>' % self.name
        self.parent.pos = self.pos


class ColorRef(WMFRecord):
    """The ColorRef Object defines the RGB color."""
    def __init__(self, parent, name):
        WMFRecord.__init__(self, parent)
        self.name = name

    def dump(self):
        print '<%s type="ColorRef">' % self.name
        self.printAndSet("Red", self.readuInt8(), hexdump=False)
        self.printAndSet("Green", self.readuInt8(), hexdump=False)
        self.printAndSet("Blue", self.readuInt8(), hexdump=False)
        self.printAndSet("Reserved", self.readuInt8(), hexdump=False)
        print '</%s>' % self.name
        self.parent.pos = self.pos

# vim:set filetype=python shiftwidth=4 softtabstop=4 expandtab:
