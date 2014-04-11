#!/usr/bin/env python2
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
#

from docdirstream import DOCDirStream


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

# vim:set filetype=python shiftwidth=4 softtabstop=4 expandtab:
