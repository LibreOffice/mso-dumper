#!/usr/bin/env python3
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
#

from . import ole
from .binarystream import BinaryStream
from .msometa import SummaryInformationStream
from .msometa import DocumentSummaryInformationStream


class VSDFile:
    """Represents the whole visio file - feed will all bytes."""
    def __init__(self, chars, params):
        self.chars = chars
        self.size = len(self.chars)
        self.params = params
        self.error = None

        self.init()

    def init(self):
        self.header = ole.Header(self.chars, self.params)
        self.pos = self.header.parse()

    def __getDirectoryObj(self):
        obj = self.header.getDirectory()
        obj.parseDirEntries()
        return obj

    def getDirectoryNames(self):
        return self.__getDirectoryObj().getDirectoryNames()

    def getDirectoryStreamByName(self, name):
        obj = self.__getDirectoryObj()
        bytes = obj.getRawStreamByName(name)
        return self.getStreamFromBytes(name, bytes)

    def getStreamFromBytes(self, name, bytes):
        if name == "\x05SummaryInformation":
            return SummaryInformationStream(bytes, self.params, doc=self)
        elif name == "\x05DocumentSummaryInformation":
            return DocumentSummaryInformationStream(bytes, self.params, doc=self)
        else:
            return BinaryStream(bytes, self.params, name, doc=self)

    def getName(self):
        return "native"


def createVSDFile(chars, params):
    return VSDFile(chars, params)

# vim:set filetype=python shiftwidth=4 softtabstop=4 expandtab:
