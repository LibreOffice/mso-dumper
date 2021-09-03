#!/usr/bin/env python3
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
#

# Simple test to show the timestamp of the root node, which is ideally empty,
# but e.g. https://bugs.freedesktop.org/show_bug.cgi?id=86729 bugdocs have it
# set.

import msodumper.docdirstream
import msodumper.globals
import msodumper.msometa
import sys
sys = reload(sys)
sys.setdefaultencoding("utf-8")
sys.path.append(sys.path[0] + "/..")


class OLEStream(msodumper.docdirstream.DOCDirStream):
    def __init__(self, bytes):
        msodumper.docdirstream.DOCDirStream.__init__(self, bytes)

    def dump(self):
        print('<stream type="OLE" size="%d">' % self.size)
        header = Header(self)
        header.dump()

        # Seek to the Root Directory Entry
        sectorSize = 2 ** header.SectorShift
        self.pos = (header.FirstDirSectorLocation + 1) * sectorSize

        DirectoryEntryName = msodumper.globals.getUTF8FromUTF16(self.readBytes(64))
        print('<DirectoryEntryName value="%s"/>' % DirectoryEntryName)
        DirectoryEntryNameLength = self.readuInt16()
        print('<DirectoryEntryNameLength value="%s"/>' % DirectoryEntryNameLength)
        ObjectType = self.readuInt8()
        print('<ObjectType value="%s"/>' % ObjectType)
        ColorFlag = self.readuInt8()
        print('<ColorFlag value="%s"/>' % ColorFlag)
        LeftSiblingID = self.readuInt32()
        print('<LeftSiblingID value="0x%x"/>' % LeftSiblingID)
        RightSiblingID = self.readuInt32()
        print('<RightSiblingID value="0x%x"/>' % RightSiblingID)
        ChildID = self.readuInt32()
        print('<ChildID value="0x%x"/>' % ChildID)
        msodumper.msometa.GUID(self, "CLSID").dump()
        StateBits = self.readuInt32()
        print('<StateBits value="0x%x"/>' % StateBits)
        msodumper.msometa.FILETIME(self, "CreationTime").dump()
        msodumper.msometa.FILETIME(self, "ModifiedTime").dump()
        print('</stream>')


class Header(msodumper.msometa.OLERecord):
    def __init__(self, parent):
        msodumper.msometa.OLERecord.__init__(self, parent)

    def dump(self):
        print('<CFHeader>')
        self.printAndSet("HeaderSignature", self.readuInt64())
        self.printAndSet("HeaderCLSID0", self.readuInt64())
        self.printAndSet("HeaderCLSID1", self.readuInt64())
        self.printAndSet("MinorVersion", self.readuInt16())
        self.printAndSet("MajorVersion", self.readuInt16())
        self.printAndSet("ByteOrder", self.readuInt16())
        self.printAndSet("SectorShift", self.readuInt16())
        self.printAndSet("MiniSectorShift", self.readuInt16())
        self.printAndSet("Reserved0", self.readuInt16())
        self.printAndSet("Reserved1", self.readuInt16())
        self.printAndSet("Reserved2", self.readuInt16())
        self.printAndSet("NumDirectorySectors", self.readuInt32())
        self.printAndSet("NumFATSectors", self.readuInt32())
        self.printAndSet("FirstDirSectorLocation", self.readuInt32())
        self.printAndSet("TransactionSignatureNumber", self.readuInt32())
        self.printAndSet("MiniStreamCutoffSize", self.readuInt32())
        self.printAndSet("FirstMiniFATSectorLocation", self.readuInt32())
        self.printAndSet("NumMiniFATSectors", self.readuInt32())
        self.printAndSet("FirstDIFATSectorLocation", self.readuInt32())
        self.printAndSet("NumDIFATSectors", self.readuInt32())
        print('<DIFAT>')
        self.DIFAT = []
        for i in range(109):
            n = self.readuInt32()
            if n == 0xffffffff:
                break
            print('<DIFAT index="%d" value="%x"/>' % (i, n))
            self.DIFAT.append(n)
        print('</DIFAT>')
        print('</CFHeader>')


class OLEDumper:
    def __init__(self, filepath):
        self.filepath = filepath

    def dump(self):
        file = open(self.filepath, 'rb')
        strm = OLEStream(file.read())
        file.close()
        print('<?xml version="1.0"?>')
        strm.dump()


def main(args):
    dumper = OLEDumper(args[1])
    dumper.dump()


if __name__ == '__main__':
    main(sys.argv)

# vim:set filetype=python shiftwidth=4 softtabstop=4 expandtab:
