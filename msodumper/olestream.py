########################################################################
#
#  Copyright (c) 2011 Kohei Yoshida
#
#  Permission is hereby granted, free of charge, to any person
#  obtaining a copy of this software and associated documentation
#  files (the "Software"), to deal in the Software without
#  restriction, including without limitation the rights to use,
#  copy, modify, merge, publish, distribute, sublicense, and/or sell
#  copies of the Software, and to permit persons to whom the
#  Software is furnished to do so, subject to the following
#  conditions:
#
#  The above copyright notice and this permission notice shall be
#  included in all copies or substantial portions of the Software.
#
#  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
#  EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
#  OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
#  NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
#  HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
#  WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
#  FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
#  OTHER DEALINGS IN THE SOFTWARE.
#
########################################################################

import sys
import globals

class CompObjStreamError(Exception): pass

class MonikerStream(object):
    def __init__ (self, bytes):
        self.strm = globals.ByteStream(bytes)

    def read (self):
        print ("moniker size: %d"%(len(self.strm.bytes)-16))
        clsID = self.strm.readBytes(16)
        print ("CLS ID: %s"%globals.getRawBytes(clsID, True, False))
        print ("stream data (implemention specific):")
        globals.dumpBytes(self.strm.readRemainingBytes())
        print ("")

class OLEStream(object):

    def __init__ (self, bytes):
        self.strm = globals.ByteStream(bytes)

    def read (self):
        ver = self.strm.readUnsignedInt(4)
        print ("version: 0x%8.8X"%ver)
        flags = self.strm.readUnsignedInt(4)
        print ("flags: %d"%flags)
        linkUpdateOption = self.strm.readUnsignedInt(4)
        print ("link update option: %d"%linkUpdateOption)
        reserved = self.strm.readUnsignedInt(4)
        print ("")

        # Reserved moniker (must be ignored)
        monikerSize = self.strm.readUnsignedInt(4)
        if monikerSize > 0:
            strm = MonikerStream(self.strm.readBytes(monikerSize+16))
            strm.read()

        # Relative source moniker (relative path to the linked object)
        monikerSize = self.strm.readUnsignedInt(4)
        if monikerSize > 0:
            strm = MonikerStream(self.strm.readBytes(monikerSize+16))
            strm.read()

        # Absolute source moniker (absolute path to the linked object)
        monikerSize = self.strm.readUnsignedInt(4)
        if monikerSize > 0:
            strm = MonikerStream(self.strm.readBytes(monikerSize+16))
            strm.read()

        clsIDIndicator = self.strm.readSignedInt(4)
        print ("cls ID indicator: %d"%clsIDIndicator)
        clsID = self.strm.readBytes(16)
        print ("CLS ID: %s"%globals.getRawBytes(clsID, True, False))
#       globals.dumpBytes(self.strm.bytes, 512)

class CompObjStream(object):

    def __init__ (self, bytes):
        self.strm = globals.ByteStream(bytes)

    def read (self):
        # CompObjHeader
        reserved = self.strm.readBytes(4)
        ver = self.strm.readUnsignedInt(4)
        reserved = self.strm.readBytes(20)
        print ("version: 0x%4.4X"%ver)

        # LengthPrefixedAnsiString
        length = self.strm.readUnsignedInt(4)
        displayName = self.strm.readBytes(length)
        if ord(displayName[-1]) != 0x00:
            # must be null-terminated.
            raise CompObjStreamError()

        print ("display name: " + displayName[:-1])

        # ClipboardFormatOrAnsiString
        marker = self.strm.readUnsignedInt(4)
        if marker == 0:
            # Don't do anything.
            pass
        elif marker == 0xFFFFFFFF or marker == 0xFFFFFFFE:
            clipFormatID = self.strm.readUnsignedInt(4)
            print ("clipboard format ID: %d"%clipFormatID)
        else:
            clipName = self.strm.readBytes(marker)
            if ord(clipName[-1]) != 0x00:
                # must be null-terminated.
                raise CompObjStreamError()
            print ("clipboard format name: %s"%clipName[:-1])

        # LengthPrefixedAnsiString
        length = self.strm.readUnsignedInt(4)
        if length == 0 or length > 0x00000028:
            # the spec says this stream is now invalid.
            raise CompObjStreamError()

        reserved = self.strm.readBytes(length)
        if ord(reserved[-1]) != 0x00:
            # must be null-terminated.
            raise CompObjStreamError()

        print ("reserved name : %s"%reserved[:-1])
        unicodeMarker = self.strm.readUnsignedInt(4)
        if unicodeMarker != 0x71B239F4:
            raise CompObjStreamError()

        # LengthPrefixedUnicodeString
        length = self.strm.readUnsignedInt(4)
        if length > 0:
            s = globals.getUTF8FromUTF16(self.strm.readBytes(length*2))
            print ("display name (unicode): %s"%s)

        # ClipboardFormatOrAnsiString
        marker = self.strm.readUnsignedInt(4)
        if marker == 0:
            # Don't do anything.
            pass
        elif marker == 0xFFFFFFFF or marker == 0xFFFFFFFE:
            clipFormatID = self.strm.readUnsignedInt(4)
            print ("clipboard format ID: %d"%clipFormatID)
        else:
            clipName = globals.getUTF8FromUTF16(self.strm.readBytes(marker*2))
            if ord(clipName[-1]) != 0x00:
                # must be null-terminated.
                raise CompObjStreamError()
            print ("clipboard format name: %s"%clipName[:-1])

class PropertySetStream(object):

    def __init__ (self, bytes):
        self.strm = globals.ByteStream(bytes)

    def read (self):
        byteorder = self.strm.readUnsignedInt(2)
        if byteorder != 0xFFFE:
            raise PropertySetStreamError()
        ver = self.strm.readUnsignedInt(2)
        print ("version: 0x%4.4X"%ver)
        sID = self.strm.readUnsignedInt(4)
        print ("system identifier: 0x%4.4X"%sID)
        clsID = self.strm.readBytes(16)
        print ("CLS ID: %s"%globals.getRawBytes(clsID, True, False))
        sets = self.strm.readUnsignedInt(4)
        print ("number of property sets: 0x%4.4X"%sets)
        fmtID0 = self.strm.readBytes(16)
        print ("FMT ID 0: %s"%globals.getRawBytes(fmtID0, True, False))
        offset0 = self.strm.readUnsignedInt(4)
        print ("offset 0: 0x%4.4X"%offset0)
        if sets > 1:
            fmtID1 = self.strm.readBytes(16)
            print ("FMT ID 1: %s"%globals.getRawBytes(fmtID0, True, False))
            offset1 = self.strm.readUnsignedInt(4)
            print ("offset 1: 0x%4.4X\n"%offset1)
        self.readSet(offset0)
        if sets > 1:
            self.strm.setCurrentPos(offset1);
            self.readSet(offset1)

    def readSet (self, setOffset):
        print ("-----------------------------")
        print ("Property set")
        print ("-----------------------------")
        size = self.strm.readUnsignedInt(4)
        print ("size: 0x%4.4X"%size)
        props = self.strm.readUnsignedInt(4)
        print ("number of properties: 0x%4.4X"%props)
        pos = 0
        while pos < props:
            self.strm.setCurrentPos(setOffset + 8 + pos*8);
            id = self.strm.readUnsignedInt(4)
            offset = self.strm.readUnsignedInt(4)
            print ("ID: 0x%4.4X offset: 0x%4.4X"%(id, offset))
            self.strm.setCurrentPos(setOffset + offset);
            type = self.strm.readUnsignedInt(2)
            padding = self.strm.readUnsignedInt(2)
            if padding != 0:
                raise PropertySetStreamError()
            print ("type: 0x%4.4X"%type)
            if type == 2:
                value = self.strm.readSignedInt(2)
                print ("VT_I2: %d"%value)
            elif type == 0x41:
                blobSize = self.strm.readUnsignedInt(4)
                print ("VT_BLOB size: 0x%4.4X"%blobSize)
                print ("------------------------------------------------------------------------")
                globals.dumpBytes(self.strm.bytes[self.strm.pos:self.strm.pos+blobSize], blobSize)
                print ("------------------------------------------------------------------------")
            else:
                print ("unknown type")
            pos += 1
        print ("")
