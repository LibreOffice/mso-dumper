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




