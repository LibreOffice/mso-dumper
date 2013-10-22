########################################################################
#
#  Copyright (c) 2010 Kohei Yoshida
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

import sys, struct, math, zipfile, xmlpp, StringIO

OutputWidth = 76

class ByteConvertError(Exception): pass

class ByteStreamError(Exception): pass


class ModelBase(object):

    class HostAppType:
        Word       = 0
        Excel      = 1
        PowerPoint = 2

    def __init__ (self, hostApp):
        self.hostApp = hostApp


class Params(object):
    """command-line parameters."""
    def __init__ (self):
        self.debug = False
        self.showSectorChain = False
        self.showStreamPos = False


class ByteStream(object):

    def __init__ (self, bytes):
        self.bytes = bytes
        self.pos = 0
        self.size = len(bytes)

    def getSize (self):
        return self.size

    def readBytes (self, length):
        if self.pos + length > self.size:
            error("reading %d bytes from position %d would exceed the current size of %d\n"%
                (length, self.pos, self.size))
            raise ByteStreamError()
        r = self.bytes[self.pos:self.pos+length]
        self.pos += length
        return r

    def readRemainingBytes (self):
        r = self.bytes[self.pos:]
        self.pos = self.size
        return r

    def getCurrentPos (self):
        return self.pos

    def setCurrentPos (self, pos):
        self.pos = pos

    def isEndOfRecord (self):
        return (self.pos == self.size)

    def readUnsignedInt (self, length):
        bytes = self.readBytes(length)
        return getUnsignedInt(bytes)

    def readSignedInt (self, length):
        bytes = self.readBytes(length)
        return getSignedInt(bytes)

    def readDouble (self):
        # double is always 8 bytes.
        bytes = self.readBytes(8)
        return getDouble(bytes)

    def readUnicodeString (self, textLen=None):
        # First 2-bytes contains the text length, followed by a 1-byte flag.
        if textLen == None:
            textLen = self.readUnsignedInt(2)
        bytes = self.bytes[self.pos:]
        text, byteLen = getRichText(bytes, textLen)
        self.moveForward (byteLen)
        return text

    def readFixedPoint (self):
        # [MS-OSHARED] Section 2.2.1.6 FixedPoint
        integral = self.readSignedInt(2)
        fractional = self.readUnsignedInt(2)
        val = 0.0
        val += integral + (fractional / 65536.0)
        return val

    def moveBack (self, byteCount):
        self.pos -= byteCount
        if self.pos < 0:
            self.pos = 0

    def moveForward (self, byteCount):
        self.pos += byteCount
        if self.pos > self.size:
            self.pos = self.size


def getValueOrUnknown (list, idx, errmsg='(unknown)'):
    listType = type(list)
    if listType == type([]):
        # list
        if idx < len(list):
            return list[idx]
    elif listType == type({}):
        # dictionary
        if list.has_key(idx):
            return list[idx]

    return errmsg


def output (msg):
    sys.stdout.write(msg)

def error (msg):
    sys.stderr.write("Error: " + msg)

def debug (msg):
    sys.stderr.write("DEBUG: %s\n"%msg)


def encodeName (name, lowOnly = False, lowLimit = 0x20):
    """Encode name that contains unprintable characters."""

    n = len(name)
    if n == 0:
        return name

    newname = ''
    for i in xrange(0, n):
        if ord(name[i]) < lowLimit or ((not lowOnly) and ord(name[i]) >= 127):
            newname += "\\x%2.2X"%ord(name[i])
        else:
            newname += name[i]

    return newname


class UnicodeRichExtText(object):
    def __init__ (self):
        self.baseText = ''
        self.phoneticBytes = []


def getUnicodeRichExtText (bytes):
    ret = UnicodeRichExtText()
    strm = ByteStream(bytes)
    textLen = strm.readUnsignedInt(2)
    flags = strm.readUnsignedInt(1)
    #  0 0 0 0 0 0 0 0
    # |-------|D|C|B|A|
    isDoubleByte = (flags & 0x01) > 0 # A
    ignored      = (flags & 0x02) > 0 # B
    hasPhonetic  = (flags & 0x04) > 0 # C
    isRichStr    = (flags & 0x08) > 0 # D

    numElem = 0
    if isRichStr:
        numElem = strm.readUnsignedInt(2)

    phoneticBytes = 0
    if hasPhonetic:
        phoneticBytes = strm.readUnsignedInt(4)
        
    if isDoubleByte:
        # double-byte string (UTF-16)
        text = ''
        for i in xrange(0, textLen):
            text += toTextBytes(strm.readBytes(2)).decode('utf-16')
        ret.baseText = text
    else:
        # single-byte string
        ret.baseText = toTextBytes(strm.readBytes(textLen))

    if isRichStr:
        for i in xrange(0, numElem):
            posChar = strm.readUnsignedInt(2)
            fontIdx = strm.readUnsignedInt(2)

    if hasPhonetic:
        ret.phoneticBytes = strm.readBytes(phoneticBytes)

    return ret, strm.getCurrentPos()


def getRichText (bytes, textLen=None):
    """parse a string of the rich-text format that Excel uses.

Note the following:

  * The 1st byte always contains flag.
  * The actual number of bytes read may differ depending on the values of the 
    flags, so the client code should pass an open-ended stream of bytes and 
    always query for the actual bytes read to adjust for the new stream 
    position when this function returns.
"""

    strm = ByteStream(bytes)
    flags = strm.readUnsignedInt(1)
    if type(flags) == type('c'):
        flags = ord(flags)
    is16Bit   = (flags & 0x01)
    isFarEast = (flags & 0x04)
    isRich    = (flags & 0x08)

    formatRuns = 0
    if isRich:
        formatRuns = strm.readUnsignedInt(2)

    extInfo = 0
    if isFarEast:
        extInfo = strm.readUnsignedInt(4)

    extraBytes = 0
    if textLen == None:
        extraBytes = formatRuns*4 + extInfo
        textLen = len(bytes) - extraBytes

    totalByteLen = strm.getCurrentPos() + textLen + extraBytes
    if is16Bit:
        totalByteLen += textLen # double the text length since each char is 2 bytes.
        text = ''
        for i in xrange(0, textLen):
            text += toTextBytes(strm.readBytes(2)).decode('utf-16')
    else:
        text = toTextBytes(strm.readBytes(textLen))

    return (text, totalByteLen)


def dumpBytes (chars, subDivide=None):
    line = 0
    subDivideLine = None
    if subDivide != None:
        subDivideLine = subDivide/16

    charLen = len(chars)
    if charLen == 0:
        # no bytes to dump.
        return

    labelWidth = int(math.ceil(math.log(charLen, 10)))
    lineBuf = '' # bytes interpreted as chars at the end of each line
    i = 0
    for i in xrange(0, charLen):
        if (i+1)%16 == 1:
            # print line header with seek position
            fmt = "%%%d.%dd: "%(labelWidth, labelWidth)
            output(fmt%i)

        byte = ord(chars[i])
        if 32 < byte and byte < 127:
            lineBuf += chars[i]
        else:
            lineBuf += '.'
        output("%2.2X "%byte)

        if (i+1)%4 == 0:
            # put extra space at every 4 bytes.
            output(" ")

        if (i+1)%16 == 0:
            # end of line
            output (lineBuf)
            output("\n")
            if subDivideLine != None and (line+1)%subDivideLine == 0:
                output("\n")
            lineBuf = ''
            line += 1

    if len(lineBuf) > 0:
        # pad with white space so that the line string gets aligned.
        i += 1
        while True:
            output ("   ")
            if (i+1)%4 == 0:
                # put extra space at every 4 bytes.
                output(" ")
            if (i+1) % 16 == 0:
                # end of line
                break
            i += 1
        output (lineBuf)
        output("\n")

def getSectorPos (secID, secSize):
    return 512 + secID*secSize


def getRawBytes (bytes, spaced=True, reverse=True):
    text = ''
    for b in bytes:
        if type(b) == type(''):
            b = ord(b)
        if len(text) == 0:
            text = "%2.2X"%b
        elif spaced:
            if reverse:
                text = "%2.2X "%b + text
            else:
                text += " %2.2X"%b
        else:
            if reverse:
                text = "%2.2X"%b + text
            else:
                text += "%2.2X"%b
    return text


def getTextBytes (bytes):
    return toTextBytes(bytes)


def toTextBytes (bytes):
    n = len(bytes)
    text = ''
    for i in xrange(0, n):
        b = bytes[i]
        if type(b) == type(0x00):
            b = struct.pack('B', b)
        text += b
    return text


def getSignedInt (bytes):
    # little endian
    n = len(bytes)
    if n == 0:
        return 0

    text = toTextBytes(bytes)
    if n == 1:
        # byte - 1 byte
        return struct.unpack('b', text)[0]
    elif n == 2:
        # short - 2 bytes
        return struct.unpack('<h', text)[0]
    elif n == 4:
        # int, long - 4 bytes
        return struct.unpack('<l', text)[0]

    raise ByteConvertError


def getUnsignedInt (bytes):
    # little endian
    n = len(bytes)
    if n == 0:
        return 0

    text = toTextBytes(bytes)
    if n == 1:
        # byte - 1 byte
        return struct.unpack('B', text)[0]
    elif n == 2:
        # short - 2 bytes
        return struct.unpack('<H', text)[0]
    elif n == 4:
        # int, long - 4 bytes
        return struct.unpack('<L', text)[0]

    raise ByteConvertError


def getFloat (bytes):
    n = len(bytes)
    if n == 0:
        return 0.0

    text = toTextBytes(bytes)
    return struct.unpack('<f', text)[0]


def getDouble (bytes):
    n = len(bytes)
    if n == 0:
        return 0.0

    text = toTextBytes(bytes)
    return struct.unpack('<d', text)[0]


def getUTF8FromUTF16 (bytes):
    # little endian utf-16 strings
    byteCount = len(bytes)
    loopCount = int(byteCount/2)
    text = ''
    for i in xrange(0, loopCount):
        code = ''
        lsbZero = bytes[i*2] == '\x00'
        msbZero = bytes[i*2+1] == '\x00'
        if msbZero and lsbZero:
            return text
        
        if not msbZero:
            code += bytes[i*2+1]
        if not lsbZero:
            code += bytes[i*2]
        try:    
            text += unicode(code, 'utf-8')
        except UnicodeDecodeError:
            text += "<%d invalid chars>"%len(code)
    return text

class StreamWrap(object):
    def __init__ (self,printer):
        self.printer = printer
        self.buffer = ""
    def write (self,string):
        self.buffer += string
    def flush (self):
        for line in self.buffer.splitlines():
            self.printer(line)

def outputZipContent (bytes, printer, width=80):
    printer("Zipped content:")
    rawFile = StringIO.StringIO(bytes)
    zipFile = zipfile.ZipFile(rawFile)
    i = 0
    # TODO: when 2.6/3.0 is in widespread use, change to infolist
    # here, names might be ambiguous
    for filename in zipFile.namelist():
        if i > 0:
            printer('-'*width)
        i += 1
        printer("")
        printer(filename + ":")
        printer('-'*width)

        contents = zipFile.read(filename)
        if filename.endswith(".xml") or contents.startswith("<?xml"):
            wrapper = StreamWrap(printer)
            xmlpp.pprint(contents,wrapper,1,80)
            wrapper.flush()
        else:
            dumpBytes(contents)
            
    zipFile.close()

def stringizeColorRef(colorRef, colorName="color"):
    def split (packedColor):
        return ((packedColor & 0xFF0000) // 0x10000, (packedColor & 0xFF00) / 0x100, (packedColor & 0xFF))
    
    colorValue = colorRef & 0xFFFFFF
    if colorRef & 0xFE000000 == 0xFE000000 or colorRef & 0xFF000000 == 0:
        colors = split(colorValue)
        return "%s = (%d,%d,%d)"%(colorName, colors[0], colors[1], colors[2])
    elif colorRef & 0x08000000 or colorRef & 0x10000000:
        return "%s = schemecolor(%d)"%(colorName, colorValue)
    elif colorRef & 0x04000000:
        return "%s = colorschemecolor(%d)"%(colorName, colorValue)
    else:
        return "%s = <unidentified color>(%4.4Xh)"%(colorName, colorValue)
