########################################################################
#
#    OpenOffice.org - a multi-platform office productivity suite
#
#    Author:
#      Kohei Yoshida  <kyoshida@novell.com>
#
#   The Contents of this file are made available subject to the terms
#   of GNU Lesser General Public License Version 2.1 and any later
#   version.
#
########################################################################

import sys, struct, math, zipfile, xmlpp, StringIO

class ByteConvertError(Exception): pass


class Params(object):
    """command-line parameters."""
    def __init__ (self):
        self.debug = False
        self.showSectorChain = False
        self.showStreamPos = False


class StreamData(object):
    """run-time stream data."""
    def __init__ (self):
        self.encrypted = False
        self.pivotCacheIDs = {}

    def appendPivotCacheId (self, newId):
        # must be 4-digit with leading '0's.
        strId = "%.4d"%newId
        self.pivotCacheIDs[strId] = True

    def isPivotCacheStream (self, name):
        return self.pivotCacheIDs.has_key(name)


class ByteStream(object):

    def __init__ (self, bytes):
        self.bytes = bytes
        self.pos = 0
        self.size = len(bytes)

    def getSize (self):
        return self.size

    def readBytes (self, length):
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


def output (msg):
    sys.stdout.write(msg)

def error (msg):
    sys.stderr.write("Error: " + msg)


def decodeName (name):
    """decode name that contains unprintable characters."""

    n = len(name)
    if n == 0:
        return name

    newname = ''
    for i in xrange(0, n):
        if ord(name[i]) <= 20:
            newname += "<%2.2Xh>"%ord(name[i])
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
    isDoubleByte = (flags     & 0x01) == 1 # A
    ignored      = ((flags/2) & 0x01) == 1 # B
    hasPhonetic  = ((flags/4) & 0x01) == 1 # C
    isRichStr    = ((flags/8) & 0x01) == 1 # D

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
    """parse a string of the rich-text format that Excel uses."""

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
        return ("<16-bit strings not supported yet>", totalByteLen)

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
    flushBytes = False
    for i in xrange(0, charLen):
        if (i+1)%16 == 1:
            # print line header with seek position
            fmt = "%%%d.%dd: "%(labelWidth, labelWidth)
            output(fmt%i)

        byte = ord(chars[i])
        output("%2.2X "%byte)
        flushBytes = True

        if (i+1)%4 == 0:
            # put extra space at every 4 bytes.
            output(" ")

        if (i+1)%16 == 0:
            output("\n")
            flushBytes = False
            if subDivideLine != None and (line+1)%subDivideLine == 0:
                output("\n")
            line += 1

    if flushBytes:
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
