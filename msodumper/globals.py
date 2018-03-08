# -*- tab-width: 4; indent-tabs-mode: nil -*-
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
from builtins import range
import sys, struct, math, zipfile, io
from . import xmlpp

PY3 = sys.version > '3'

# Python3 bytes[i] is an integer, Python2  str[i] is an str of length 1
# These functions explicitely return a given type for both 2 and 3
# - indexbytes(): return bytes element at given index, as integer (uses
#   ord() for python2)
# - indexedbytetobyte: convert a value obtained by indexing into a bytes list
#   (or from 'for x in somebytes') to a bytes list of len 1
# - indexedbytetoint() return the same as an int (uses ord() for python2)
#   
if PY3:

    def indexbytes(data, i):
        return data[i]
    def indexedbytetobyte(i):
        return i.to_bytes(1, byteorder='big')        
    def indexedbytetoint(i):
        return i
    nullchar = 0
else:
    def indexbytes(data, i):
        return ord(data[i])
    def indexedbytetobyte(i):
        return i
    def indexedbytetoint(i):
        return ord(i)
    nullchar = chr(0)

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
        self.noStructOutput = False
        self.dumpText = False
        self.dumpedIds = []
        self.noRawDump = False
        self.catchExceptions = False
        self.utf8 = False
        
# Global parameters / run configuration, to be set up by the main
# program during initialization
params = Params()

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
        if idx in list:
            return list[idx]

    return errmsg

textdump = b""

def dumptext():
    data = textdump.replace(b"\r", b"\n")
    if sys.platform == "win32":
        import msvcrt
        msvcrt.setmode(sys.stdout.fileno(), os.O_BINARY)
    if PY3:
        sys.stdout.buffer.write(data)
    else:
        sys.stdout.write(data)

# Write msg to stdout, as bytes (encode it if needed)
def output (msg, recordType = -1):
    if type(msg) == type(u''):
        msg = msg.encode('utf-8')
    if params.noStructOutput:
        return
    if recordType == -1 or not params.dumpedIds or \
           recordType in params.dumpedIds:
        if PY3:
            sys.stdout.buffer.write(msg)
        else:
            sys.stdout.write(msg)

def outputln(msg, recordType = -1):
    if type(msg) == type(u''):
        output(msg + "\n", recordType)
    else:
        output(msg + b"\n", recordType)

# Replace sys.stdout as arg to prettyPrint and call our output method
class utfwriter:
    def write(self, s):
        output(s)

def error (msg):
    sys.stderr.write("Error: %s\n"%msg)

def debug (msg):
    sys.stderr.write("DEBUG: %s\n"%msg)

def nulltrunc(bytes):
    '''Return input truncated to first 0 byte, allowing, e.g., comparison
    with a simple literal bytes string'''
    try:
        firstnull = bytes.index(nullchar)
        bytes = bytes[:firstnull]
    except:
        pass
    return bytes


# This is syntaxically identical for python2 and python3 if the input is str 
def encodeName (name, lowOnly = False, lowLimit = 0x20):
    """Encode name that contains unprintable characters."""

    n = len(name)
    if n == 0:
        return name
    
    if PY3 and (type(name) == type(b'')):
        return _encodeNameBytes(name, n, lowOnly, lowLimit)

    newname = ''
    for i in range(0, n):
        if name[i] == '&'[0]:
            newname += "&amp;"
        elif name[i] == '<'[0]:
            newname += "&lt;"
        elif name[i] == '>'[0]:
            newname += "&gt;"
        elif ord(name[i]) < lowLimit or ((not lowOnly) and ord(name[i]) >= 127):
            newname += "\\x%2.2X"%ord(name[i])
        else:
            newname += name[i]
    
    return newname

# Python3 only. Same as above but accept bytes as input.
def _encodeNameBytes (name, n, lowOnly = False, lowLimit = 0x20):
    """Encode name that contains unprintable characters."""

    newname = ''
    for i in range(0, n):
        if name[i] == b'&'[0]:
            newname += "&amp;"
        elif name[i] == '<'[0]:
            newname += b"&lt;"
        elif name[i] == '>'[0]:
            newname += b"&gt;"
        elif name[i] < lowLimit or ((not lowOnly) and name[i] >= 127):
            newname += "\\x%2.2X"%name[i]
        else:
            newname += chr(name[i])

    return newname
    
# Uncompress "compressed" UTF-16. This compression strips high bytes
# from a string when they are all 0. Just restore them.
def uncompCompUnicode(bytes):
    out = b""
    for b in bytes:
        out += indexedbytetobyte(b)
        out += b'\0'
    return out

class UnicodeRichExtText(object):
    def __init__ (self):
        self.baseText = u''
        self.phoneticBytes = []

# Search sorted list for first element strictly bigger than input
# value. Should be binary search, but CONTINUE record offsets list are
# usually small. Return list size if value >= last list element
def findFirstBigger(ilist, value):
    i = 0
    while i < len(ilist) and value >= ilist[i]:
        i +=1
    return i

def getUnicodeRichExtText (bytes, offset = 0, rofflist = []):
    if len(rofflist) == 0:
        rofflist = [len(bytes)]
    ret = UnicodeRichExtText()
    # Avoid myriad of messages when in "catching" mode
    if params.catchExceptions and (bytes is None or len(bytes) == 0):
        return ret, 0

    if len(rofflist) == 0 or rofflist[len(rofflist)-1] != len(bytes):
        error("bad input to getUnicodeRichExtText: empty offset list or last offset != size. size %d list %s" % (len(bytes), str(rofflist)))
        raise ByteStreamError()

    strm = ByteStream(bytes)
    strm.setCurrentPos(offset)

    try:
        textLen = strm.readUnsignedInt(2)
        flags = strm.readUnsignedInt(1)
        #  0 0 0 0 0 0 0 0
        # |-------|D|C|B|A|
        if (flags & 0x01) > 0: # A
            bytesPerChar = 2
        else:
            bytesPerChar = 1
        ignored      = (flags & 0x02) > 0 # B
        hasPhonetic  = (flags & 0x04) > 0 # C
        isRichStr    = (flags & 0x08) > 0 # D

        numElem = 0
        if isRichStr:
            numElem = strm.readUnsignedInt(2)

        phoneticBytes = 0
        if hasPhonetic:
            phoneticBytes = strm.readUnsignedInt(4)

        # Reading the string proper. This is made a bit more
        # complicated by the fact that the format can switch from
        # compressed (latin data with high zeros stripped) to normal
        # (UTF-16LE) whenever a string encounters a CONTINUE record
        # boundary. The new format is indicated by a single byte at
        # the start of the CONTINUE record payload.
        while textLen > 0:
            #print("Reading Unicode with bytesPerChar %d" % bytesPerChar)
            bytesToRead = textLen * bytesPerChar

            # Truncate to next record boundary
            ibound = findFirstBigger(rofflist, strm.getCurrentPos())
            if ibound == len(rofflist):
                # Just try to read and let the stream raise an exception
                strm.readBytes(bytesToRead)
                return
            
            bytesToRead = min(bytesToRead, \
                              rofflist[ibound]- strm.getCurrentPos())
            newdata = strm.readBytes(bytesToRead)
            if bytesPerChar == 1:
                newdata = uncompCompUnicode(newdata)

            ret.baseText +=  newdata.decode('UTF-16LE', errors='replace')

            textLen -= bytesToRead // bytesPerChar
            
            # If there is still data to read, we hit a record boundary. Read
            # the grbit byte for detecting possible compression switch
            if textLen > 0:
                grbit = strm.readUnsignedInt(1)
                if (grbit & 1) != 0:
                    bytesPerChar = 2
                else:
                    bytesPerChar = 1
                
        if isRichStr:
            for i in range(0, numElem):
                posChar = strm.readUnsignedInt(2)
                fontIdx = strm.readUnsignedInt(2)

        if hasPhonetic:
            ret.phoneticBytes = strm.readBytes(phoneticBytes)
    except Exception as e:
        if not params.catchExceptions:
            raise
        error("getUnicodeRichExtText: %s\n" % e)
        return ret, len(bytes)
    return ret, strm.getCurrentPos() - offset


def getRichText (bytes, textLen=None):
    """parse a string of the rich-text format that Excel uses.
Return python2/3 unicode()/str()

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
        text = strm.readBytes(2*textLen).decode('UTF-16LE', errors='replace')
    else:
        if params.utf8:
            # Compressed Unicode-> latin1
            text = strm.readBytes(textLen).decode('cp1252')
        else:
            # Old behaviour with hex dump
            text = strm.readBytes(textLen).decode('cp1252')

    return (text, totalByteLen)

if PY3:
    def toCharOrDot(char):
        if 32 < char and char < 127:
            return chr(char)
        else:
            return '.'
else:
    def toCharOrDot(char):
        if 32 < ord(char) and ord(char) < 127:
            return char
        else:
            return '.'

def dumpBytes (chars, subDivide=None):
    if params.noStructOutput or params.noRawDump:
        return
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
    for i in range(0, charLen):
        if (i+1)%16 == 1:
            # print line header with seek position
            fmt = "%%%d.%dd: "%(labelWidth, labelWidth)
            output(fmt%i)

        byte = indexbytes(chars, i)
        lineBuf += toCharOrDot(chars[i])
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


# TBD: getTextBytes is now only called from pptrecord.
# getTextBytes() and toTextBytes() are probably not
# needed any more now that we store text as str not list.
# toTextBytes() has been changed to do nothing until we're sure we can dump it
def getTextBytes (bytes):
    return toTextBytes(bytes)

def toTextBytes (bytes):
    return bytes

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
    elif n == 8:
        # long long - 8 bytes
        return struct.unpack('<Q', text)[0]

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
    loopCount = len(bytes) // 2

    # Truncate input to first null doublet
    for i in range(0, loopCount):
        if indexbytes(bytes, i*2) == 0:
            if indexbytes(bytes, i*2+1) == 0:
                bytes = bytes[0:i*2]
                break

    # Convert from utf-16 and return utf-8, using markers for
    # conversion errors
    if type(bytes) != type(u''):
        bytes = bytes.decode('UTF-16LE', errors='replace')
    return bytes.encode('UTF-8')

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
    rawFile = io.BytesIO(bytes)
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
        if filename.endswith(".xml") or contents.startswith(b"<?xml"):
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

# vim:set filetype=python shiftwidth=4 softtabstop=4 expandtab:
