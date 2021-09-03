#!/usr/bin/env python3
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
#

from .binarystream import BinaryStream
from . import globals
import time


PIDDSI = {
    0x00000001: "PIDDSI_CODEPAGE",
    0x00000002: "PIDDSI_CATEGORY",
    0x00000003: "PIDDSI_PRESFORMAT",
    0x00000004: "PIDDSI_BYTECOUNT",
    0x00000005: "PIDDSI_LINECOUNT",
    0x00000006: "PIDDSI_PARACOUNT",
    0x00000007: "PIDDSI_SLIDECOUNT",
    0x00000008: "PIDDSI_NOTECOUNT",
    0x00000009: "PIDDSI_HIDDENCOUNT",
    0x0000000A: "PIDDSI_MMCLIPCOUNT",
    0x0000000B: "PIDDSI_SCALE",
    0x0000000C: "PIDDSI_HEADINGPAIR",
    0x0000000D: "PIDDSI_DOCPARTS",
    0x0000000E: "PIDDSI_MANAGER",
    0x0000000F: "PIDDSI_COMPANY",
    0x00000010: "PIDDSI_LINKSDIRTY",
    0x00000011: "PIDDSI_CCHWITHSPACES",
    0x00000013: "PIDDSI_SHAREDDOC",
    0x00000014: "PIDDSI_LINKBASE",
    0x00000015: "PIDDSI_HLINKS",
    0x00000016: "PIDDSI_HYPERLINKSCHANGED",
    0x00000017: "PIDDSI_VERSION",
    0x00000018: "PIDDSI_DIGSIG",
    0x0000001A: "PIDDSI_CONTENTTYPE",
    0x0000001B: "PIDDSI_CONTENTSTATUS",
    0x0000001C: "PIDDSI_LANGUAGE",
    0x0000001D: "PIDDSI_DOCVERSION",
}


class DocumentSummaryInformationStream(BinaryStream):
    def __init__(self, bytes, params, doc):
        BinaryStream.__init__(self, bytes, params, "\x05DocumentSummaryInformation", doc=doc)

    def dump(self):
        print('<stream name="\\x05DocumentSummaryInformation" size="%d">' % self.size)
        PropertySetStream(self, PIDDSI).dump()
        print('</stream>')


PIDSI = {
    0x00000001: "CODEPAGE_PROPERTY_IDENTIFIER",
    0x00000002: "PIDSI_TITLE",
    0x00000003: "PIDSI_SUBJECT",
    0x00000004: "PIDSI_AUTHOR",
    0x00000005: "PIDSI_KEYWORDS",
    0x00000006: "PIDSI_COMMENTS",
    0x00000007: "PIDSI_TEMPLATE",
    0x00000008: "PIDSI_LASTAUTHOR",
    0x00000009: "PIDSI_REVNUMBER",
    0x0000000A: "PIDSI_EDITTIME",
    0x0000000B: "PIDSI_LASTPRINTED",
    0x0000000C: "PIDSI_CREATE_DTM",
    0x0000000D: "PIDSI_LASTSAVE_DTM",
    0x0000000E: "PIDSI_PAGECOUNT",
    0x0000000F: "PIDSI_WORDCOUNT",
    0x00000010: "PIDSI_CHARCOUNT",
    0x00000011: "PIDSI_THUMBNAIL",
    0x00000012: "PIDSI_APPNAME",
    0x00000013: "PIDSI_DOC_SECURITY",
}


class SummaryInformationStream(BinaryStream):
    def __init__(self, bytes, params, doc):
        BinaryStream.__init__(self, bytes, params, "\x05SummaryInformation", doc=doc)

    def dump(self):
        print('<stream name="\\x05SummaryInformation" size="%d">' % self.size)
        PropertySetStream(self, PIDSI).dump()
        print('</stream>')


class PropertySetStream(BinaryStream):
    """Specified by [MS-OLEPS] 2.21, the stream format for simple property sets."""
    def __init__(self, parent, PropertyIds):
        BinaryStream.__init__(self, parent.bytes)
        self.parent = parent
        self.propertyIds = PropertyIds

    def dump(self):
        print('<propertySetStream type="PropertySetStream" offset="%s">' % self.pos)
        self.printAndSet("ByteOrder", self.readuInt16())
        self.printAndSet("Version", self.readuInt16())
        self.printAndSet("SystemIdentifier", self.readuInt32())
        self.printAndSet("CLSID0", self.readuInt32())
        self.printAndSet("CLSID1", self.readuInt32())
        self.printAndSet("CLSID2", self.readuInt32())
        self.printAndSet("CLSID3", self.readuInt32())
        self.printAndSet("NumPropertySets", self.readuInt32())
        GUID(self, "FMTID0").dump()
        self.printAndSet("Offset0", self.readuInt32())
        PropertySet(self, self.Offset0).dump()
        if self.NumPropertySets == 0x00000002:
            GUID(self, "FMTID1").dump()
            self.printAndSet("Offset1", self.readuInt32())
            self.propertyIds = {}
            # The spec says: if NumPropertySets has the value 0x00000002,
            # FMTID1 must be set to FMTID_UserDefinedProperties.
            PropertySet(self, self.Offset1, userDefined=True).dump()
        print('</propertySetStream>')


class PropertySet(BinaryStream):
    def __init__(self, parent, offset, userDefined=False):
        BinaryStream.__init__(self, parent.bytes)
        self.parent = parent
        self.pos = offset
        self.userDefined = userDefined

    def getCodePage(self):
        for index, idAndOffset in enumerate(self.idsAndOffsets):
            if idAndOffset.PropertyIdentifier == 0x00000001:  # CODEPAGE_PROPERTY_IDENTIFIER
                return self.typedPropertyValues[index].Value

    def dump(self):
        if self.pos > self.size:
            return

        self.posOrig = self.pos
        print('<propertySet type="PropertySet" offset="%s">' % self.pos)
        self.printAndSet("Size", self.readuInt32())
        self.printAndSet("NumProperties", self.readuInt32())
        self.idsAndOffsets = []
        for i in range(self.NumProperties):
            idAndOffset = PropertyIdentifierAndOffset(self, i)
            idAndOffset.dump()
            self.idsAndOffsets.append(idAndOffset)
        self.typedPropertyValues = []
        for i in range(self.NumProperties):
            if self.userDefined and self.idsAndOffsets[i].PropertyIdentifier == 0x00000000:
                # [MS-OLEPS] 2.18.1 says the Dictionary property (id=0 in user-defined sets) has a different type.
                dictionary = Dictionary(self, i)
                dictionary.dump()
                self.typedPropertyValues.append("Dictionary")
            else:
                typedPropertyValue = TypedPropertyValue(self, i)
                typedPropertyValue.dump()
                self.typedPropertyValues.append(typedPropertyValue)
        print('</propertySet>')


class PropertyIdentifierAndOffset(BinaryStream):
    def __init__(self, parent, index):
        BinaryStream.__init__(self, parent.bytes)
        self.parent = parent
        self.index = index
        self.pos = parent.pos

    def dump(self):
        print('<propertyIdentifierAndOffset%s type="PropertyIdentifierAndOffset" offset="%s">' % (self.index, self.pos))
        self.printAndSet("PropertyIdentifier", self.readuInt32(), dict=self.parent.parent.propertyIds, default="unknown")
        self.printAndSet("Offset", self.readuInt32())
        print('</propertyIdentifierAndOffset%s>' % self.index)
        self.parent.pos = self.pos


PropertyType = {
    0x0000: "VT_EMPTY",
    0x0001: "VT_NULL",
    0x0002: "VT_I2",
    0x0003: "VT_I4",
    0x0004: "VT_R4",
    0x0005: "VT_R8",
    0x0006: "VT_CY",
    0x0007: "VT_DATE",
    0x0008: "VT_BSTR",
    0x000A: "VT_ERROR",
    0x000B: "VT_BOOL",
    0x000E: "VT_DECIMAL",
    0x0010: "VT_I1",
    0x0011: "VT_UI1",
    0x0012: "VT_UI2",
    0x0013: "VT_UI4",
    0x0014: "VT_I8",
    0x0015: "VT_UI8",
    0x0016: "VT_INT",
    0x0017: "VT_UINT",
    0x001E: "VT_LPSTR",
    0x001F: "VT_LPWSTR",
    0x0040: "VT_FILETIME",
    0x0041: "VT_BLOB",
    0x0042: "VT_STREAM",
    0x0043: "VT_STORAGE",
    0x0044: "VT_STREAMED_Object",
    0x0045: "VT_STORED_Object",
    0x0046: "VT_BLOB_Object",
    0x0047: "VT_CF",
    0x0048: "VT_CLSID",
    0x0049: "VT_VERSIONED_STREAM",
    0x1002: "VT_VECTOR | VT_I2",
    0x1003: "VT_VECTOR | VT_I4",
    0x1004: "VT_VECTOR | VT_R4",
    0x1005: "VT_VECTOR | VT_R8",
    0x1006: "VT_VECTOR | VT_CY",
    0x1007: "VT_VECTOR | VT_DATE",
    0x1008: "VT_VECTOR | VT_BSTR",
    0x100A: "VT_VECTOR | VT_ERROR",
    0x100B: "VT_VECTOR | VT_BOOL",
    0x100C: "VT_VECTOR | VT_VARIANT",
    0x1010: "VT_VECTOR | VT_I1",
    0x1011: "VT_VECTOR | VT_UI1",
    0x1012: "VT_VECTOR | VT_UI2",
    0x1013: "VT_VECTOR | VT_UI4",
    0x1014: "VT_VECTOR | VT_I8",
    0x1015: "VT_VECTOR | VT_UI8",
    0x101E: "VT_VECTOR | VT_LPSTR",
    0x101F: "VT_VECTOR | VT_LPWSTR",
    0x1040: "VT_VECTOR | VT_FILETIME",
    0x1047: "VT_VECTOR | VT_CF",
    0x1048: "VT_VECTOR | VT_CLSID",
    0x2002: "VT_ARRAY | VT_I2",
    0x2003: "VT_ARRAY | VT_I4",
    0x2004: "VT_ARRAY | VT_R4",
    0x2005: "VT_ARRAY | VT_R8",
    0x2006: "VT_ARRAY | VT_CY",
    0x2007: "VT_ARRAY | VT_DATE",
    0x2008: "VT_ARRAY | VT_BSTR",
    0x200A: "VT_ARRAY | VT_ERROR",
    0x200B: "VT_ARRAY | VT_BOOL",
    0x200C: "VT_ARRAY | VT_VARIANT",
    0x200E: "VT_ARRAY | VT_DECIMAL",
    0x2010: "VT_ARRAY | VT_I1",
    0x2011: "VT_ARRAY | VT_UI1",
    0x2012: "VT_ARRAY | VT_UI2",
    0x2013: "VT_ARRAY | VT_UI4",
    0x2016: "VT_ARRAY | VT_INT",
    0x2017: "VT_ARRAY | VT_UINT",

}


class DictionaryEntry(BinaryStream):
    """"Specified by [MS-OLEPS] 2.16, represents a mapping between a property
    identifier and a property name."""
    def __init__(self, parent, index):
        BinaryStream.__init__(self, parent.bytes)
        self.parent = parent
        self.pos = parent.pos
        self.index = index

    def dump(self):
        print('<dictionaryEntry offset="%s" index="%s">' % (self.pos, self.index))
        self.printAndSet("PropertyIdentifier", self.readuInt32())
        self.printAndSet("Length", self.readuInt32())

        bytes = []
        for dummy in range(self.Length):
            c = self.readuInt8()
            if c == 0:
                break
            bytes.append(c)
        # TODO support non-latin1
        encoding = "latin1"
        if globals.PY3:
            print('<Name value="%s"/>' % globals.encodeName(b"".join(map(lambda c: globals.indexedbytetobyte(c), bytes)).decode(encoding), lowOnly=True).encode('utf-8'))
        else:
            print('<Name value="%s"/>' % globals.encodeName("".join(map(lambda c: chr(c), bytes)).decode(encoding), lowOnly=True).encode('utf-8'))

        print('</dictionaryEntry>')
        self.parent.pos = self.pos


class Dictionary(BinaryStream):
    """Specified by [MS-OLEPS] 2.17, represents all mappings between property
    identifiers and property names in a property set."""
    def __init__(self, parent, index):
        BinaryStream.__init__(self, parent.bytes)
        self.parent = parent
        self.index = index
        self.pos = parent.posOrig + parent.idsAndOffsets[index].Offset

    def dump(self):
        print('<dictionary%s type="Dictionary" offset="%s">' % (self.index, self.pos))
        self.printAndSet("NumEntries", self.readuInt32())
        for i in range(self.NumEntries):
            dictionaryEntry = DictionaryEntry(self, i)
            dictionaryEntry.dump()
        print('</dictionary%s>' % self.index)


class TypedPropertyValue(BinaryStream):
    def __init__(self, parent, index):
        BinaryStream.__init__(self, parent.bytes)
        self.parent = parent
        self.index = index
        self.pos = parent.posOrig + parent.idsAndOffsets[index].Offset

    def dump(self):
        print('<typedPropertyValue%s type="TypedPropertyValue" offset="%s">' % (self.index, self.pos))
        self.printAndSet("Type", self.readuInt16(), dict=PropertyType)
        self.printAndSet("Padding", self.readuInt16())
        if self.Type == 0x0002:  # VT_I2
            self.printAndSet("Value", self.readInt16())
        elif self.Type == 0x0003:  # VT_I4
            self.printAndSet("Value", self.readInt32())
        elif self.Type == 0x000b:  # VARIANT_BOOL
            self.printAndSet("Value", self.readuInt32())
        elif self.Type == 0x0040:  # VT_FILETIME
            FILETIME(self, "Value").dump()
        elif self.Type == 0x001E:  # VT_LPSTR
            CodePageString(self, "Value", self.parent.getCodePage()).dump()
        elif self.Type == 0x101E:  # VT_VECTOR | VT_LPSTR
            VectorHeader(self, "Value", self.parent.getCodePage()).dump()
        else:
            print('<todo what="TypedPropertyValue::dump: unhandled Type %s"/>' % hex(self.Type))
        print('</typedPropertyValue%s>' % self.index)


class VectorHeader(BinaryStream):
    """Defined by [MS-OLEPS] 2.14.2, represents the number of scalar values in
    a vector property type."""
    def __init__(self, parent, name, codepage):
        BinaryStream.__init__(self, parent.bytes)
        self.pos = parent.pos
        self.parent = parent
        self.name = name
        self.codepage = codepage

    def dump(self):
        print('<%s type="VectorHeader">' % self.name)
        self.printAndSet("Length", self.readuInt32())
        for dummy in range(self.Length):
            CodePageString(self, "String", self.codepage).dump()
        print('</%s>' % self.name)


class CodePageString(BinaryStream):
    def __init__(self, parent, name, codepage):
        BinaryStream.__init__(self, parent.bytes)
        self.pos = parent.pos
        self.parent = parent
        self.name = name
        self.codepage = codepage

    def dump(self):
        print('<%s type="CodePageString">' % self.name)
        self.printAndSet("Size", self.readuInt32())
        bytes = []
        for dummy in range(self.Size):
            c = self.readuInt8()
            if c == 0:
                break
            bytes.append(c)
        codepage = self.codepage
        if (codepage is not None) and (codepage < 0):
            codepage += 2 ** 16  # signed -> unsigned
        encoding = ""
        # http://msdn.microsoft.com/en-us/goglobal/bb964654
        if codepage == 1252:
            encoding = "latin1"
        elif codepage == 1250:
            encoding = "latin2"
        elif codepage == 65001:
            # http://msdn.microsoft.com/en-us/library/windows/desktop/dd374130%28v=vs.85%29.aspx
            encoding = "utf-8"
        if len(encoding):
            if globals.PY3:
                print('<Characters value="%s"/>' % globals.encodeName(b"".join(map(lambda c: globals.indexedbytetobyte(c), bytes)).decode(encoding), lowOnly=True).encode('utf-8'))
            else:
                # Argh cant use indexedbytetobyte because we actually have ints
                print('<Characters value="%s"/>' % globals.encodeName("".join(map(lambda c: chr(c), bytes)).decode(encoding), lowOnly=True).encode('utf-8'))
        else:
            print('<todo what="CodePageString::dump: unhandled codepage %s"/>' % codepage)
        print('</%s>' % self.name)


class GUID(BinaryStream):
    def __init__(self, parent, name):
        BinaryStream.__init__(self, parent.bytes)
        self.pos = parent.pos
        self.parent = parent
        self.name = name

    def dump(self):
        Data1 = self.readuInt32()
        Data2 = self.readuInt16()
        Data3 = self.readuInt16()
        Data4 = []
        for dummy in range(8):
            Data4.append(self.readuInt8())
        value = "%08x-%04x-%04x-%02x%02x-%02x%02x%02x%02x%02x%02x" % (Data1, Data2, Data3, Data4[0], Data4[1], Data4[2], Data4[3], Data4[4], Data4[5], Data4[6], Data4[7])
        print('<%s type="GUID" value="%s"/>' % (self.name, value))
        self.parent.pos = self.pos


class OLERecord(BinaryStream):
    def __init__(self, parent):
        BinaryStream.__init__(self, parent.bytes)
        self.parent = parent
        self.pos = parent.pos


class FILETIME(OLERecord):
    def __init__(self, parent, name):
        OLERecord.__init__(self, parent)
        self.name = name

    def dump(self):
        # ft is number of 100ns since Jan 1 1601
        ft = self.readuInt64()
        if ft > 0:
            epoch = 11644473600
            sec = (ft / 10000000) - epoch
        else:
            sec = ft
        try:
            pretty = time.strftime("%Y-%m-%dT%H:%M:%SZ", time.localtime(sec))
        except ValueError:
            pretty = "ValueError"
        print('<%s type="FILETIME" value="%d" pretty="%s"/>' % (self.name, sec, pretty))
        self.parent.pos = self.pos


# vim:set filetype=python shiftwidth=4 softtabstop=4 expandtab:
