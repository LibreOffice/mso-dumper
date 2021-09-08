#!/usr/bin/env python3
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
#

from .binarystream import BinaryStream

PlaceableKey = {
    0x9ac6cdd7: "META_PLACEABLE",
}

FileType = {
    0x00: "Memory",
    0x01: "Disk",
}

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

# No idea what's the proper name of this thing, see
# http://msdn.microsoft.com/en-us/library/dd145130%28VS.85%29.aspx
RasterPolishMap = {
    0x00000042: "0",
    0x00010289: "DPSoon",
    0x00020C89: "DPSona",
    0x000300AA: "PSon",
    0x00040C88: "SDPona",
    0x000500A9: "DPon",
    0x00060865: "PDSxnon",
    0x000702C5: "PDSaon",
    0x00080F08: "SDPnaa",
    0x00090245: "PDSxon",
    0x000A0329: "DPna",
    0x000B0B2A: "PSDnaon",
    0x000C0324: "SPna",
    0x000D0B25: "PDSnaon",
    0x000E08A5: "PDSonon",
    0x000F0001: "Pn",
    0x00100C85: "PDSona",
    0x001100A6: "DSon",
    0x00120868: "SDPxnon",
    0x001302C8: "SDPaon",
    0x00140869: "DPSxnon",
    0x001502C9: "DPSaon",
    0x00165CCA: "PSDPSanaxx",
    0x00171D54: "SSPxDSxaxn",
    0x00180D59: "SPxPDxa",
    0x00191CC8: "SDPSanaxn",
    0x001A06C5: "PDSPaox",
    0x001B0768: "SDPSxaxn",
    0x001C06CA: "PSDPaox",
    0x001D0766: "DSPDxaxn",
    0x001E01A5: "PDSox",
    0x001F0385: "PDSoan",
    0x00200F09: "DPSnaa",
    0x00210248: "SDPxon",
    0x00220326: "DSna",
    0x00230B24: "SPDnaon",
    0x00240D55: "SPxDSxa",
    0x00251CC5: "PDSPanaxn",
    0x002606C8: "SDPSaox",
    0x00271868: "SDPSxnox",
    0x00280369: "DPSxa",
    0x002916CA: "PSDPSaoxxn",
    0x002A0CC9: "DPSana",
    0x002B1D58: "SSPxPDxaxn",
    0x002C0784: "SPDSoax",
    0x002D060A: "PSDnox",
    0x002E064A: "PSDPxox",
    0x002F0E2A: "PSDnoan",
    0x0030032A: "PSna",
    0x00310B28: "SDPnaon",
    0x00320688: "SDPSoox",
    0x00330008: "Sn",
    0x003406C4: "SPDSaox",
    0x00351864: "SPDSxnox",
    0x003601A8: "SDPox",
    0x00370388: "SDPoan",
    0x0038078A: "PSDPoax",
    0x00390604: "SPDnox",
    0x003A0644: "SPDSxox",
    0x003B0E24: "SPDnoan",
    0x003C004A: "PSx",
    0x003D18A4: "SPDSonox",
    0x003E1B24: "SPDSnaox",
    0x003F00EA: "PSan",
    0x00400F0A: "PSDnaa",
    0x00410249: "DPSxon",
    0x00420D5D: "SDxPDxa",
    0x00431CC4: "SPDSanaxn",
    0x00440328: "SDna",
    0x00450B29: "DPSnaon",
    0x004606C6: "DSPDaox",
    0x0047076A: "PSDPxaxn",
    0x00480368: "SDPxa",
    0x004916C5: "PDSPDaoxxn",
    0x004A0789: "DPSDoax",
    0x004B0605: "PDSnox",
    0x004C0CC8: "SDPana",
    0x004D1954: "SSPxDSxoxn",
    0x004E0645: "PDSPxox",
    0x004F0E25: "PDSnoan",
    0x00500325: "PDna",
    0x00510B26: "DSPnaon",
    0x005206C9: "DPSDaox",
    0x00530764: "SPDSxaxn",
    0x005408A9: "DPSonon",
    0x00550009: "Dn",
    0x005601A9: "DPSox",
    0x00570389: "DPSoan",
    0x00580785: "PDSPoax",
    0x00590609: "DPSnox",
    0x005A0049: "DPx",
    0x005B18A9: "DPSDonox",
    0x005C0649: "DPSDxox",
    0x005D0E29: "DPSnoan",
    0x005E1B29: "DPSDnaox",
    0x005F00E9: "DPan",
    0x00600365: "PDSxa",
    0x006116C6: "DSPDSaoxxn",
    0x00620786: "DSPDoax",
    0x00630608: "SDPnox",
    0x00640788: "SDPSoax",
    0x00650606: "DSPnox",
    0x00660046: "DSx",
    0x006718A8: "SDPSonox",
    0x006858A6: "DSPDSonoxxn",
    0x00690145: "PDSxxn",
    0x006A01E9: "DPSax",
    0x006B178A: "PSDPSoaxxn",
    0x006C01E8: "SDPax",
    0x006D1785: "PDSPDoaxxn",
    0x006E1E28: "SDPSnoax",
    0x006F0C65: "PDSxnan",
    0x00700CC5: "PDSana",
    0x00711D5C: "SSDxPDxaxn",
    0x00720648: "SDPSxox",
    0x00730E28: "SDPnoan",
    0x00740646: "DSPDxox",
    0x00750E26: "DSPnoan",
    0x00761B28: "SDPSnaox",
    0x007700E6: "DSan",
    0x007801E5: "PDSax",
    0x00791786: "DSPDSoaxxn",
    0x007A1E29: "DPSDnoax",
    0x007B0C68: "SDPxnan",
    0x007C1E24: "SPDSnoax",
    0x007D0C69: "DPSxnan",
    0x007E0955: "SPxDSxo",
    0x007F03C9: "DPSaan",
    0x008003E9: "DPSaa",
    0x00810975: "SPxDSxon",
    0x00820C49: "DPSxna",
    0x00831E04: "SPDSnoaxn",
    0x00840C48: "SDPxna",
    0x00851E05: "PDSPnoaxn",
    0x008617A6: "DSPDSoaxx",
    0x008701C5: "PDSaxn",
    0x008800C6: "DSa",
    0x00891B08: "SDPSnaoxn",
    0x008A0E06: "DSPnoa",
    0x008B0666: "DSPDxoxn",
    0x008C0E08: "SDPnoa",
    0x008D0668: "SDPSxoxn",
    0x008E1D7C: "SSDxPDxax",
    0x008F0CE5: "PDSanan",
    0x00900C45: "PDSxna",
    0x00911E08: "SDPSnoaxn",
    0x009217A9: "DPSDPoaxx",
    0x009301C4: "SPDaxn",
    0x009417AA: "PSDPSoaxx",
    0x009501C9: "DPSaxn",
    0x00960169: "DPSxx",
    0x0097588A: "PSDPSonoxx",
    0x00981888: "SDPSonoxn",
    0x00990066: "DSxn",
    0x009A0709: "DPSnax",
    0x009B07A8: "SDPSoaxn",
    0x009C0704: "SPDnax",
    0x009D07A6: "DSPDoaxn",
    0x009E16E6: "DSPDSaoxx",
    0x009F0345: "PDSxan",
    0x00A000C9: "DPa",
    0x00A11B05: "PDSPnaoxn",
    0x00A20E09: "DPSnoa",
    0x00A30669: "DPSDxoxn",
    0x00A41885: "PDSPonoxn",
    0x00A50065: "PDxn",
    0x00A60706: "DSPnax",
    0x00A707A5: "PDSPoaxn",
    0x00A803A9: "DPSoa",
    0x00A90189: "DPSoxn",
    0x00AA0029: "D",
    0x00AB0889: "DPSono",
    0x00AC0744: "SPDSxax",
    0x00AD06E9: "DPSDaoxn",
    0x00AE0B06: "DSPnao",
    0x00AF0229: "DPno",
    0x00B00E05: "PDSnoa",
    0x00B10665: "PDSPxoxn",
    0x00B21974: "SSPxDSxox",
    0x00B30CE8: "SDPanan",
    0x00B4070A: "PSDnax",
    0x00B507A9: "DPSDoaxn",
    0x00B616E9: "DPSDPaoxx",
    0x00B70348: "SDPxan",
    0x00B8074A: "PSDPxax",
    0x00B906E6: "DSPDaoxn",
    0x00BA0B09: "DPSnao",
    0x00BB0226: "DSno",
    0x00BC1CE4: "SPDSanax",
    0x00BD0D7D: "SDxPDxan",
    0x00BE0269: "DPSxo",
    0x00BF08C9: "DPSano",
    0x00C000CA: "PSa",
    0x00C11B04: "SPDSnaoxn",
    0x00C21884: "SPDSonoxn",
    0x00C3006A: "PSxn",
    0x00C40E04: "SPDnoa",
    0x00C50664: "SPDSxoxn",
    0x00C60708: "SDPnax",
    0x00C707AA: "PSDPoaxn",
    0x00C803A8: "SDPoa",
    0x00C90184: "SPDoxn",
    0x00CA0749: "DPSDxax",
    0x00CB06E4: "SPDSaoxn",
    0x00CC0020: "S",
    0x00CD0888: "SDPono",
    0x00CE0B08: "SDPnao",
    0x00CF0224: "SPno",
    0x00D00E0A: "PSDnoa",
    0x00D1066A: "PSDPxoxn",
    0x00D20705: "PDSnax",
    0x00D307A4: "SPDSoaxn",
    0x00D41D78: "SSPxPDxax",
    0x00D50CE9: "DPSanan",
    0x00D616EA: "PSDPSaoxx",
    0x00D70349: "DPSxan",
    0x00D80745: "PDSPxax",
    0x00D906E8: "SDPSaoxn",
    0x00DA1CE9: "DPSDanax",
    0x00DB0D75: "SPxDSxan",
    0x00DC0B04: "SPDnao",
    0x00DD0228: "SDno",
    0x00DE0268: "SDPxo",
    0x00DF08C8: "SDPano",
    0x00E003A5: "PDSoa",
    0x00E10185: "PDSoxn",
    0x00E20746: "DSPDxax",
    0x00E306EA: "PSDPaoxn",
    0x00E40748: "SDPSxax",
    0x00E506E5: "PDSPaoxn",
    0x00E61CE8: "SDPSanax",
    0x00E70D79: "SPxPDxan",
    0x00E81D74: "SSPxDSxax",
    0x00E95CE6: "DSPDSanaxxn",
    0x00EA02E9: "DPSao",
    0x00EB0849: "DPSxno",
    0x00EC02E8: "SDPao",
    0x00ED0848: "SDPxno",
    0x00EE0086: "DSo",
    0x00EF0A08: "SDPnoo",
    0x00F00021: "P",
    0x00F10885: "PDSono",
    0x00F20B05: "PDSnao",
    0x00F3022A: "PSno",
    0x00F40B0A: "PSDnao",
    0x00F50225: "PDno",
    0x00F60265: "PDSxo",
    0x00F708C5: "PDSano",
    0x00F802E5: "PDSao",
    0x00F90845: "PDSxno",
    0x00FA0089: "DPo",
    0x00FB0A09: "DPSnoo",
    0x00FC008A: "PSo",
    0x00FD0A0A: "PSDnoo",
    0x00FE02A9: "DPSoo",
    0x00FF0062: "1"
}


class WMFStream(BinaryStream):
    def __init__(self, bytes):
        BinaryStream.__init__(self, bytes)

    def dump(self):
        print('<stream type="WMF" size="%d">' % self.size)
        wmfHeader = WmfHeader(self)
        wmfHeader.dump()

        MAXIMUM_RECORDS = 100000
        for i in range(1, MAXIMUM_RECORDS):
            size = self.getuInt32()
            id = self.getuInt16(pos=self.pos + 4)
            record = RecordType.get(id, ["INVALID"])
            type = record[0]
            # WmfHeader is already dumped
            if i:
                print('<record index="%s" type="%s">' % (i, type))
                if len(record) > 1:
                    handler = record[1](self)
                    handler.dump()
                else:
                    print('<todo/>')
                print('</record>')
            # META_EOF
            if type == "META_EOF":
                break
            if self.pos + size * 2 <= self.size:
                self.pos += size * 2
            else:
                print('<Error value="Unexpected end of file" />')
                break
        print('</stream>')


class WMFRecord(BinaryStream):
    def __init__(self, parent):
        BinaryStream.__init__(self, parent.bytes)
        self.parent = parent
        self.pos = parent.pos


class Rect(WMFRecord):
    """The Rect Object defines a rectangle."""
    def __init__(self, parent, name=None):
        WMFRecord.__init__(self, parent)
        if name:
            self.name = name
        else:
            self.name = "rect"

    def dump(self):
        print('<%s type="Rect">' % self.name)
        self.printAndSet("Left", self.readInt16(), hexdump=False)
        self.printAndSet("Top", self.readInt16(), hexdump=False)
        self.printAndSet("Right", self.readInt16(), hexdump=False)
        self.printAndSet("Bottom", self.readInt16(), hexdump=False)
        print('</%s>' % self.name)
        self.parent.pos = self.pos


class RectL(WMFRecord):
    """The RectL Object defines a rectangle."""

    def __init__(self, parent, name=None):
        WMFRecord.__init__(self, parent)
        if name:
            self.name = name
        else:
            self.name = "rectL"

    def dump(self):
        print('<%s type="RectL">' % self.name)
        self.printAndSet("Left", self.readInt32(), hexdump=False)
        self.printAndSet("Top", self.readInt32(), hexdump=False)
        self.printAndSet("Right", self.readInt32(), hexdump=False)
        self.printAndSet("Bottom", self.readInt32(), hexdump=False)
        print('</%s>' % self.name)
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
        print('<%s type="SizeL">' % self.name)
        self.printAndSet("cx", self.readuInt32(), hexdump=False)
        self.printAndSet("cy", self.readuInt32(), hexdump=False)
        print('</%s>' % self.name)
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
        print('<%s type="PointL">' % self.name)
        self.printAndSet("x", self.readInt32(), hexdump=False)
        self.printAndSet("y", self.readInt32(), hexdump=False)
        print('</%s>' % self.name)
        self.parent.pos = self.pos


class PointS(WMFRecord):
    """The PointS Object defines the x- and y-coordinates of a point."""

    def __init__(self, parent, name):
        WMFRecord.__init__(self, parent)
        self.name = name

    def dump(self):
        print('<%s type="PointS">' % self.name)
        self.printAndSet("x", self.readInt16(), hexdump=False)
        self.printAndSet("y", self.readInt16(), hexdump=False)
        print('</%s>' % self.name)
        self.parent.pos = self.pos


class ColorRef(WMFRecord):
    """The ColorRef Object defines the RGB color."""

    def __init__(self, parent, name):
        WMFRecord.__init__(self, parent)
        self.name = name

    def dump(self):
        print('<%s type="ColorRef">' % self.name)
        self.printAndSet("Red", self.readuInt8(), hexdump=False)
        self.printAndSet("Green", self.readuInt8(), hexdump=False)
        self.printAndSet("Blue", self.readuInt8(), hexdump=False)
        self.printAndSet("Reserved", self.readuInt8(), hexdump=False)
        print('</%s>' % self.name)
        self.parent.pos = self.pos


class WMFLineto(WMFRecord):
    """Draws a line from the current position up to, but not including, the
    specified point."""
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        posOrig = self.pos
        self.printAndSet("Type", self.readuInt32())
        self.printAndSet("Size", self.readuInt32(), hexdump=False)
        PointL(self, "Point").dump()


class WmfHeader(WMFRecord):
    """The HEADER record types define the starting points of WMF metafiles."""
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print('<wmfHeader>')
        self.header = Header(self)
        self.header.dump()
        self.parent.pos = self.pos
        print('</wmfHeader>')


class Header(WMFRecord):
    """The Header object defines the WMF metafile header."""
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        posOrig = self.pos
        self.Size = 18
        if PlaceableHeader(self).isPlaceable():
            PlaceableHeader(self).dump()
            self.Size += 22
        print("<header>")
        self.printAndSet("FileType", self.readuInt16(), dict=FileType)
        self.printAndSet("HeaderSize", self.readuInt16(), hexdump=False)
        self.printAndSet("Version", self.readuInt16(), hexdump=False)
        self.printAndSet("FileSize", self.readuInt32(), hexdump=False)
        self.printAndSet("NumObjects", self.readuInt16(), hexdump=False)
        self.printAndSet("MaxRecordSize", self.readuInt32(), hexdump=False)
        self.printAndSet("NumOfParams", self.readuInt16(), hexdump=False)
        print("</header>")
        assert self.pos - posOrig == self.Size
        self.parent.pos = self.pos


class PlaceableHeader(WMFRecord):
    """The header for the placeable WMF"""
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<placeableHeader>")
        self.header = Header(self)
        self.Size = 22
        pos = self.pos
        dataPos = self.pos
        self.printAndSet("Key", self.readuInt32(), dict=PlaceableKey)
        self.printAndSet("HWmf", self.readuInt16())
        Rect(self, "BoundingBox").dump()
        self.printAndSet("Inch", self.readuInt16(), hexdump=False)
        self.printAndSet("Reserved", self.readuInt32())
        self.printAndSet("Checksum", self.readuInt16())
        assert self.pos == dataPos + self.Size
        self.parent.pos = self.pos
        print("</placeableHeader>")

    def isPlaceable(self):
        self.header = Header(self)
        key = self.readuInt32()
        if key in PlaceableKey.keys():
            return True
        return False


class WmfSetviewportorgex(WMFRecord):
    """Defines the viewport origin."""

    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        posOrig = self.pos
        self.printAndSet("Type", self.readuInt32())
        self.printAndSet("Size", self.readuInt32(), hexdump=False)
        PointL(self, "Origin").dump()
        assert self.pos - posOrig == self.Size


"""The RecordType enumeration defines values that uniquely identify EMF records."""
RecordType = {
    0x0000: ['META_EOF'],
    0x0035: ['META_REALIZEPALETTE'],
    0x0037: ['META_SETPALENTRIES'],
    0x0102: ['META_SETBKMODE'],
    0x0103: ['META_SETMAPMODE'],
    0x0104: ['META_SETROP2'],
    0x0105: ['META_SETRELABS'],
    0x0106: ['META_SETPOLYFILLMODE'],
    0x0107: ['META_SETSTRETCHBLTMODE'],
    0x0108: ['META_SETTEXTCHAREXTRA'],
    0x0127: ['META_RESTOREDC'],
    0x0139: ['META_RESIZEPALETTE'],
    0x0142: ['META_DIBCREATEPATTERNBRUSH'],
    0x0149: ['META_SETLAYOUT'],
    0x0201: ['META_SETBKCOLOR'],
    0x0209: ['META_SETTEXTCOLOR'],
    0x0211: ['META_OFFSETVIEWPORTORG'],
    0x0213: ['META_LINETO', WMFLineto],
    0x0214: ['META_MOVETO'],
    0x0220: ['META_OFFSETCLIPRGN'],
    0x0228: ['META_FILLREGION'],
    0x0231: ['META_SETMAPPERFLAGS'],
    0x0234: ['META_SELECTPALETTE'],
    0x0324: ['META_POLYGON'],
    0x0325: ['META_POLYLINE'],
    0x020A: ['META_SETTEXTJUSTIFICATION'],
    0x020B: ['META_SETWINDOWORG'],
    0x020C: ['META_SETWINDOWEXT'],
    0x020D: ['META_SETVIEWPORTORG'],
    0x020E: ['META_SETVIEWPORTEXT'],
    0x020F: ['META_OFFSETWINDOWORG'],
    0x0410: ['META_SCALEWINDOWEXT'],
    0x0412: ['META_SCALEVIEWPORTEXT'],
    0x0415: ['META_EXCLUDECLIPRECT'],
    0x0416: ['META_INTERSECTCLIPRECT'],
    0x0418: ['META_ELLIPSE'],
    0x0419: ['META_FLOODFILL'],
    0x0429: ['META_FRAMEREGION'],
    0x0436: ['META_ANIMATEPALETTE'],
    0x0521: ['META_TEXTOUT'],
    0x0538: ['META_POLYPOLYGON'],
    0x0548: ['META_EXTFLOODFILL'],
    0x041B: ['META_RECTANGLE'],
    0x041F: ['META_SETPIXEL'],
    0x061C: ['META_ROUNDRECT'],
    0x061D: ['META_PATBLT'],
    0x001E: ['META_SAVEDC'],
    0x081A: ['META_PIE'],
    0x0B23: ['META_STRETCHBLT'],
    0x0626: ['META_ESCAPE'],
    0x012A: ['META_INVERTREGION'],
    0x012B: ['META_PAINTREGION'],
    0x012C: ['META_SELECTCLIPREGION'],
    0x012D: ['META_SELECTOBJECT'],
    0x012E: ['META_SETTEXTALIGN'],
    0x0817: ['META_ARC'],
    0x0830: ['META_CHORD'],
    0x0922: ['META_BITBLT'],
    0x0a32: ['META_EXTTEXTOUT'],
    0x0d33: ['META_SETDIBTODEV'],
    0x0940: ['META_DIBBITBLT'],
    0x0b41: ['META_DIBSTRETCHBLT'],
    0x0f43: ['META_STRETCHDIB'],
    0x01f0: ['META_DELETEOBJECT'],
    0x00f7: ['META_CREATEPALETTE'],
    0x01F9: ['META_CREATEPATTERNBRUSH'],
    0x02FA: ['META_CREATEPENINDIRECT'],
    0x02FB: ['META_CREATEFONTINDIRECT'],
    0x02FC: ['META_CREATEBRUSHINDIRECT'],
    0x06FF: ['META_CREATEREGION'],
}

# vim:set filetype=python shiftwidth=4 softtabstop=4 expandtab:
