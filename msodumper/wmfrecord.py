#!/usr/bin/env python3
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
#

from .binarystream import BinaryStream
import base64

PlaceableKey = {
    0x9ac6cdd7: "META_PLACEABLE",
}

BinaryRasterOperation = {
    0x0001: "R2_BLACK",
    0x0002: "R2_NOTMERGEPEN",
    0x0003: "R2_MASKNOTPEN",
    0x0004: "R2_NOTCOPYPEN",
    0x0005: "R2_MASKPENNOT",
    0x0006: "R2_NOT",
    0x0007: "R2_XORPEN",
    0x0008: "R2_NOTMASKPEN",
    0x0009: "R2_MASKPEN",
    0x000A: "R2_NOTXORPEN",
    0x000B: "R2_NOP",
    0x000C: "R2_MERGENOTPEN",
    0x000D: "R2_COPYPEN",
    0x000E: "R2_MERGEPENNOT",
    0x000F: "R2_MERGEPEN",
    0x0010: "R2_WHITE",
}

BitCount = {
    0x0000: "BI_BITCOUNT_0",
    0x0001: "BI_BITCOUNT_1",
    0x0004: "BI_BITCOUNT_2",
    0x0008: "BI_BITCOUNT_3",
    0x0010: "BI_BITCOUNT_4",
    0x0018: "BI_BITCOUNT_5",
    0x0020: "BI_BITCOUNT_6",
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

CharacterSet = {
    0x00000001: "DEFAULT_CHARSET",
    0x00000002: "SYMBOL_CHARSET",
    0x0000004D: "MAC_CHARSET",
    0x00000080: "SHIFTJIS_CHARSET",
    0x00000081: "HANGUL_CHARSET",
    0x00000082: "JOHAB_CHARSET",
    0x00000086: "GB2312_CHARSET",
    0x00000088: "CHINESEBIG5_CHARSET",
    0x000000A1: "GREEK_CHARSET",
    0x000000A2: "TURKISH_CHARSET",
    0x000000A3: "VIETNAMESE_CHARSET",
    0x000000B1: "HEBREW_CHARSET",
    0x000000B2: "ARABIC_CHARSET",
    0x000000BA: "BALTIC_CHARSET",
    0x000000CC: "RUSSIAN_CHARSET",
    0x000000DE: "THAI_CHARSET",
    0x000000EE: "EASTEUROPE_CHARSET",
    0x000000FF: "OEM_CHARSET",
}

ColorUsage = {
    0x0000: "DIB_RGB_COLORS",
    0x0001: "DIB_PAL_COLORS",
    0x0002: "DIB_PAL_INDICES",
}

Compression = {
    0x0000: "BI_RGB",
    0x0001: "BI_RLE8",
    0x0002: "BI_RLE4",
    0x0003: "BI_BITFIELDS",
    0x0004: "BI_JPEG",
    0x0005: "BI_PNG",
    0x000B: "BI_CMYK",
    0x000C: "BI_CMYKRLE8",
    0x000D: "BI_CMYKRLE4",
}

FamilyFont = {
    0x00: "FF_DONTCARE",
    0x01: "FF_ROMAN",
    0x02: "FF_SWISS",
    0x03: "FF_MODERN",
    0x04: "FF_SCRIPT",
    0x05: "FF_DECORATIVE",
}

FloodFill = {
    0x0000: "FLOODFILLBORDER",
    0x0001: "FLOODFILLSURFACE",
}

FontQuality = {
    0x00: "DEFAULT_QUALITY",
    0x01: "DRAFT_QUALITY",
    0x02: "PROOF_QUALITY",
    0x03: "NONANTIALIASED_QUALITY",
    0x04: "ANTIALIASED_QUALITY",
    0x05: "CLEARTYPE_QUALITY",
}

GamutMappingIntent = {
    0x00000008: "LCS_GM_ABS_COLORIMETRIC",
    0x00000001: "LCS_GM_BUSINESS",
    0x00000002: "LCS_GM_GRAPHICS",
    0x00000004: "LCS_GM_IMAGES",
}

ColorUsage = {
    0x0000: "DIB_RGB_COLORS",
    0x0001: "DIB_PAL_COLORS",
    0x0002: "DIB_PAL_INDICES",
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

Layout = {
    0x0000: "LAYOUT_LTR",
    0x0001: "LAYOUT_RTL",
    0x0008: "LAYOUT_BITMAPORIENTATIONPRESERVED",
}

LogicalColorSpace = {
    0x00000000: "LCS_CALIBRATED_RGB",
    0x73524742: "LCS_sRGB",
    0x57696E20: "LCS_WINDOWS_COLOR_SPACE",
}

LogicalColorSpaceV5 = {
    0x4C494E4B: "LCS_PROFILE_LINKED",
    0x4D424544: "LCS_PROFILE_EMBEDDED",
}

MapMode = {
    0x0001: "MM_TEXT",
    0x0002: "MM_LOMETRIC",
    0x0003: "MM_HIMETRIC",
    0x0004: "MM_LOENGLISH",
    0x0005: "MM_HIENGLISH",
    0x0006: "MM_TWIPS",
    0x0007: "MM_ISOTROPIC",
    0x0008: "MM_ANISOTROPIC",
}

MetafileEscapes = {
    0x0001: "NEWFRAME",
    0x0002: "ABORTDOC",
    0x0003: "NEXTBAND",
    0x0004: "SETCOLORTABLE",
    0x0005: "GETCOLORTABLE",
    0x0006: "FLUSHOUT",
    0x0007: "DRAFTMODE",
    0x0008: "QUERYESCSUPPORT",
    0x0009: "SETABORTPROC",
    0x000A: "STARTDOC",
    0x000B: "ENDDOC",
    0x000C: "GETPHYSPAGESIZE",
    0x000D: "GETPRINTINGOFFSET",
    0x000E: "GETSCALINGFACTOR",
    0x000F: "META_ESCAPE_ENHANCED_METAFILE",
    0x0010: "SETPENWIDTH",
    0x0011: "SETCOPYCOUNT",
    0x0012: "SETPAPERSOURCE",
    0x0013: "PASSTHROUGH",
    0x0014: "GETTECHNOLOGY",
    0x0015: "SETLINECAP",
    0x0016: "SETLINEJOIN",
    0x0017: "SETMITERLIMIT",
    0x0018: "BANDINFO",
    0x0019: "DRAWPATTERNRECT",
    0x001A: "GETVECTORPENSIZE",
    0x001B: "GETVECTORBRUSHSIZE",
    0x001C: "ENABLEDUPLEX",
    0x001D: "GETSETPAPERBINS",
    0x001E: "GETSETPRINTORIENT",
    0x001F: "ENUMPAPERBINS",
    0x0020: "SETDIBSCALING",
    0x0021: "EPSPRINTING",
    0x0022: "ENUMPAPERMETRICS",
    0x0023: "GETSETPAPERMETRICS",
    0x0025: "POSTSCRIPT_DATA",
    0x0026: "POSTSCRIPT_IGNORE",
    0x002A: "GETDEVICEUNITS",
    0x0100: "GETEXTENDEDTEXTMETRICS",
    0x0102: "GETPAIRKERNTABLE",
    0x0200: "EXTTEXTOUT",
    0x0201: "GETFACENAME",
    0x0202: "DOWNLOADFACE",
    0x0801: "METAFILE_DRIVER",
    0x0C01: "QUERYDIBSUPPORT",
    0x1000: "BEGIN_PATH",
    0x1001: "CLIP_TO_PATH",
    0x1002: "END_PATH",
    0x100E: "OPENCHANNEL",
    0x100F: "DOWNLOADHEADER",
    0x1010: "CLOSECHANNEL",
    0x1013: "POSTSCRIPT_PASSTHROUGH",
    0x1014: "ENCAPSULATED_POSTSCRIPT",
    0x1015: "POSTSCRIPT_IDENTIFY",
    0x1016: "POSTSCRIPT_INJECTION",
    0x1017: "CHECKJPEGFORMAT",
    0x1018: "CHECKPNGFORMAT",
    0x1019: "GET_PS_FEATURESETTING",
    0x101A: "MXDC_ESCAPE",
    0x11D8: "SPCLPASSTHROUGH2",
}

MetafileType = {
    0x0001: "MEMORYMETAFILE",
    0x0002: "DISKMETAFILE",
}

MetafileVersion = {
    0x0100: "METAVERSION100",
    0x0300: "METAVERSION300",
}

MixMode = {
    0x0001: "TRANSPARENT",
    0x0002: "OPAQUE",
}

OutPrecision = {
    0x00000000: "OUT_DEFAULT_PRECIS",
    0x00000001: "OUT_STRING_PRECIS",
    0x00000003: "OUT_STROKE_PRECIS",
    0x00000004: "OUT_TT_PRECIS",
    0x00000005: "OUT_DEVICE_PRECIS",
    0x00000006: "OUT_RASTER_PRECIS",
    0x00000007: "OUT_TT_ONLY_PRECIS",
    0x00000008: "OUT_OUTLINE_PRECIS",
    0x00000009: "OUT_SCREEN_OUTLINE_PRECIS",
    0x0000000A: "OUT_PS_ONLY_PRECIS",
}

PaletteEntryFlag = {
    0x01: "PC_RESERVED",
    0x02: "PC_EXPLICIT",
    0x04: "PC_NOCOLLAPSE",
}

PostScriptCap = {
    -2: "PostScriptNotSet",
    0: "PostScriptFlatCap",
    1: "PostScriptRoundCap",
    2: "PostScriptSquareCap",
}

PostScriptClipping = {
    0x0000: "CLIP_SAVE",
    0x0001: "CLIP_RESTORE",
    0x0002: "CLIP_INCLUSIVE",
}

PostScriptFeatureSetting = {
    0x00000000: "FEATURESETTING_NUP",
    0x00000001: "FEATURESETTING_OUTPUT",
    0x00000002: "FEATURESETTING_PSLEVEL",
    0x00000003: "FEATURESETTING_CUSTPAPER",
    0x00000004: "FEATURESETTING_MIRROR",
    0x00000005: "FEATURESETTING_NEGATIVE",
    0x00000006: "FEATURESETTING_PROTOCOL",
    0x00001000: "FEATURESETTING_PRIVATE_BEGIN",
    0x00001FFF: "FEATURESETTING_PRIVATE_END",
}

PostScriptJoin = {
    -2: "PostScriptNotSet",
    0: "PostScriptMiterJoin",
    1: "PostScriptRoundJoin",
    2: "PostScriptBevelJoin",
}

StretchMode = {
    0x0001: "BLACKONWHITE",
    0x0002: "WHITEONBLACK",
    0x0003: "COLORONCOLOR",
    0x0004: "HALFTONE",
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
            if (self.pos + size * 2) <= self.size:
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


class Lineto(WMFRecord):
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
        self.printAndSet("FileType", self.readuInt16(), dict=MetafileType)
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


class Setviewportorgex(WMFRecord):
    """Defines the viewport origin."""

    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        posOrig = self.pos
        self.printAndSet("Type", self.readuInt32())
        self.printAndSet("Size", self.readuInt32(), hexdump=False)
        PointL(self, "Origin").dump()
        assert self.pos - posOrig == self.Size


class RealizePalette(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class SetPaletteEntries(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class SetBkMode(WMFRecord):
    """The SetBkMode record is used to define the background raster operation
       mix mode (pens, text, hatched brushes, and inside of filled objects
       with background colors)"""
    def __init__(self, parent, name=None):
        WMFRecord.__init__(self, parent)
        if name:
            self.name = name
        else:
            self.name = "setbkmode"

    def dump(self):
        pass
        dataPos = self.pos
        print('<%s type="SetBkMode">' % self.name)
        self.printAndSet("RecordSize", self.readuInt32(), hexdump=False)
        self.printAndSet("RecordFunction", self.readuInt16(), hexdump=True)
        self.printAndSet("BkMode", self.readuInt16(), hexdump=False)
        # Check optional reserved value if the size shows that it exists
        if self.RecordSize == 5:
            self.printAndSet("Reserved", self.readuInt16(), hexdump=False)
        print('</%s>' % self.name)
        assert self.pos == dataPos + self.RecordSize * 2


class SetMapMode(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class SetROP2(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class SetRelAbs(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class SetPolyFillMode(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class SetStretchBltMode(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class SetTextCharacterExtra(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class RestoreDC:
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class ResizePalette(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class CreateDIBPatternBrush(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class SetLayout(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class SetTextColor(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class SetBkColor(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class SetTextColor(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class MoveTo(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class OffsetClipRgn(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class FillRgn(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class SetMapperFlags(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class SelectPalette(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class Polygon(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class Polyline(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class SetTextJustification(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class SetWindowOrg(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class SetWindowExt:
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class SetViewportOrg(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class SetViewportExt(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class OffsetWindowOrg(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class ScaleWindowExt(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class ScaleViewportExt(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class ExcludeClipRect(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class IntersectClipRect(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class Ellipse(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class FrameRgn(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class AnimatePalette(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class TextOut(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class PolyPolygon(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class ExtFloodFill(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class Rectangle(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class SetPixel(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class RoundRect(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class SetPixel(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class SetPixel(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class SetPixel(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class RoundRect(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class SetPixel(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class PatBlt(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class RoundRect(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class SaveDC(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class SaveDC(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class Pie(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class StretchBlt(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class Escape(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class InvertRgn(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class PaintRgn(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class SelectClipRgn(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class SelectClipRgn(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class SelectObject(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class SetTextAlign(WMFRecord):
    """The SetTextAlign record is used to define the text alignment"""
    def __init__(self, parent, name=None):
        WMFRecord.__init__(self, parent)
        if name:
            self.name = name
        else:
            self.name = "settextalign"

    def dump(self):
        dataPos = self.pos
        print('<%s type="SetTextAlign">' % self.name)
        self.printAndSet("RecordSize", self.readuInt32(), hexdump=False)
        self.printAndSet("RecordFunction", self.readuInt16(), hexdump=True)
        self.printAndSet("TextAlignmentMode", self.readuInt16(), hexdump=False)
        # Check optional reserved value if the size shows that it exists
        if self.RecordSize == 5:
            self.printAndSet("Reserved", self.readuInt16(), hexdump=False)
        print('</%s>' % self.name)
        assert self.pos == dataPos + self.RecordSize * 2


class Arc(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class Chord(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class BitBlt(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class ExtTextOut(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class SetDIBitsToDevice(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class StretchDIBits(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class StretchDIBits(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class StretchDIBits(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class DeleteObject(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class CreatePalette(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class CreatePatternBrush(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class CreatePenIndirect(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class CreateFontIndirect(WMFRecord):
    """The CreateFontIndirect record is used to create a font object"""
    def __init__(self, parent, name=None):
        WMFRecord.__init__(self, parent)
        if name:
            self.name = name
        else:
            self.name = "createfontindirect"

    def dump(self):
        dataPos = self.pos
        print('<%s type="CreateFontIndirect">' % self.name)
        self.printAndSet("RecordSize", self.readuInt32(), hexdump=False)
        self.printAndSet("RecordFunction", self.readuInt16(), hexdump=True)
        # Check optional reserved value if the size shows that it exists
        if self.RecordSize > 3:
            Font(self, "Font").dump()
        print('</%s>' % self.name)
        # RecordSize is described in words, so we should double for bytes
        assert self.pos == dataPos + self.RecordSize * 2


class Font(WMFRecord):
    """The Font object describes a logical font and its attributes"""
    def __init__(self, parent, name=None):
        WMFRecord.__init__(self, parent)
        if name:
            self.name = name
        else:
            self.name = "Font"

    def dump(self):
        dataPos = self.pos
        print('<%s type="Font">' % self.name)
        self.printAndSet("Height", self.readInt16(), hexdump=False)
        self.printAndSet("Width", self.readInt16(), hexdump=False)
        self.printAndSet("Escapement", self.readInt16(), hexdump=False)
        self.printAndSet("Orientation", self.readInt16(), hexdump=False)
        self.printAndSet("Weight", self.readInt16(), hexdump=False)
        self.printAndSet("Italic", self.readuInt8(), hexdump=False)
        self.printAndSet("Underline", self.readuInt8(), hexdump=False)
        self.printAndSet("StrikeOut", self.readuInt8(), hexdump=False)
        self.printAndSet("CharSet", self.readuInt8(), hexdump=False)
        self.printAndSet("OutPrecision", self.readuInt8(), hexdump=False)
        self.printAndSet("ClipPrecision", self.readuInt8(), hexdump=False)
        self.printAndSet("Quality", self.readuInt8(), hexdump=False)
        self.printAndSet("PitchAndFamily", self.readuInt8(), hexdump=False)
        name = self.readBytes(32)
        self.FaceName = ""
        # Use characters until null byte
        for i in range(32):
            if name[i] == 0:
                break
            self.FaceName += chr(name[i])
        print('<FaceName value="%s"/>' % self.FaceName)
        print('</%s>' % self.name)
        assert self.pos == dataPos + 50
        self.parent.pos = self.pos


class CreateBrushIndirect(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


class CreateRectRgn(WMFRecord):
    def __init__(self, parent):
        WMFRecord.__init__(self, parent)

    def dump(self):
        print("<todo/>")
        pass


# GDI Functions: https://docs.microsoft.com/en-us/windows/win32/api/_gdi/
# Wine API / GDI: https://source.winehq.org/WineAPI/gdi.html
"""The RecordType enumeration defines values that uniquely identify WMF records."""
RecordType = {
    0x0000: ['META_EOF'],
    0x0035: ['META_REALIZEPALETTE', RealizePalette],
    0x0037: ['META_SETPALENTRIES', SetPaletteEntries],
    0x0102: ['META_SETBKMODE', SetBkMode],
    0x0103: ['META_SETMAPMODE', SetMapMode],
    0x0104: ['META_SETROP2', SetROP2],
    0x0105: ['META_SETRELABS', SetRelAbs],
    0x0106: ['META_SETPOLYFILLMODE', SetPolyFillMode],
    0x0107: ['META_SETSTRETCHBLTMODE', SetStretchBltMode],
    0x0108: ['META_SETTEXTCHAREXTRA', SetTextCharacterExtra],
    0x0127: ['META_RESTOREDC', RestoreDC],
    0x0139: ['META_RESIZEPALETTE', ResizePalette],
    0x0142: ['META_DIBCREATEPATTERNBRUSH', CreateDIBPatternBrush],
    0x0149: ['META_SETLAYOUT', SetLayout],
    0x0201: ['META_SETBKCOLOR', SetBkColor],
    0x0209: ['META_SETTEXTCOLOR', SetTextColor],
    0x0211: ['META_OFFSETVIEWPORTORG', Setviewportorgex],
    0x0213: ['META_LINETO', Lineto],
    0x0214: ['META_MOVETO', MoveTo],
    0x0220: ['META_OFFSETCLIPRGN', OffsetClipRgn],
    0x0228: ['META_FILLREGION', FillRgn],
    0x0231: ['META_SETMAPPERFLAGS', SetMapperFlags],
    0x0234: ['META_SELECTPALETTE', SelectPalette],
    0x0324: ['META_POLYGON', Polygon],
    0x0325: ['META_POLYLINE', Polyline],
    0x020A: ['META_SETTEXTJUSTIFICATION', SetTextJustification],
    0x020B: ['META_SETWINDOWORG', SetWindowOrg],
    0x020C: ['META_SETWINDOWEXT', SetWindowExt],
    0x020D: ['META_SETVIEWPORTORG', SetViewportOrg],
    0x020E: ['META_SETVIEWPORTEXT', SetViewportExt],
    0x020F: ['META_OFFSETWINDOWORG', OffsetWindowOrg],
    0x0410: ['META_SCALEWINDOWEXT', ScaleWindowExt],
    0x0412: ['META_SCALEVIEWPORTEXT', ScaleViewportExt],
    0x0415: ['META_EXCLUDECLIPRECT', ExcludeClipRect],
    0x0416: ['META_INTERSECTCLIPRECT', IntersectClipRect],
    0x0418: ['META_ELLIPSE', Ellipse],
    0x0419: ['META_FLOODFILL', FloodFill],
    0x0429: ['META_FRAMEREGION', FrameRgn],
    0x0436: ['META_ANIMATEPALETTE', AnimatePalette],
    0x0521: ['META_TEXTOUT', TextOut],
    0x0538: ['META_POLYPOLYGON', PolyPolygon],
    0x0548: ['META_EXTFLOODFILL', ExtFloodFill],
    0x041B: ['META_RECTANGLE', Rectangle],
    0x041F: ['META_SETPIXEL', SetPixel],
    0x061C: ['META_ROUNDRECT', RoundRect],
    0x061D: ['META_PATBLT', PatBlt],
    0x001E: ['META_SAVEDC', SaveDC],
    0x081A: ['META_PIE', Pie],
    0x0B23: ['META_STRETCHBLT', StretchBlt],
    0x0626: ['META_ESCAPE', Escape],
    0x012A: ['META_INVERTREGION', InvertRgn],
    0x012B: ['META_PAINTREGION', PaintRgn],
    0x012C: ['META_SELECTCLIPREGION', SelectClipRgn],
    0x012D: ['META_SELECTOBJECT', SelectObject],
    0x012E: ['META_SETTEXTALIGN', SetTextAlign],
    0x0817: ['META_ARC', Arc],
    0x0830: ['META_CHORD', Chord],
    0x0922: ['META_BITBLT', BitBlt],
    0x0a32: ['META_EXTTEXTOUT', ExtTextOut],
    0x0d33: ['META_SETDIBTODEV', SetDIBitsToDevice],
    0x0940: ['META_DIBBITBLT', BitBlt],
    0x0b41: ['META_DIBSTRETCHBLT', StretchBlt],
    0x0f43: ['META_STRETCHDIB', StretchDIBits],
    0x01f0: ['META_DELETEOBJECT', DeleteObject],
    0x00f7: ['META_CREATEPALETTE', CreatePalette],
    0x01F9: ['META_CREATEPATTERNBRUSH', CreatePatternBrush],
    0x02FA: ['META_CREATEPENINDIRECT', CreatePenIndirect],
    0x02FB: ['META_CREATEFONTINDIRECT', CreateFontIndirect],
    0x02FC: ['META_CREATEBRUSHINDIRECT', CreateBrushIndirect],
    0x06FF: ['META_CREATEREGION', CreateRectRgn],
}

# vim:set filetype=python shiftwidth=4 softtabstop=4 expandtab:
