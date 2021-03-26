#!/usr/bin/env python2
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
#

from .binarystream import BinaryStream
from . import wmfrecord
import base64


# The FormatSignature enumeration defines valuesembedded data in EMF records.
FormatSignature = {
    0x464D4520: "ENHMETA_SIGNATURE",
    0x46535045: "EPS_SIGNATURE"
}

RegionMode = {
    0x01: "RGN_AND",
    0x02: "RGN_OR",
    0x03: "RGN_XOR",
    0x04: "RGN_DIFF",
    0x05: "RGN_COPY"
}

DIBColors = {
    0x00: "DIB_RGB_COLORS",
    0x01: "DIB_PAL_COLORS",
    0x02: "DIB_PAL_INDICES"
}

# The PenStyle enumeration defines the attributes of pens that can be used in graphics operations.
PenStyle = {
    0x00000000: "PS_COSMETIC",
    0x00000000: "PS_ENDCAP_ROUND",
    0x00000000: "PS_JOIN_ROUND",
    0x00000000: "PS_SOLID",
    0x00000001: "PS_DASH",
    0x00000002: "PS_DOT",
    0x00000003: "PS_DASHDOT",
    0x00000004: "PS_DASHDOTDOT",
    0x00000005: "PS_NULL",
    0x00000006: "PS_INSIDEFRAME",
    0x00000007: "PS_USERSTYLE",
    0x00000008: "PS_ALTERNATE",
    0x00000100: "PS_ENDCAP_SQUARE",
    0x00000200: "PS_ENDCAP_FLAT",
    0x00001000: "PS_JOIN_BEVEL",
    0x00002000: "PS_JOIN_MITER",
    0x00010000: "PS_GEOMETRIC",
    # Additional combinations
    0x00010200: "PS_GEOMETRIC, PS_ENDCAP_FLAT",
    0x00011100: "PS_GEOMETRIC, PS_JOIN_BEVEL, PS_ENDCAP_SQUARE",
}


class EMFStream(BinaryStream):
    def __init__(self, bytes):
        BinaryStream.__init__(self, bytes)

    def dump(self):
        print('<stream type="EMF" size="%d">' % self.size)
        emrHeader = EmrHeader(self)
        emrHeader.dump()
        for i in range(emrHeader.header.Records):
            id = self.getuInt32()
            record = RecordType[id]
            type = record[0]
            size = self.getuInt32(pos=self.pos + 4)
            # EmrHeader is already dumped
            if i:
                print('<record index="%s" type="%s">' % (i, type))
                if len(record) > 1:
                    handler = record[1](self)
                    handler.dump()
                else:
                    print('<todo/>')
                print('</record>')
            self.pos += size
        print('</stream>')


class EMFRecord(BinaryStream):
    def __init__(self, parent):
        BinaryStream.__init__(self, parent.bytes)
        self.parent = parent
        self.pos = parent.pos


class EmrSavedc(EMFRecord):
    """This record saves the current state of the playback device context."""
    def __init__(self, parent):
        EMFRecord.__init__(self, parent)

    def dump(self):
        posOrig = self.pos
        self.printAndSet("Type", self.readuInt32())
        self.printAndSet("Size", self.readuInt32(), hexdump=False)
        assert self.pos - posOrig == self.Size


class EmrRestoredc(EMFRecord):
    """This record saves the current state of the playback device context."""
    def __init__(self, parent):
        EMFRecord.__init__(self, parent)

    def dump(self):
        posOrig = self.pos
        self.printAndSet("Type", self.readuInt32())
        self.printAndSet("Size", self.readuInt32(), hexdump=False)
        self.printAndSet("SavedDC", self.readInt32(), hexdump=False)
        assert self.pos - posOrig == self.Size


class EmrCreatebrushindirect(EMFRecord):
    """Defines a logical brush with a LogBrushEx object."""
    def __init__(self, parent):
        EMFRecord.__init__(self, parent)

    def dump(self):
        posOrig = self.pos
        self.printAndSet("Type", self.readuInt32())
        self.printAndSet("Size", self.readuInt32(), hexdump=False)
        self.printAndSet("ihBrush", self.readuInt32(), hexdump=False)
        LogBrushEx(self, "LogBrush").dump()
        assert self.pos - posOrig == self.Size


# The HatchStyle enumeration is an extension to the WMF HatchStyle enumeration.
EmfHatchStyle = {
    0x0006: "HS_SOLIDCLR",
    0x0007: "HS_DITHEREDCLR",
    0x0008: "HS_SOLIDTEXTCLR",
    0x0009: "HS_DITHEREDTEXTCLR",
    0x000A: "HS_SOLIDBKCLR",
    0x000B: "HS_DITHEREDBKCLR"
}
HatchStyle = dict(wmfrecord.HatchStyle.items() + EmfHatchStyle.items())


class LogBrushEx(EMFRecord):
    """The LogBrushEx object defines the style, color, and pattern of a device-independent brush."""
    def __init__(self, parent, name):
        EMFRecord.__init__(self, parent)
        self.name = name

    def dump(self):
        print('<%s>' % self.name)
        self.printAndSet("BrushStyle", self.readuInt32(), dict=wmfrecord.BrushStyle)
        wmfrecord.ColorRef(self, "Color").dump()
        self.printAndSet("BrushHatch", self.readuInt32(), dict=HatchStyle)
        print('</%s>' % self.name)
        self.parent.pos = self.pos


# Defines modes for using specified transform data to modify the world-space to page-space transform of the playback device context.
ModifyWorldTransformMode = {
    0x01: "MWT_IDENTITY",
    0x02: "MWT_LEFTMULTIPLY",
    0x03: "MWT_RIGHTMULTIPLY",
    0x04: "MWT_SET"
}


class EmrModifyworldtransform(EMFRecord):
    """Modifies the current world-space to page-space transform."""
    def __init__(self, parent):
        EMFRecord.__init__(self, parent)

    def dump(self):
        posOrig = self.pos
        self.printAndSet("Type", self.readuInt32())
        self.printAndSet("Size", self.readuInt32(), hexdump=False)
        XForm(self, "Xform").dump()
        self.printAndSet("ModifyWorldTransformMode", self.readuInt32(), dict=ModifyWorldTransformMode)
        assert self.pos - posOrig == self.Size


class XForm(EMFRecord):
    """The XForm object defines a two-dimensional, linear transform matrix."""
    def __init__(self, parent, name):
        EMFRecord.__init__(self, parent)
        self.name = name

    def dump(self):
        print('<%s>' % self.name)
        self.printAndSet("M11", self.readFloat32())
        self.printAndSet("M12", self.readFloat32())
        self.printAndSet("M21", self.readFloat32())
        self.printAndSet("M22", self.readFloat32())
        self.printAndSet("Dx", self.readFloat32())
        self.printAndSet("Dy", self.readFloat32())
        print('</%s>' % self.name)
        self.parent.pos = self.pos


# The ICMMode enumeration defines values that specify when to turn on and off ICM (Image Color Management).
ICMMode = {
    0x01: "ICM_OFF",
    0x02: "ICM_ON",
    0x03: "ICM_QUERY",
    0x04: "ICM_DONE_OUTSIDEDC"
}


class EmrSeticmmode(EMFRecord):
    """Specifies ICM to be enabled, disabled, or queried on the playback device context."""
    def __init__(self, parent):
        EMFRecord.__init__(self, parent)

    def dump(self):
        posOrig = self.pos
        self.printAndSet("Type", self.readuInt32())
        self.printAndSet("Size", self.readuInt32(), hexdump=False)
        self.printAndSet("ICMMode", self.readuInt32(), dict=ICMMode)
        assert self.pos - posOrig == self.Size


# The FormatSignature enumeration defines values that are used to identify the format of embedded
# data in EMF records.
FormatSignature = {
    0x464D4520: "ENHMETA_SIGNATURE",
    0x46535045: "EPS_SIGNATURE",
    0x50444620: "PDF ",  # not in [MS-EMF]
}


class EmrFormat(EMFRecord):
    """
    The EmrFormat object contains information that identifies the format of image data in an
    EMR_COMMENT_MULTIFORMATS record.
    """
    def __init__(self, parent, index):
        EMFRecord.__init__(self, parent)
        self.index = index

    def dump(self):
        print("<emrFormat index='%s'>" % self.index)
        self.printAndSet("Signature", self.readuInt32(), dict=FormatSignature)
        self.printAndSet("Version", self.readuInt32())
        self.printAndSet("SizeData", self.readuInt32(), hexdump=False)
        self.printAndSet("offData", self.readuInt32(), hexdump=False)
        print("</emrFormat>")


class EmrCommentMultiformats(EMFRecord):
    """The EMR_COMMENT_MULTIFORMATS record specifies an image in multiple graphics formats."""
    def __init__(self, parent):
        EMFRecord.__init__(self, parent)

    def dump(self):
        print("<emrCommentMultiFormats>")
        wmfrecord.RectL(self, "OutputRect").dump()
        self.printAndSet("CountFormats", self.readuInt32())
        for formatIndex in range(self.CountFormats):
            EmrFormat(self, formatIndex).dump()
        print("</emrCommentMultiFormats>")


# Defines the types of data that a public comment record can contain.
EmrCommentEnum = {
    0x80000001: "EMR_COMMENT_WINDOWS_METAFILE",
    0x00000002: "EMR_COMMENT_BEGINGROUP",
    0x00000003: "EMR_COMMENT_ENDGROUP",
    0x40000004: "EMR_COMMENT_MULTIFORMATS",
    0x00000040: "EMR_COMMENT_UNICODE_STRING",
    0x00000080: "EMR_COMMENT_UNICODE_END",
}


class EmrCommentPublic(EMFRecord):
    """The EMR_COMMENT_PUBLIC record types specify extensions to EMF processing."""
    def __init__(self, parent):
        EMFRecord.__init__(self, parent)

    def dump(self):
        print("<emrCommentPublic>")
        self.printAndSet("PublicCommentIdentifier", self.readuInt32(), dict=EmrCommentEnum)
        if self.PublicCommentIdentifier == 0x40000004:  # EMR_COMMENT_MULTIFORMATS
            EmrCommentMultiformats(self).dump()
        print("</emrCommentPublic>")


class EMFPlusRecordHeader(EMFRecord):
    """The common start of EmfPlus* records."""
    def __init__(self, parent):
        EMFRecord.__init__(self, parent)

    def peek(self):
        print("<emfPlusRecordHeader>")
        id = self.readuInt16()
        record = RecordType[id]
        type = record[0]
        print('<Type value="%s" name="%s"/>' % (hex(id), type))
        self.Type = id
        self.printAndSet("Flags", self.readuInt16())
        self.printAndSet("Size", self.readuInt32())
        self.printAndSet("DataSize", self.readuInt32())
        print("</emfPlusRecordHeader>")


class EmfPlusHeader(EMFRecord):
    """The EmfPlusHeader record specifies the start of EMF+ data in the metafile."""
    def __init__(self, parent, header):
        EMFRecord.__init__(self, parent)
        self.header = header

    def dump(self):
        print("<emfPlusHeader>")
        pos = self.pos
        self.pos += 12  # header size
        dataPos = self.pos
        self.printAndSet("Version", self.readuInt32())
        self.printAndSet("EmfPlusFlags", self.readuInt32())
        self.printAndSet("LogicalDpiX", self.readuInt32())
        self.printAndSet("LogicalDpiY", self.readuInt32())
        assert self.pos == dataPos + self.header.DataSize
        self.parent.pos = pos + self.header.Size
        print("</emfPlusHeader>")


class EmrCommentEmfplus(EMFRecord):
    """The EMR_COMMENT_EMFPLUS record contains embedded EMF+ records."""
    def __init__(self, parent):
        EMFRecord.__init__(self, parent)

    def dump(self):
        size = self.parent.DataSize - 4  # CommentIdentifier doesn't count
        print("<emrCommentEmfplus size='%s'>" % size)
        end = self.pos + size
        while self.pos < end:
            header = EMFPlusRecordHeader(self)
            header.peek()
            if header.Type == 0x4001:  # EmfPlusHeader
                EmfPlusHeader(self, header).dump()
            else:
                print('<todo what="EmrCommentEmfplus::dump(): handle %s"/>' % RecordType[header.Type][0])
                self.pos += header.Size
        print("</emrCommentEmfplus>")


CommentIdentifier = {
    0x00000000: "EMR_COMMENT_EMFSPOOL",
    0x2B464D45: "EMR_COMMENT_EMFPLUS",
    0x43494447: "EMR_COMMENT_PUBLIC",
}


class EmrComment(EMFRecord):
    """The EMR_COMMENT record contains arbitrary private data."""
    def __init__(self, parent):
        EMFRecord.__init__(self, parent)

    def dump(self):
        self.printAndSet("Type", self.readuInt32())
        self.printAndSet("Size", self.readuInt32(), hexdump=False)
        self.printAndSet("DataSize", self.readuInt32(), hexdump=False)
        self.printAndSet("CommentIdentifier", self.readuInt32(), dict=CommentIdentifier, default="")
        if self.CommentIdentifier == 0x00000000:  # EMR_COMMENT_EMFSPOOL
            print('<todo what="EmrComment::dump(): handle EMR_COMMENT_EMFSPOOL"/>')
        elif self.CommentIdentifier == 0x2B464D45:  # EMR_COMMENT_EMFPLUS
            EmrCommentEmfplus(self).dump()
        elif self.CommentIdentifier == 0x43494447:  # EMR_COMMENT_PUBLIC
            EmrCommentPublic(self).dump()
        else:
            print('<todo what="EmrComment::dump(): handle EMR_COMMENT: %s"/>' % hex(self.CommentIdentifier))


class EmrSetviewportorgex(EMFRecord):
    """Defines the viewport origin."""
    def __init__(self, parent):
        EMFRecord.__init__(self, parent)

    def dump(self):
        posOrig = self.pos
        self.printAndSet("Type", self.readuInt32())
        self.printAndSet("Size", self.readuInt32(), hexdump=False)
        wmfrecord.PointL(self, "Origin").dump()
        assert self.pos - posOrig == self.Size


# Defines values that specify how to calculate the region of a polygon that is to be filled.
PolygonFillMode = {
    0x01: "ALTERNATE",  # Selects alternate mode (fills the area between odd-numbered and even-numbered polygon sides on each scan line).
    0x02: "WINDING"     # Selects winding mode (fills any region with a nonzero winding value).
}


class EmrSetpolyfillmode(EMFRecord):
    """Defines polygon fill mode."""
    def __init__(self, parent):
        EMFRecord.__init__(self, parent)

    def dump(self):
        posOrig = self.pos
        self.printAndSet("Type", self.readuInt32())
        self.printAndSet("Size", self.readuInt32(), hexdump=False)
        self.printAndSet("PolygonFillMode", self.readuInt32(), dict=PolygonFillMode)
        assert self.pos - posOrig == self.Size


# Used to specify how color data is added to or removed from bitmaps that are
# stretched or compressed.
StretchMode = {
    0x01: "STRETCH_ANDSCANS",
    0x02: "STRETCH_ORSCANS",
    0x03: "STRETCH_DELETESCANS",
    0x04: "STRETCH_HALFTONE",
}


class EmrSetstretchbltmode(EMFRecord):
    """Specifies bitmap stretch mode."""
    def __init__(self, parent):
        EMFRecord.__init__(self, parent)

    def dump(self):
        posOrig = self.pos
        self.printAndSet("Type", self.readuInt32())
        self.printAndSet("Size", self.readuInt32(), hexdump=False)
        self.printAndSet("StretchMode", self.readuInt32(), dict=StretchMode)
        assert self.pos - posOrig == self.Size


class EmrExtselectcliprgn(EMFRecord):
    """Combines the specified region with the current clip region using the specified mode."""
    def __init__(self, parent):
        EMFRecord.__init__(self, parent)

    def dump(self):
        posOrig = self.pos
        self.printAndSet("Type", self.readuInt32())
        self.printAndSet("Size", self.readuInt32(), hexdump=False)
        self.printAndSet("RgnDataSize", self.readuInt32())
        self.printAndSet("RegionMode", self.readuInt32(), dict=RegionMode)
        RegionData(self, "RgnData", self.RgnDataSize).dump()
        assert self.pos - posOrig == self.Size


# Specifies the indexes of predefined logical graphics objects that can be used in graphics operations.
StockObject = {
    0x80000000: "WHITE_BRUSH",
    0x80000001: "LTGRAY_BRUSH",
    0x80000002: "GRAY_BRUSH",
    0x80000003: "DKGRAY_BRUSH",
    0x80000004: "BLACK_BRUSH",
    0x80000005: "NULL_BRUSH",
    0x80000006: "WHITE_PEN",
    0x80000007: "BLACK_PEN",
    0x80000008: "NULL_PEN",
    0x8000000A: "OEM_FIXED_FONT",
    0x8000000B: "ANSI_FIXED_FONT",
    0x8000000C: "ANSI_VAR_FONT",
    0x8000000D: "SYSTEM_FONT",
    0x8000000E: "DEVICE_DEFAULT_FONT",
    0x8000000F: "DEFAULT_PALETTE",
    0x80000010: "SYSTEM_FIXED_FONT",
    0x80000011: "DEFAULT_GUI_FONT",
    0x80000012: "DC_BRUSH",
    0x80000013: "DC_PEN"
}


class EmrSelectobject(EMFRecord):
    """Combines the specified region with the current clip region using the specified mode."""
    def __init__(self, parent):
        EMFRecord.__init__(self, parent)

    def dump(self):
        posOrig = self.pos
        self.printAndSet("Type", self.readuInt32())
        self.printAndSet("Size", self.readuInt32(), hexdump=False)
        ihObject = self.getuInt32(pos=self.pos)
        if ihObject < 0x80000000:
            self.printAndSet("ihObject", self.readuInt32())
        else:
            self.printAndSet("ihObject", self.readuInt32(), dict=StockObject)
        assert self.pos - posOrig == self.Size


class EmrDeleteobject(EMFRecord):
    """Specifies the index of the object to be deleted from the EMF Object
    Table."""
    def __init__(self, parent):
        EMFRecord.__init__(self, parent)

    def dump(self):
        posOrig = self.pos
        self.printAndSet("Type", self.readuInt32())
        self.printAndSet("Size", self.readuInt32(), hexdump=False)
        ihObject = self.getuInt32(pos=self.pos)
        if ihObject < 0x80000000:
            self.printAndSet("ihObject", self.readuInt32())
        else:
            self.printAndSet("ihObject", self.readuInt32(), dict=StockObject)
        assert self.pos - posOrig == self.Size


class EmrPolygon16(EMFRecord):
    """Draws a polygon consisting of two or more vertexes connected by straight lines."""
    def __init__(self, parent):
        EMFRecord.__init__(self, parent)

    def dump(self):
        posOrig = self.pos
        self.printAndSet("Type", self.readuInt32())
        self.printAndSet("Size", self.readuInt32(), hexdump=False)
        wmfrecord.RectL(self, "Bounds").dump()
        self.printAndSet("Count", self.readuInt32(), hexdump=False)
        print('<aPoints>')
        for i in range(self.Count):
            wmfrecord.PointS(self, "aPoint%d" % i).dump()
        print('</aPoints>')
        assert self.pos - posOrig == self.Size


class EmrPolypolygon16(EMFRecord):
    """Paints a series of closed polygons."""
    def __init__(self, parent):
        EMFRecord.__init__(self, parent)

    def dump(self):
        posOrig = self.pos
        self.printAndSet("Type", self.readuInt32())
        self.printAndSet("Size", self.readuInt32(), hexdump=False)
        wmfrecord.RectL(self, "Bounds").dump()
        self.printAndSet("NumberOfPolygons", self.readuInt32(), hexdump=False)
        self.printAndSet("Count", self.readuInt32(), hexdump=False)
        print('<PolygonPointCounts>')
        for i in range(self.NumberOfPolygons):
            self.printAndSet("PolygonPointCount%d" % i, self.readuInt32(), hexdump=False)
        print('</PolygonPointCounts>')
        print('<aPoints>')
        for i in range(self.Count):
            wmfrecord.PointS(self, "aPoint").dump()
        print('</aPoints>')
        assert self.pos - posOrig == self.Size


class EmrPolylineto16(EMFRecord):
    """Draws one or more straight lines based upon the current position."""
    def __init__(self, parent):
        EMFRecord.__init__(self, parent)

    def dump(self):
        posOrig = self.pos
        self.printAndSet("Type", self.readuInt32())
        self.printAndSet("Size", self.readuInt32(), hexdump=False)
        wmfrecord.RectL(self, "Bounds").dump()
        self.printAndSet("Count", self.readuInt32(), hexdump=False)
        print('<aPoints>')
        for i in range(self.Count):
            wmfrecord.PointS(self, "aPoint%d" % i).dump()
        print('</aPoints>')
        assert self.pos - posOrig == self.Size


class EmrPolybezierto16(EMFRecord):
    """Draws one or more Bezier curves based on the current position."""
    def __init__(self, parent):
        EMFRecord.__init__(self, parent)

    def dump(self):
        posOrig = self.pos
        self.printAndSet("Type", self.readuInt32())
        self.printAndSet("Size", self.readuInt32(), hexdump=False)
        wmfrecord.RectL(self, "Bounds").dump()
        self.printAndSet("Count", self.readuInt32(), hexdump=False)
        print('<aPoints>')
        for i in range(self.Count):
            wmfrecord.PointS(self, "aPoint%d" % i).dump()
        print('</aPoints>')
        assert self.pos - posOrig == self.Size


class EmrMovetoex(EMFRecord):
    """Specifies the coordinates of a new drawing position, in logical
    units."""
    def __init__(self, parent):
        EMFRecord.__init__(self, parent)

    def dump(self):
        posOrig = self.pos
        self.printAndSet("Type", self.readuInt32())
        self.printAndSet("Size", self.readuInt32(), hexdump=False)
        wmfrecord.PointL(self, "Offset").dump()
        assert self.pos - posOrig == self.Size


class EmrLineto(EMFRecord):
    """Draws a line from the current position up to, but not including, the
    specified point."""
    def __init__(self, parent):
        EMFRecord.__init__(self, parent)

    def dump(self):
        posOrig = self.pos
        self.printAndSet("Type", self.readuInt32())
        self.printAndSet("Size", self.readuInt32(), hexdump=False)
        wmfrecord.PointL(self, "Point").dump()
        assert self.pos - posOrig == self.Size


class EmrSelectclippath(EMFRecord):
    """Specifies the current path as a clipping region for the playback device
    context, combining the new region with any existing clipping region using
    the specified mode."""
    def __init__(self, parent):
        EMFRecord.__init__(self, parent)

    def dump(self):
        posOrig = self.pos
        self.printAndSet("Type", self.readuInt32())
        self.printAndSet("Size", self.readuInt32(), hexdump=False)
        self.printAndSet("RegionMode", self.readuInt32(), dict=RegionMode)
        assert self.pos - posOrig == self.Size


class EmrBeginpath(EMFRecord):
    """This record opens a path bracket in the current playback device context."""
    def __init__(self, parent):
        EMFRecord.__init__(self, parent)

    def dump(self):
        posOrig = self.pos
        self.printAndSet("Type", self.readuInt32())
        self.printAndSet("Size", self.readuInt32(), hexdump=False)
        assert self.pos - posOrig == self.Size


class EmrEndpath(EMFRecord):
    """This record closes a path bracket and selects the path defined by the bracket into the playback device context."""
    def __init__(self, parent):
        EMFRecord.__init__(self, parent)

    def dump(self):
        posOrig = self.pos
        self.printAndSet("Type", self.readuInt32())
        self.printAndSet("Size", self.readuInt32(), hexdump=False)
        assert self.pos - posOrig == self.Size


class EmrClosefigure(EMFRecord):
    """This record closes an open figure in a path."""
    def __init__(self, parent):
        EMFRecord.__init__(self, parent)

    def dump(self):
        posOrig = self.pos
        self.printAndSet("Type", self.readuInt32())
        self.printAndSet("Size", self.readuInt32(), hexdump=False)
        assert self.pos - posOrig == self.Size


class EmrFillpath(EMFRecord):
    """Closes any open figures in the current path and fills the path's
    interior with the current brush and polygon-filling mode."""
    def __init__(self, parent):
        EMFRecord.__init__(self, parent)

    def dump(self):
        posOrig = self.pos
        self.printAndSet("Type", self.readuInt32())
        self.printAndSet("Size", self.readuInt32(), hexdump=False)
        wmfrecord.RectL(self, "Bounds").dump()
        assert self.pos - posOrig == self.Size


class EmrStrokeandfillpath(EMFRecord):
    """The EMR_STROKEANDFILLPATH record closes any open figures in a path, strokes the outline of the
    path by using the current pen, and fills its interior by using the current brush."""
    def __init__(self, parent):
        EMFRecord.__init__(self, parent)

    def dump(self):
        posOrig = self.pos
        self.printAndSet("Type", self.readuInt32())
        self.printAndSet("Size", self.readuInt32(), hexdump=False)
        wmfrecord.RectL(self, "Bounds").dump()
        assert self.pos - posOrig == self.Size


class EmrBitblt(EMFRecord):
    """Specifies a block transfer of pixels from a source bitmap to a destination rectangle."""
    def __init__(self, parent):
        EMFRecord.__init__(self, parent)

    def dump(self):
        posOrig = self.pos
        self.printAndSet("Type", self.readuInt32())
        self.printAndSet("Size", self.readuInt32(), hexdump=False)
        wmfrecord.RectL(self, "Bounds").dump()
        self.printAndSet("xDest", self.readInt32(), hexdump=False)
        self.printAndSet("yDest", self.readInt32(), hexdump=False)
        self.printAndSet("cxDest", self.readInt32(), hexdump=False)
        self.printAndSet("cyDest", self.readInt32(), hexdump=False)
        self.printAndSet("BitBltRasterOperation", self.readuInt32(), dict=wmfrecord.RasterPolishMap)
        self.printAndSet("xSrc", self.readInt32(), hexdump=False)
        self.printAndSet("ySrc", self.readInt32(), hexdump=False)
        XForm(self, "XformSrc").dump()
        wmfrecord.ColorRef(self, "BkColorSrc").dump()
        self.printAndSet("UsageSrc", self.readInt32(), dict=DIBColors)
        self.printAndSet("offBmiSrc", self.readuInt32())
        self.printAndSet("cbBmiSrc", self.readuInt32())
        self.printAndSet("offBitsSrc", self.readuInt32())
        self.printAndSet("cbBitsSrc", self.readuInt32())
        assert self.pos - posOrig == self.Size


class LogPenEx(EMFRecord):
    """The LogPenEx object specifies the style, width, and color of an extended logical pen."""
    def __init__(self, parent, name):
        EMFRecord.__init__(self, parent)
        self.name = name

    def dump(self):
        print('<%s type="LogPenEx">' % self.name)
        self.printAndSet("PenStyle", self.readuInt32(), dict=PenStyle)
        self.printAndSet("Width", self.readuInt32())
        self.printAndSet("BrushStyle", self.readuInt32(), dict=wmfrecord.BrushStyle)
        wmfrecord.ColorRef(self, "ColorRef").dump()
        if self.BrushStyle == 0x0002:  # "BS_HATCHED"
            self.printAndSet("BrushHatch", self.readuInt32(), dict=wmfrecord.HatchStyle)
        else:
            self.printAndSet("BrushHatch", self.readuInt32())
        self.printAndSet("NumStyleEntries", self.readuInt32())
        if self.NumStyleEntries > 0:
            print('<todo what="LogPenEx::dump(): self.NumStyleEntries != 0"/>')
        print('</%s>' % self.name)
        self.parent.pos = self.pos


class EmrExtcreatepen(EMFRecord):
    """Defines an extended logical pen."""
    def __init__(self, parent):
        EMFRecord.__init__(self, parent)

    def dump(self):
        self.printAndSet("Type", self.readuInt32())
        self.printAndSet("Size", self.readuInt32(), hexdump=False)
        self.printAndSet("ihPen", self.readuInt32(), hexdump=False)
        self.printAndSet("offBmi", self.readuInt32(), hexdump=False)
        self.printAndSet("cbBmi", self.readuInt32(), hexdump=False)
        self.printAndSet("offBits", self.readuInt32(), hexdump=False)
        self.printAndSet("cbBits", self.readuInt32(), hexdump=False)
        LogPenEx(self, "elp").dump()
        if self.cbBmi:
            print('<todo what="LogPenEx::dump(): self.cbBmi != 0"/>')
        if self.cbBits:
            print('<todo what="LogPenEx::dump(): self.cbBits != 0"/>')


class EmrStretchdibits(EMFRecord):
    """Specifies a block transfer of pixels from a source bitmap to a
    destination rectangle."""
    def __init__(self, parent):
        EMFRecord.__init__(self, parent)

    def dump(self):
        posOrig = self.pos
        self.printAndSet("Type", self.readuInt32())
        self.printAndSet("Size", self.readuInt32(), hexdump=False)
        wmfrecord.RectL(self, "Bounds").dump()
        self.printAndSet("xDest", self.readInt32(), hexdump=False)
        self.printAndSet("yDest", self.readInt32(), hexdump=False)
        self.printAndSet("xSrc", self.readInt32(), hexdump=False)
        self.printAndSet("ySrc", self.readInt32(), hexdump=False)
        self.printAndSet("cxSrc", self.readInt32(), hexdump=False)
        self.printAndSet("cySrc", self.readInt32(), hexdump=False)
        self.printAndSet("offBmiSrc", self.readuInt32(), hexdump=False)
        self.printAndSet("cbBmiSrc", self.readuInt32(), hexdump=False)
        self.printAndSet("offBitsSrc", self.readuInt32(), hexdump=False)
        self.printAndSet("cbBitsSrc", self.readuInt32(), hexdump=False)
        self.printAndSet("UsageSrc", self.readInt32(), dict=DIBColors)
        self.printAndSet("BitBltRasterOperation", self.readuInt32(), dict=wmfrecord.RasterPolishMap)
        self.printAndSet("cxDest", self.readInt32(), hexdump=False)
        self.printAndSet("cyDest", self.readInt32(), hexdump=False)
        print('<BitmapBuffer>')
        if self.cbBmiSrc:
            self.pos = posOrig + self.offBmiSrc
            self.BmiSrc = self.readBytes(self.cbBmiSrc)
            print('<BmiSrc value="%s"/>' % base64.b64encode(self.BmiSrc))
        if self.cbBitsSrc:
            self.pos = posOrig + self.offBitsSrc
            self.BitsSrc = self.readBytes(self.cbBitsSrc)
            print('<BitsSrc value="%s"/>' % base64.b64encode(self.BitsSrc))
        print('</BitmapBuffer>')
        assert self.pos - posOrig == self.Size


class EmrEof(EMFRecord):
    """Indicates the end of the metafile and specifies a palette."""
    def __init__(self, parent):
        EMFRecord.__init__(self, parent)

    def dump(self):
        posOrig = self.pos
        self.printAndSet("Type", self.readuInt32())
        self.printAndSet("Size", self.readuInt32(), hexdump=False)
        self.printAndSet("nPalEntries", self.readuInt32(), hexdump=False)
        self.printAndSet("offPalEntries", self.readuInt32(), hexdump=False)
        if self.nPalEntries > 0:
            print('<todo what="EmrEof::dump(): handle nPalEntries > 0"/>')
        self.printAndSet("SizeLast", self.readuInt32(), hexdump=False)
        assert self.pos - posOrig == self.Size


class RegionData(EMFRecord):
    """The RegionData object specifies data that defines a region, which is made of non-overlapping rectangles."""
    def __init__(self, parent, name, size):
        EMFRecord.__init__(self, parent)
        self.name = name
        self.size = size

    def dump(self):
        print('<%s>' % self.name)
        header = RegionDataHeader(self)
        header.dump()
        for i in range(header.CountRects):
            wmfrecord.RectL(self, "Data%d" % i).dump()
        print('</%s>' % self.name)
        self.parent.pos = self.pos


class RegionDataHeader(EMFRecord):
    """The RegionDataHeader object describes the properties of a RegionData object."""
    def __init__(self, parent):
        EMFRecord.__init__(self, parent)

    def dump(self):
        self.printAndSet("Size", self.readuInt32())
        self.printAndSet("Type", self.readuInt32())
        self.printAndSet("CountRects", self.readuInt32())
        self.printAndSet("RgnSize", self.readuInt32())
        wmfrecord.RectL(self, "Bounds").dump()
        self.parent.pos = self.pos


class EmrHeader(EMFRecord):
    """The EMR_HEADER record types define the starting points of EMF metafiles."""
    def __init__(self, parent):
        EMFRecord.__init__(self, parent)

    def dump(self):
        print('<emrHeader>')
        self.printAndSet("Type", self.readuInt32())
        self.printAndSet("Size", self.readuInt32(), hexdump=False)
        self.header = Header(self)
        self.header.dump()
        if self.Size >= 100:
            HeaderExtension1(self).dump()
        if self.Size >= 108:
            HeaderExtension2(self).dump()
        print('</emrHeader>')


class Header(EMFRecord):
    """The Header object defines the EMF metafile header."""
    def __init__(self, parent):
        EMFRecord.__init__(self, parent)

    def dump(self):
        posOrig = self.pos
        print("<header>")
        wmfrecord.RectL(self, "Bounds").dump()
        wmfrecord.RectL(self, "Frame").dump()
        self.printAndSet("RecordSignature", self.readuInt32(), dict=FormatSignature)
        self.printAndSet("Version", self.readuInt32())
        self.printAndSet("Bytes", self.readuInt32(), hexdump=False)
        self.printAndSet("Records", self.readuInt32(), hexdump=False)
        self.printAndSet("Handles", self.readuInt16(), hexdump=False)
        self.printAndSet("Reserved", self.readuInt16(), hexdump=False)
        self.printAndSet("nDescription", self.readuInt32(), hexdump=False)
        self.printAndSet("offDescription", self.readuInt32(), hexdump=False)
        self.printAndSet("nPalEntries", self.readuInt32(), hexdump=False)
        wmfrecord.SizeL(self, "Device").dump()
        wmfrecord.SizeL(self, "Millimeters").dump()
        print("</header>")
        assert posOrig == self.pos - 80
        self.parent.pos = self.pos


class HeaderExtension1(EMFRecord):
    """The HeaderExtension1 object defines the first extension to the EMF metafile header."""
    def __init__(self, parent):
        EMFRecord.__init__(self, parent)

    def dump(self):
        posOrig = self.pos
        print("<headerExtension1>")
        self.printAndSet("cbPixelFormat", self.readuInt32(), hexdump=False)
        self.printAndSet("offPixelFormat", self.readuInt32(), hexdump=False)
        self.printAndSet("bOpenGL", self.readuInt32())
        print("</headerExtension1>")
        assert posOrig == self.pos - 12
        self.parent.pos = self.pos


class HeaderExtension2(EMFRecord):
    """The HeaderExtension1 object defines the first extension to the EMF metafile header."""
    def __init__(self, parent):
        EMFRecord.__init__(self, parent)

    def dump(self):
        posOrig = self.pos
        print("<headerExtension2>")
        self.printAndSet("MicrometersX", self.readuInt32(), hexdump=False)
        self.printAndSet("MicrometersY", self.readuInt32(), hexdump=False)
        print("</headerExtension2>")
        assert posOrig == self.pos - 8
        self.parent.pos = self.pos


"""The RecordType enumeration defines values that uniquely identify EMF records."""
RecordType = {
    0x00000001: ['EMR_HEADER'],
    0x00000002: ['EMR_POLYBEZIER'],
    0x00000003: ['EMR_POLYGON'],
    0x00000004: ['EMR_POLYLINE'],
    0x00000005: ['EMR_POLYBEZIERTO'],
    0x00000006: ['EMR_POLYLINETO'],
    0x00000007: ['EMR_POLYPOLYLINE'],
    0x00000008: ['EMR_POLYPOLYGON'],
    0x00000009: ['EMR_SETWINDOWEXTEX'],
    0x0000000A: ['EMR_SETWINDOWORGEX'],
    0x0000000B: ['EMR_SETVIEWPORTEXTEX'],
    0x0000000C: ['EMR_SETVIEWPORTORGEX', EmrSetviewportorgex],
    0x0000000D: ['EMR_SETBRUSHORGEX'],
    0x0000000E: ['EMR_EOF', EmrEof],
    0x0000000F: ['EMR_SETPIXELV'],
    0x00000010: ['EMR_SETMAPPERFLAGS'],
    0x00000011: ['EMR_SETMAPMODE'],
    0x00000012: ['EMR_SETBKMODE'],
    0x00000013: ['EMR_SETPOLYFILLMODE', EmrSetpolyfillmode],
    0x00000014: ['EMR_SETROP2'],
    0x00000015: ['EMR_SETSTRETCHBLTMODE', EmrSetstretchbltmode],
    0x00000016: ['EMR_SETTEXTALIGN'],
    0x00000017: ['EMR_SETCOLORADJUSTMENT'],
    0x00000018: ['EMR_SETTEXTCOLOR'],
    0x00000019: ['EMR_SETBKCOLOR'],
    0x0000001A: ['EMR_OFFSETCLIPRGN'],
    0x0000001B: ['EMR_MOVETOEX', EmrMovetoex],
    0x0000001C: ['EMR_SETMETARGN'],
    0x0000001D: ['EMR_EXCLUDECLIPRECT'],
    0x0000001E: ['EMR_INTERSECTCLIPRECT'],
    0x0000001F: ['EMR_SCALEVIEWPORTEXTEX'],
    0x00000020: ['EMR_SCALEWINDOWEXTEX'],
    0x00000021: ['EMR_SAVEDC', EmrSavedc],
    0x00000022: ['EMR_RESTOREDC', EmrRestoredc],
    0x00000023: ['EMR_SETWORLDTRANSFORM'],
    0x00000024: ['EMR_MODIFYWORLDTRANSFORM', EmrModifyworldtransform],
    0x00000025: ['EMR_SELECTOBJECT', EmrSelectobject],
    0x00000026: ['EMR_CREATEPEN'],
    0x00000027: ['EMR_CREATEBRUSHINDIRECT', EmrCreatebrushindirect],
    0x00000028: ['EMR_DELETEOBJECT', EmrDeleteobject],
    0x00000029: ['EMR_ANGLEARC'],
    0x0000002A: ['EMR_ELLIPSE'],
    0x0000002B: ['EMR_RECTANGLE'],
    0x0000002C: ['EMR_ROUNDRECT'],
    0x0000002D: ['EMR_ARC'],
    0x0000002E: ['EMR_CHORD'],
    0x0000002F: ['EMR_PIE'],
    0x00000030: ['EMR_SELECTPALETTE'],
    0x00000031: ['EMR_CREATEPALETTE'],
    0x00000032: ['EMR_SETPALETTEENTRIES'],
    0x00000033: ['EMR_RESIZEPALETTE'],
    0x00000034: ['EMR_REALIZEPALETTE'],
    0x00000035: ['EMR_EXTFLOODFILL'],
    0x00000036: ['EMR_LINETO', EmrLineto],
    0x00000037: ['EMR_ARCTO'],
    0x00000038: ['EMR_POLYDRAW'],
    0x00000039: ['EMR_SETARCDIRECTION'],
    0x0000003A: ['EMR_SETMITERLIMIT'],
    0x0000003B: ['EMR_BEGINPATH', EmrBeginpath],
    0x0000003C: ['EMR_ENDPATH', EmrEndpath],
    0x0000003D: ['EMR_CLOSEFIGURE', EmrClosefigure],
    0x0000003E: ['EMR_FILLPATH', EmrFillpath],
    0x0000003F: ['EMR_STROKEANDFILLPATH', EmrStrokeandfillpath],
    0x00000040: ['EMR_STROKEPATH'],
    0x00000041: ['EMR_FLATTENPATH'],
    0x00000042: ['EMR_WIDENPATH'],
    0x00000043: ['EMR_SELECTCLIPPATH', EmrSelectclippath],
    0x00000044: ['EMR_ABORTPATH'],
    0x00000046: ['EMR_COMMENT', EmrComment],
    0x00000047: ['EMR_FILLRGN'],
    0x00000048: ['EMR_FRAMERGN'],
    0x00000049: ['EMR_INVERTRGN'],
    0x0000004A: ['EMR_PAINTRGN'],
    0x0000004B: ['EMR_EXTSELECTCLIPRGN', EmrExtselectcliprgn],
    0x0000004C: ['EMR_BITBLT', EmrBitblt],
    0x0000004D: ['EMR_STRETCHBLT'],
    0x0000004E: ['EMR_MASKBLT'],
    0x0000004F: ['EMR_PLGBLT'],
    0x00000050: ['EMR_SETDIBITSTODEVICE'],
    0x00000051: ['EMR_STRETCHDIBITS', EmrStretchdibits],
    0x00000052: ['EMR_EXTCREATEFONTINDIRECTW'],
    0x00000053: ['EMR_EXTTEXTOUTA'],
    0x00000054: ['EMR_EXTTEXTOUTW'],
    0x00000055: ['EMR_POLYBEZIER16'],
    0x00000056: ['EMR_POLYGON16', EmrPolygon16],
    0x00000057: ['EMR_POLYLINE16'],
    0x00000058: ['EMR_POLYBEZIERTO16', EmrPolybezierto16],
    0x00000059: ['EMR_POLYLINETO16', EmrPolylineto16],
    0x0000005A: ['EMR_POLYPOLYLINE16'],
    0x0000005B: ['EMR_POLYPOLYGON16', EmrPolypolygon16],
    0x0000005C: ['EMR_POLYDRAW16'],
    0x0000005D: ['EMR_CREATEMONOBRUSH'],
    0x0000005E: ['EMR_CREATEDIBPATTERNBRUSHPT'],
    0x0000005F: ['EMR_EXTCREATEPEN', EmrExtcreatepen],
    0x00000060: ['EMR_POLYTEXTOUTA'],
    0x00000061: ['EMR_POLYTEXTOUTW'],
    0x00000062: ['EMR_SETICMMODE', EmrSeticmmode],
    0x00000063: ['EMR_CREATECOLORSPACE'],
    0x00000064: ['EMR_SETCOLORSPACE'],
    0x00000065: ['EMR_DELETECOLORSPACE'],
    0x00000066: ['EMR_GLSRECORD'],
    0x00000067: ['EMR_GLSBOUNDEDRECORD'],
    0x00000068: ['EMR_PIXELFORMAT'],
    0x00000069: ['EMR_DRAWESCAPE'],
    0x0000006A: ['EMR_EXTESCAPE'],
    0x0000006C: ['EMR_SMALLTEXTOUT'],
    0x0000006D: ['EMR_FORCEUFIMAPPING'],
    0x0000006E: ['EMR_NAMEDESCAPE'],
    0x0000006F: ['EMR_COLORCORRECTPALETTE'],
    0x00000070: ['EMR_SETICMPROFILEA'],
    0x00000071: ['EMR_SETICMPROFILEW'],
    0x00000072: ['EMR_ALPHABLEND'],
    0x00000073: ['EMR_SETLAYOUT'],
    0x00000074: ['EMR_TRANSPARENTBLT'],
    0x00000076: ['EMR_GRADIENTFILL'],
    0x00000077: ['EMR_SETLINKEDUFIS'],
    0x00000078: ['EMR_SETTEXTJUSTIFICATION'],
    0x00000079: ['EMR_COLORMATCHTOTARGETW'],
    0x0000007A: ['EMR_CREATECOLORSPACEW'],
    # EmfPlus
    0x4001: ['EmfPlusHeader'],
    0x4002: ['EmfPlusEndOfFile'],
    0x4003: ['EmfPlusComment'],
    0x4004: ['EmfPlusGetDC'],
    0x4005: ['EmfPlusMultiFormatStart'],
    0x4006: ['EmfPlusMultiFormatSection'],
    0x4007: ['EmfPlusMultiFormatEnd'],
    0x4008: ['EmfPlusObject'],
    0x4009: ['EmfPlusClear'],
    0x400A: ['EmfPlusFillRects'],
    0x400B: ['EmfPlusDrawRects'],
    0x400C: ['EmfPlusFillPolygon'],
    0x400D: ['EmfPlusDrawLines'],
    0x400E: ['EmfPlusFillEllipse'],
    0x400F: ['EmfPlusDrawEllipse'],
    0x4010: ['EmfPlusFillPie'],
    0x4011: ['EmfPlusDrawPie'],
    0x4012: ['EmfPlusDrawArc'],
    0x4013: ['EmfPlusFillRegion'],
    0x4014: ['EmfPlusFillPath'],
    0x4015: ['EmfPlusDrawPath'],
    0x4016: ['EmfPlusFillClosedCurve'],
    0x4017: ['EmfPlusDrawClosedCurve'],
    0x4018: ['EmfPlusDrawCurve'],
    0x4019: ['EmfPlusDrawBeziers'],
    0x401A: ['EmfPlusDrawImage'],
    0x401B: ['EmfPlusDrawImagePoints'],
    0x401C: ['EmfPlusDrawString'],
    0x401D: ['EmfPlusSetRenderingOrigin'],
    0x401E: ['EmfPlusSetAntiAliasMode'],
    0x401F: ['EmfPlusSetTextRenderingHint'],
    0x4020: ['EmfPlusSetTextContrast'],
    0x4021: ['EmfPlusSetInterpolationMode'],
    0x4022: ['EmfPlusSetPixelOffsetMode'],
    0x4023: ['EmfPlusSetCompositingMode'],
    0x4024: ['EmfPlusSetCompositingQuality'],
    0x4025: ['EmfPlusSave'],
    0x4026: ['EmfPlusRestore'],
    0x4027: ['EmfPlusBeginContainer'],
    0x4028: ['EmfPlusBeginContainerNoParams'],
    0x4029: ['EmfPlusEndContainer'],
    0x402A: ['EmfPlusSetWorldTransform'],
    0x402B: ['EmfPlusResetWorldTransform'],
    0x402C: ['EmfPlusMultiplyWorldTransform'],
    0x402D: ['EmfPlusTranslateWorldTransform'],
    0x402E: ['EmfPlusScaleWorldTransform'],
    0x402F: ['EmfPlusRotateWorldTransform'],
    0x4030: ['EmfPlusSetPageTransform'],
    0x4031: ['EmfPlusResetClip'],
    0x4032: ['EmfPlusSetClipRect'],
    0x4033: ['EmfPlusSetClipPath'],
    0x4034: ['EmfPlusSetClipRegion'],
    0x4035: ['EmfPlusOffsetClip'],
    0x4036: ['EmfPlusDrawDriverstring'],
    0x4037: ['EmfPlusStrokeFillPath'],
    0x4038: ['EmfPlusSerializableObject'],
    0x4039: ['EmfPlusSetTSGraphics'],
    0x403A: ['EmfPlusSetTSClip'],
}

# vim:set filetype=python shiftwidth=4 softtabstop=4 expandtab:
