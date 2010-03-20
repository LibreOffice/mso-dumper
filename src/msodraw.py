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

import globals

def indent (level):
    return '  '*level

class RecordHeader:
    def __init__ (self):
        self.recVer = None
        self.recInstance = None
        self.recType = None
        self.recLen = None


class ColorRef:
    def __init__ (self, byte):
        self.red   = (byte & 0x000000FF)
        self.green = (byte & 0x0000FF00) / 256 
        self.blue  = (byte & 0x00FF0000) / 65536
        self.flag  = (byte & 0xFF000000) / 16777216

        self.paletteIndex = (self.flag & 0x01) != 0
        self.paletteRGB   = (self.flag & 0x02) != 0
        self.systemRGB    = (self.flag & 0x04) != 0
        self.schemeIndex  = (self.flag & 0x08) != 0
        self.sysIndex     = (self.flag & 0x10) != 0

    def appendLine (self, recHdl, level):
        if self.paletteIndex:
            # red and green and used as an unsigned index into the current color palette.
            paletteId = self.green * 256 + self.red
            recHdl.appendLine(indent(level) + "color index in current palette: %d"%paletteId)
        if self.sysIndex:
            # red and green are used as an unsigned 16-bit index into the system color table.
            sysId = self.green * 256 + self.red
            recHdl.appendLine(indent(level) + "system index: %d"%sysId)
        elif self.schemeIndex:
            # the red value is used as as a color scheme index
            recHdl.appendLine(indent(level) + "color scheme index: %d"%self.red)

        else:
            recHdl.appendLine(indent(level) + "color: (red=%d, green=%d, blue=%d)    flag: 0x%2.2X"%
                (self.red, self.green, self.blue, self.flag))
            recHdl.appendLine(indent(level) + "palette index: %s"%recHdl.getTrueFalse(self.paletteIndex))
            recHdl.appendLine(indent(level) + "palette RGB: %s"%recHdl.getTrueFalse(self.paletteRGB))
            recHdl.appendLine(indent(level) + "system RGB: %s"%recHdl.getTrueFalse(self.systemRGB))
            recHdl.appendLine(indent(level) + "system RGB: %s"%recHdl.getTrueFalse(self.systemRGB))
            recHdl.appendLine(indent(level) + "scheme index: %s"%recHdl.getTrueFalse(self.schemeIndex))
            recHdl.appendLine(indent(level) + "system index: %s"%recHdl.getTrueFalse(self.sysIndex))



class FDG:
    def __init__ (self):
        self.shapeCount = None
        self.lastShapeID = -1

    def appendLines (self, recHdl, rh):
        recHdl.appendLine("FDG content (drawing data):")
        recHdl.appendLine("  ID of this shape: %d"%rh.recInstance)
        recHdl.appendLine("  shape count: %d"%self.shapeCount)
        recHdl.appendLine("  last shape ID: %d"%self.lastShapeID)


class FOPT:
    """property table for a shape instance"""

    class TextBoolean:

        def appendLines (self, recHdl, prop, level):
            indent = '  '*level
            A = (prop.value & 0x00000001) != 0
            B = (prop.value & 0x00000002) != 0
            C = (prop.value & 0x00000004) != 0
            D = (prop.value & 0x00000008) != 0
            E = (prop.value & 0x00000010) != 0
            F = (prop.value & 0x00010000) != 0
            G = (prop.value & 0x00020000) != 0
            H = (prop.value & 0x00040000) != 0
            I = (prop.value & 0x00080000) != 0
            J = (prop.value & 0x00100000) != 0
            recHdl.appendLineBoolean(indent + "fit shape to text",     B)
            recHdl.appendLineBoolean(indent + "auto text margin",      D)
            recHdl.appendLineBoolean(indent + "select text",           E)
            recHdl.appendLineBoolean(indent + "use fit shape to text", G)
            recHdl.appendLineBoolean(indent + "use auto text margin",  I)
            recHdl.appendLineBoolean(indent + "use select text",       J)

    class CXStyle:
        style = [
            'straight connector',     # 0x00000000
            'elbow-shaped connector', # 0x00000001
            'curved connector',       # 0x00000002
            'no connector'            # 0x00000003
        ]

        def appendLines (self, recHdl, prop, level):
            styleName = globals.getValueOrUnknown(FOPT.CXStyle.style, prop.value)
            recHdl.appendLine(indent(level) + "connector style: %s (0x%8.8X)"%(styleName, prop.value))

    class FillColor:

        def appendLines (self, recHdl, prop, level):
            color = ColorRef(prop.value)
            color.appendLine(recHdl, level)

    class FillStyle:

        def appendLines (self, recHdl, prop, level):
            flag1 = recHdl.readUnsignedInt(1)
            recHdl.moveForward(1)
            flag2 = recHdl.readUnsignedInt(1)
            recHdl.moveForward(1)
            A = (flag1 & 0x01) != 0 # fNoFillHitTest
            B = (flag1 & 0x02) != 0 # fillUseRect
            C = (flag1 & 0x04) != 0 # fillShape
            D = (flag1 & 0x08) != 0 # fHitTestFill
            E = (flag1 & 0x10) != 0 # fFilled
            F = (flag1 & 0x20) != 0 # fUseShapeAnchor
            G = (flag1 & 0x40) != 0 # fRecolorFillAsPicture

            H = (flag2 & 0x01) != 0 # fUseNoFillHitTest
            I = (flag2 & 0x02) != 0 # fUsefillUseRect
            J = (flag2 & 0x04) != 0 # fUsefillShape
            K = (flag2 & 0x08) != 0 # fUsefHitTestFill
            L = (flag2 & 0x10) != 0 # fUsefFilled
            M = (flag2 & 0x20) != 0 # fUsefUseShapeAnchor
            N = (flag2 & 0x40) != 0 # fUsefRecolorFillAsPicture

            recHdl.appendLine(indent(level)+"fNoFillHitTest            : %s"%recHdl.getTrueFalse(A))
            recHdl.appendLine(indent(level)+"fillUseRect               : %s"%recHdl.getTrueFalse(B))
            recHdl.appendLine(indent(level)+"fillShape                 : %s"%recHdl.getTrueFalse(C))
            recHdl.appendLine(indent(level)+"fHitTestFill              : %s"%recHdl.getTrueFalse(D))
            recHdl.appendLine(indent(level)+"fFilled                   : %s"%recHdl.getTrueFalse(E))
            recHdl.appendLine(indent(level)+"fUseShapeAnchor           : %s"%recHdl.getTrueFalse(F))
            recHdl.appendLine(indent(level)+"fRecolorFillAsPicture     : %s"%recHdl.getTrueFalse(G))

            recHdl.appendLine(indent(level)+"fUseNoFillHitTest         : %s"%recHdl.getTrueFalse(H))
            recHdl.appendLine(indent(level)+"fUsefillUseRect           : %s"%recHdl.getTrueFalse(I))
            recHdl.appendLine(indent(level)+"fUsefillShape             : %s"%recHdl.getTrueFalse(J))
            recHdl.appendLine(indent(level)+"fUsefHitTestFill          : %s"%recHdl.getTrueFalse(K))
            recHdl.appendLine(indent(level)+"fUsefFilled               : %s"%recHdl.getTrueFalse(L))
            recHdl.appendLine(indent(level)+"fUsefUseShapeAnchor       : %s"%recHdl.getTrueFalse(M))
            recHdl.appendLine(indent(level)+"fUsefRecolorFillAsPicture : %s"%recHdl.getTrueFalse(N))

    class LineColor:

        def appendLines (self, recHdl, prop, level):
            color = ColorRef(prop.value)
            color.appendLine(recHdl, level)

    class GroupShape:

        flagNames = [
            'fPrint',                 # A
            'fHidden',                # B
            'fOneD',                  # C
            'fIsButton',              # D
            'fOnDblClickNotify',      # E
            'fBehindDocument',        # F
            'fEditedWrap',            # G
            'fScriptAnchor',          # H
            'fReallyHidden',          # I
            'fAllowOverlap',          # J
            'fUserDrawn',             # K
            'fHorizRule',             # L
            'fNoshadeHR',             # M
            'fStandardHR',            # N
            'fIsBullet',              # O
            'fLayoutInCell',          # P
            'fUsefPrint',             # Q
            'fUsefHidden',            # R
            'fUsefOneD',              # S
            'fUsefIsButton',          # T
            'fUsefOnDblClickNotify',  # U
            'fUsefBehindDocument',    # V
            'fUsefEditedWrap',        # W
            'fUsefScriptAnchor',      # X
            'fUsefReallyHidden',      # Y
            'fUsefAllowOverlap',      # Z
            'fUsefUserDrawn',         # a
            'fUsefHorizRule',         # b
            'fUsefNoshadeHR',         # c
            'fUsefStandardHR',        # d
            'fUsefIsBullet',          # e
            'fUsefLayoutInCell'       # f
        ]

        def appendLines (self, recHdl, prop, level):
            flag = prop.value
            flagCount = len(FOPT.GroupShape.flagNames)
            recHdl.appendLine(indent(level)+"flag: 0x%8.8X"%flag)
            for i in xrange(0, flagCount):
                bval = (flag & 0x00000001)
                recHdl.appendLine(indent(level)+"%s: %s"%(FOPT.GroupShape.flagNames[i], recHdl.getTrueFalse(bval)))
                flag /= 2

    propTable = {
        0x00BF: ['Text Boolean Properties', TextBoolean],
        0x0181: ['Fill Color', FillColor],
        0x01BF: ['Fill Style Boolean Properties', FillStyle],
        0x01C0: ['Line Color', LineColor],
        0x0303: ['Connector Shape Style (cxstyle)', CXStyle],
        0x03BF: ['Group Shape Boolean Properties', GroupShape]
    }

    class E:
        """single property entry in a property table"""
        def __init__ (self):
            self.ID          = None
            self.flagBid     = False
            self.flagComplex = False
            self.value       = None
            self.extra       = None

    def __init__ (self):
        self.properties = []

    def appendLines (self, recHdl, rh):
        recHdl.appendLine("FOPT content (property table):")
        recHdl.appendLine("  property count: %d"%rh.recInstance)
        for i in xrange(0, rh.recInstance):
            recHdl.appendLine("    "+"-"*57)
            prop = self.properties[i]
            if FOPT.propTable.has_key(prop.ID):
                # We have a handler for this property.
                # propData is expected to have two elements: name (0) and handler (1).
                propHdl = FOPT.propTable[prop.ID]
                recHdl.appendLine("    property name: %s (0x%4.4X)"%(propHdl[0], prop.ID))
                propHdl[1]().appendLines(recHdl, prop, 2)
            else:
                recHdl.appendLine("    property ID: 0x%4.4X"%prop.ID)
                if prop.flagComplex:
                    recHdl.appendLine("    complex property: %s"%globals.getRawBytes(prop.extra, True, False))
                elif prop.flagBid:
                    recHdl.appendLine("    blip ID: %d"%prop.value)
                else:
                    # regular property value
                    recHdl.appendLine("    property value: 0x%8.8X"%prop.value)


class FRIT:
    def __init__ (self):
        self.lastGroupID = None
        self.secondLastGroupID = None

    def appendLines (self, recHdl, rh):
        pass


class FSP:
    def __init__ (self):
        self.spid = None
        self.flag = None

    def appendLines (self, recHdl, rh):
        recHdl.appendLine("FSP content (instance of a shape):")
        recHdl.appendLine("  ID of this shape: %d"%self.spid)
        groupShape     = (self.flag & 0x0001) != 0
        childShape     = (self.flag & 0x0002) != 0
        topMostInGroup = (self.flag & 0x0004) != 0
        deleted        = (self.flag & 0x0008) != 0
        oleObject      = (self.flag & 0x0010) != 0
        haveMaster     = (self.flag & 0x0020) != 0
        flipHorizontal = (self.flag & 0x0040) != 0
        flipVertical   = (self.flag & 0x0080) != 0
        isConnector    = (self.flag & 0x0100) != 0
        haveAnchor     = (self.flag & 0x0200) != 0
        background     = (self.flag & 0x0400) != 0
        haveProperties = (self.flag & 0x0800) != 0
        recHdl.appendLineBoolean("  group shape", groupShape)
        recHdl.appendLineBoolean("  child shape", childShape)
        recHdl.appendLineBoolean("  topmost in group", topMostInGroup)
        recHdl.appendLineBoolean("  deleted", deleted)
        recHdl.appendLineBoolean("  OLE object shape", oleObject)
        recHdl.appendLineBoolean("  have valid master", haveMaster)
        recHdl.appendLineBoolean("  horizontally flipped", flipHorizontal)
        recHdl.appendLineBoolean("  vertically flipped", flipVertical)
        recHdl.appendLineBoolean("  connector shape", isConnector)
        recHdl.appendLineBoolean("  have anchor", haveAnchor)
        recHdl.appendLineBoolean("  background shape", background)
        recHdl.appendLineBoolean("  have shape type property", haveProperties)


class FSPGR:
    def __init__ (self):
        self.left   = None
        self.top    = None
        self.right  = None
        self.bottom = None

    def appendLines (self, recHdl, rh):
        recHdl.appendLine("FSPGR content (coordinate system of group shape):")
        recHdl.appendLine("  left boundary: %d"%self.left)
        recHdl.appendLine("  top boundary: %d"%self.top)
        recHdl.appendLine("  right boundary: %d"%self.right)
        recHdl.appendLine("  bottom boundary: %d"%self.bottom)


class FConnectorRule:
    def __init__ (self):
        self.ruleID = None
        self.spIDA = None
        self.spIDB = None
        self.spIDC = None
        self.conSiteIDA = None
        self.conSiteIDB = None

    def appendLines (self, recHdl, rh):
        recHdl.appendLine("FConnectorRule content:")
        recHdl.appendLine("  rule ID: %d"%self.ruleID)
        recHdl.appendLine("  ID of the shape where the connector starts: %d"%self.spIDA)
        recHdl.appendLine("  ID of the shape where the connector ends: %d"%self.spIDB)
        recHdl.appendLine("  ID of the connector shape: %d"%self.spIDB)
        recHdl.appendLine("  ID of the connection site in the begin shape: %d"%self.conSiteIDA)
        recHdl.appendLine("  ID of the connection site in the end shape: %d"%self.conSiteIDB)


class Type:
    dgContainer     = 0xF002
    spgrContainer   = 0xF003
    spContainer     = 0xF004
    solverContainer = 0xF005
    FDG             = 0xF008
    FSPGR           = 0xF009
    FSP             = 0xF00A
    FOPT            = 0xF00B
    FConnectorRule  = 0xF012

containerTypeNames = {
    Type.dgContainer:      'OfficeArtDgContainer',
    Type.spContainer:      'OfficeArtSpContainer',
    Type.spgrContainer:    'OfficeArtSpgrContainer',
    Type.solverContainer:  'OfficeArtSolverContainer',
    Type.FDG:              'OfficeArtFDG',
    Type.FOPT:             'OfficeArtFOPT',
    Type.FSP:              'OfficeArtFSP',
    Type.FSPGR:            'OfficeArtFSPGR',
    Type.FConnectorRule:   'OfficeArtFConnectorRule'
}

# ----------------------------------------------------------------------------

class MSODrawHandler(globals.ByteStream):

    singleIndent = ' '*2

    def __init__ (self, bytes, parent):
        """The 'parent' instance must have appendLine() method that takes one string argument."""

        globals.ByteStream.__init__(self, bytes)
        self.parent = parent

    def printRecordHeader (self, rh, level=0):
        self.parent.appendLine(indent(level+1) + "record header:")
        self.parent.appendLine(indent(level+1) + "recVer: 0x%1.1X"%rh.recVer)
        self.parent.appendLine(indent(level+1) + "recInstance: 0x%3.3X"%rh.recInstance)
        self.parent.appendLine(indent(level+1) + "recType: 0x%4.4X (%s)"%(rh.recType, MSODrawHandler.getRecTypeName(rh)))
        self.parent.appendLine(indent(level+1) + "recLen: %d"%rh.recLen)

    def readFDG (self):
        fdg = FDG()
        fdg.shapeCount  = self.readUnsignedInt(4)
        fdg.lastShapeID = self.readUnsignedInt(4)
        return fdg

    def readFOPT (self, rh):
        fopt = FOPT()
        strm = globals.ByteStream(self.readBytes(rh.recLen))
        while not strm.isEndOfRecord():
            entry = FOPT.E()
            val = strm.readUnsignedInt(2)
            entry.ID          = (val & 0x3FFF)
            entry.flagBid     = (val & 0x4000) # if true, the value is a blip ID.
            entry.flagComplex = (val & 0x8000) # if true, the value stores the size of the extra bytes.
            entry.value = strm.readSignedInt(4)
            if entry.flagComplex:
                entry.extra = strm.readBytes(entry.value)
            fopt.properties.append(entry)

        return fopt

    def readFRIT (self):
        frit = FRIT()
        frit.lastGroupID = self.readUnsignedInt(2)
        frit.secondLastGroupID = self.readUnsignedInt(2)
        return frit

    def readFSP (self):
        fsp = FSP()
        fsp.spid = self.readUnsignedInt(4)
        fsp.flag = self.readUnsignedInt(4)
        return fsp

    def readFSPGR (self):
        fspgr = FSPGR()
        fspgr.left   = self.readSignedInt(4)
        fspgr.top    = self.readSignedInt(4)
        fspgr.right  = self.readSignedInt(4)
        fspgr.bottom = self.readSignedInt(4)
        return fspgr

    def readFConnectorRule (self):
        fcon = FConnectorRule()
        fcon.ruleID = self.readUnsignedInt(4)
        fcon.spIDA = self.readUnsignedInt(4)
        fcon.spIDB = self.readUnsignedInt(4)
        fcon.spIDC = self.readUnsignedInt(4)
        fcon.conSiteIDA = self.readUnsignedInt(4)
        fcon.conSiteIDB = self.readUnsignedInt(4)
        return fcon

    def readRecordHeader (self):
        rh = RecordHeader()
        mixed = self.readUnsignedInt(2)
        rh.recVer = (mixed & 0x000F)
        rh.recInstance = (mixed & 0xFFF0) / 16
        rh.recType = self.readUnsignedInt(2)
        rh.recLen  = self.readUnsignedInt(4)
        return rh

    @staticmethod
    def getRecTypeName (rh):
        if containerTypeNames.has_key(rh.recType):
            return containerTypeNames[rh.recType]
        return 'unknown'

    def parseBytes (self):
        firstRec = True
        while not self.isEndOfRecord():
            if firstRec:
                firstRec = False
            else:
                self.parent.appendLine("-"*61)
            rh = self.readRecordHeader()
            self.printRecordHeader(rh)
            # if rh.recType == Type.dgContainer:
            if rh.recVer == 0xF:
                # container
                pass
            elif rh.recType == Type.FDG:
                fdg = self.readFDG()
                fdg.appendLines(self.parent, rh)
            elif rh.recType == Type.FOPT:
                fopt = self.readFOPT(rh)
                fopt.appendLines(self.parent, rh)
            elif rh.recType == Type.FSPGR:
                fspgr = self.readFSPGR()
                fspgr.appendLines(self.parent, rh)
            elif rh.recType == Type.FSP:
                fspgr = self.readFSP()
                fspgr.appendLines(self.parent, rh)
            elif rh.recType == Type.FConnectorRule:
                fcon = self.readFConnectorRule()
                fcon.appendLines(self.parent, rh)
            else:
                # unknown object
                self.readBytes(rh.recLen)
