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

class RecordHeader:
    def __init__ (self):
        self.recVer = None
        self.recInstance = None
        self.recType = None
        self.recLen = None


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
            indent = '  '*level
            styleName = getValueOrUnknown(FOPT.CXStyle.style, prop.value)
            recHdl.appendLine(indent + "connector style: %s (0x%8.8X)"%(styleName, prop.value))

    propTable = {
        0x00BF: ['Text Boolean Properties', TextBoolean],
        0x0303: ['Connector Shape Style (cxstyle)', CXStyle]
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
