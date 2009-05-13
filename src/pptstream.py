########################################################################
#
#    OpenOffice.org - a multi-platform office productivity suite
#
#    Author:
#      Kohei Yoshida  <kyoshida@novell.com>
#      Thorsten Behrens <tbehrens@novell.com>
#
#   The Contents of this file are made available subject to the terms
#   of GNU Lesser General Public License Version 2.1 and any later
#   version.
#
########################################################################

import sys
import ole, globals, pptrecord
from globals import output

class EndOfStream(Exception): pass

# IDs from OOo's svx/inc/svx/msdffdef.hxx (slightly adapted)
#
# opcode: [canonical name, handler (optional)]

recData = {

    0:  ["DFF_PST_Unknown"],
    1:  ["DFF_PST_SubContainerCompleted"],
    2:  ["DFF_PST_IRRAtom"],
    3:  ["DFF_PST_PSS"],
    4:  ["DFF_PST_SubContainerException"],
    6:  ["DFF_PST_ClientSignal1"],
    7:  ["DFF_PST_ClientSignal2"],
   10:  ["DFF_PST_PowerPointStateInfoAtom"],
 1000:  ["DFF_PST_Document"],
 1001:  ["DFF_PST_DocumentAtom", pptrecord.DocAtom],
 1002:  ["DFF_PST_EndDocument"],
 1003:  ["DFF_PST_SlidePersist"],
 1004:  ["DFF_PST_SlideBase"],
 1005:  ["DFF_PST_SlideBaseAtom"],
 1006:  ["DFF_PST_Slide"],
 1007:  ["DFF_PST_SlideAtom"],
 1008:  ["DFF_PST_Notes"],
 1009:  ["DFF_PST_NotesAtom"],
 1010:  ["DFF_PST_Environment"],
 1011:  ["DFF_PST_SlidePersistAtom"],
 1012:  ["DFF_PST_Scheme"],
 1013:  ["DFF_PST_SchemeAtom"],
 1014:  ["DFF_PST_DocViewInfo"],
 1015:  ["DFF_PST_SslideLayoutAtom"],
 1016:  ["DFF_PST_MainMaster"],
 1017:  ["DFF_PST_SSSlideInfoAtom"],
 1018:  ["DFF_PST_SlideViewInfo"],
 1019:  ["DFF_PST_GuideAtom"],
 1020:  ["DFF_PST_ViewInfo"],
 1021:  ["DFF_PST_ViewInfoAtom"],
 1022:  ["DFF_PST_SlideViewInfoAtom"],
 1023:  ["DFF_PST_VBAInfo"],
 1024:  ["DFF_PST_VBAInfoAtom"],
 1025:  ["DFF_PST_SSDocInfoAtom"],
 1026:  ["DFF_PST_Summary"],
 1027:  ["DFF_PST_Texture"],
 1028:  ["DFF_PST_VBASlideInfo"],
 1029:  ["DFF_PST_VBASlideInfoAtom"],
 1030:  ["DFF_PST_DocRoutingSlip"],
 1031:  ["DFF_PST_OutlineViewInfo"],
 1032:  ["DFF_PST_SorterViewInfo"],
 1033:  ["DFF_PST_ExObjList"],
 1034:  ["DFF_PST_ExObjListAtom"],
 1035:  ["DFF_PST_PPDrawingGroup"],
 1036:  ["DFF_PST_PPDrawing"],
 1038:  ["DFF_PST_Theme", pptrecord.ZipRecord],
 1039:  ["DFF_PST_ColorMapping"],
 1040:  ["DFF_PST_NamedShows"],
 1041:  ["DFF_PST_NamedShow"],
 1042:  ["DFF_PST_NamedShowSlides"],
 1052:  ["DFF_PST_OriginalMainMasterId"],
 1054:  ["DFF_PST_RoundTripContentMasterInfo", pptrecord.ZipRecord],
 1055:  ["DFF_PST_RoundTripShapeId"],
 1058:  ["DFF_PST_RoundTripContentMasterId"],
 1059:  ["DFF_PST_RoundTripOArtTextStyles", pptrecord.ZipRecord],
 1064:  ["DFF_PST_RoundTripCustomTableStyles"],
 2000:  ["DFF_PST_List"],
 2005:  ["DFF_PST_FontCollection"],
 2017:  ["DFF_PST_ListPlaceholder"],
 2019:  ["DFF_PST_BookmarkCollection"],
 2020:  ["DFF_PST_SoundCollection"],
 2021:  ["DFF_PST_SoundCollAtom"],
 2022:  ["DFF_PST_Sound"],
 2023:  ["DFF_PST_SoundData"],
 2025:  ["DFF_PST_BookmarkSeedAtom"],
 2026:  ["DFF_PST_GuideList"],
 2028:  ["DFF_PST_RunArray"],
 2029:  ["DFF_PST_RunArrayAtom"],
 2030:  ["DFF_PST_ArrayElementAtom"],
 2031:  ["DFF_PST_Int4ArrayAtom"],
 2032:  ["DFF_PST_ColorSchemeAtom", pptrecord.ColorScheme],
 3008:  ["DFF_PST_OEShape"],
 3009:  ["DFF_PST_ExObjRefAtom"],
 3011:  ["DFF_PST_OEPlaceholderAtom"],
 3020:  ["DFF_PST_GrColor"],
 3025:  ["DFF_PST_GrectAtom"],
 3031:  ["DFF_PST_GratioAtom"],
 3032:  ["DFF_PST_Gscaling"],
 3034:  ["DFF_PST_GpointAtom"],
 3035:  ["DFF_PST_OEShapeAtom"],
 3998:  ["DFF_PST_OutlineTextRefAtom"],
 3999:  ["DFF_PST_TextHeaderAtom", pptrecord.TextHeader],
 4000:  ["DFF_PST_TextCharsAtom", pptrecord.ShapeUniString],
 4001:  ["DFF_PST_StyleTextPropAtom", pptrecord.TextStyles],
 4002:  ["DFF_PST_BaseTextPropAtom", pptrecord.TextStyles],
 4003:  ["DFF_PST_TxMasterStyleAtom", pptrecord.MasterTextStyles],
 4004:  ["DFF_PST_TxCFStyleAtom"],
 4005:  ["DFF_PST_TxPFStyleAtom"],
 4006:  ["DFF_PST_TextRulerAtom"],
 4007:  ["DFF_PST_TextBookmarkAtom"],
 4008:  ["DFF_PST_TextBytesAtom", pptrecord.ShapeString],
 4009:  ["DFF_PST_TxSIStyleAtom"],
 4010:  ["DFF_PST_TextSpecInfoAtom"],
 4011:  ["DFF_PST_DefaultRulerAtom"],
 4023:  ["DFF_PST_FontEntityAtom"],
 4024:  ["DFF_PST_FontEmbedData"],
 4025:  ["DFF_PST_TypeFace"],
 4026:  ["DFF_PST_CString", pptrecord.UniString],
 4027:  ["DFF_PST_ExternalObject"],
 4033:  ["DFF_PST_MetaFile"],
 4034:  ["DFF_PST_ExOleObj"],
 4035:  ["DFF_PST_ExOleObjAtom"],
 4036:  ["DFF_PST_ExPlainLinkAtom"],
 4037:  ["DFF_PST_CorePict"],
 4038:  ["DFF_PST_CorePictAtom"],
 4039:  ["DFF_PST_ExPlainAtom"],
 4040:  ["DFF_PST_SrKinsoku"],
 4041:  ["DFF_PST_Handout"],
 4044:  ["DFF_PST_ExEmbed"],
 4045:  ["DFF_PST_ExEmbedAtom"],
 4046:  ["DFF_PST_ExLink"],
 4047:  ["DFF_PST_ExLinkAtom_old"],
 4048:  ["DFF_PST_BookmarkEntityAtom"],
 4049:  ["DFF_PST_ExLinkAtom"],
 4050:  ["DFF_PST_SrKinsokuAtom"],
 4051:  ["DFF_PST_ExHyperlinkAtom"],
 4053:  ["DFF_PST_ExPlain"],
 4054:  ["DFF_PST_ExPlainLink"],
 4055:  ["DFF_PST_ExHyperlink"],
 4056:  ["DFF_PST_SlideNumberMCAtom"],
 4057:  ["DFF_PST_HeadersFooters"],
 4058:  ["DFF_PST_HeadersFootersAtom"],
 4062:  ["DFF_PST_RecolorEntryAtom"],
 4063:  ["DFF_PST_TxInteractiveInfoAtom"],
 4065:  ["DFF_PST_EmFormatAtom"],
 4066:  ["DFF_PST_CharFormatAtom"],
 4067:  ["DFF_PST_ParaFormatAtom"],
 4068:  ["DFF_PST_MasterText"],
 4071:  ["DFF_PST_RecolorInfoAtom"],
 4073:  ["DFF_PST_ExQuickTime"],
 4074:  ["DFF_PST_ExQuickTimeMovie"],
 4075:  ["DFF_PST_ExQuickTimeMovieData"],
 4076:  ["DFF_PST_ExSubscription"],
 4077:  ["DFF_PST_ExSubscriptionSection"],
 4078:  ["DFF_PST_ExControl"],
 4091:  ["DFF_PST_ExControlAtom"],
 4080:  ["DFF_PST_SlideListWithText"],
 4081:  ["DFF_PST_AnimationInfoAtom"],
 4082:  ["DFF_PST_InteractiveInfo"],
 4083:  ["DFF_PST_InteractiveInfoAtom"],
 4084:  ["DFF_PST_SlideList"],
 4085:  ["DFF_PST_UserEditAtom"],
 4086:  ["DFF_PST_CurrentUserAtom"],
 4087:  ["DFF_PST_DateTimeMCAtom"],
 4088:  ["DFF_PST_GenericDateMCAtom"],
 4089:  ["DFF_PST_HeaderMCAtom"],
 4090:  ["DFF_PST_FooterMCAtom"],
 4100:  ["DFF_PST_ExMediaAtom"],
 4101:  ["DFF_PST_ExVideo"],
 4102:  ["DFF_PST_ExAviMovie"],
 4103:  ["DFF_PST_ExMCIMovie"],
 4109:  ["DFF_PST_ExMIDIAudio"],
 4110:  ["DFF_PST_ExCDAudio"],
 4111:  ["DFF_PST_ExWAVAudioEmbedded"],
 4112:  ["DFF_PST_ExWAVAudioLink"],
 4113:  ["DFF_PST_ExOleObjStg"],
 4114:  ["DFF_PST_ExCDAudioAtom"],
 4115:  ["DFF_PST_ExWAVAudioEmbeddedAtom"],
 4116:  ["DFF_PST_AnimationInfo"],
 4117:  ["DFF_PST_RTFDateTimeMCAtom"],
 5000:  ["DFF_PST_ProgTags"],
 5001:  ["DFF_PST_ProgStringTag"],
 5002:  ["DFF_PST_ProgBinaryTag"],
 5003:  ["DFF_PST_BinaryTagData"],
 6000:  ["DFF_PST_PrintOptions"],
 6001:  ["DFF_PST_PersistPtrFullBlock"],
 6002:  ["DFF_PST_PersistPtrIncrementalBlock"],
10000:  ["DFF_PST_RulerIndentAtom"],
10001:  ["DFF_PST_GscalingAtom"],
10002:  ["DFF_PST_GrColorAtom"],
10003:  ["DFF_PST_GLPointAtom"],
10004:  ["DFF_PST_GlineAtom"],

0xF000: ["DFF_msofbtDggContainer"],
0xF006: ["DFF_msofbtDgg"],
0xF016: ["DFF_msofbtCLSID"],
0xF00B: ["DFF_msofbtOPT", pptrecord.Property],
0xF11A: ["DFF_msofbtColorMRU"],
0xF11E: ["DFF_msofbtSplitMenuColors"],
0xF001: ["DFF_msofbtBstoreContainer"],
0xF007: ["DFF_msofbtBSE"],
0xF018: ["DFF_msofbtBlipFirst"],
0xF117: ["DFF_msofbtBlipLast"],

0xF002: ["DFF_msofbtDgContainer"],
0xF008: ["DFF_msofbtDg"],
0xF118: ["DFF_msofbtRegroupItems"],
0xF120: ["DFF_msofbtColorScheme"],
0xF003: ["DFF_msofbtSpgrContainer"],
0xF004: ["DFF_msofbtSpContainer"],
0xF009: ["DFF_msofbtSpgr", pptrecord.Rect],
0xF00A: ["DFF_msofbtSp", pptrecord.Shape],
0xF00C: ["DFF_msofbtTextbox"],
0xF00D: ["DFF_msofbtClientTextbox"],
0xF00E: ["DFF_msofbtAnchor"],
0xF00F: ["DFF_msofbtChildAnchor", pptrecord.Rect],
0xF010: ["DFF_msofbtClientAnchor", pptrecord.Rect],
0xF011: ["DFF_msofbtClientData"],
0xF11F: ["DFF_msofbtOleObject"],
0xF11D: ["DFF_msofbtDeletedPspl"],
0xF122: ["DFF_msofbtUDefProp", pptrecord.Property],
0xF005: ["DFF_msofbtSolverContainer"],
0xF012: ["DFF_msofbtConnectorRule"],
0xF013: ["DFF_msofbtAlignRule"],
0xF014: ["DFF_msofbtArcRule"],
0xF015: ["DFF_msofbtClientRule"],
0xF017: ["DFF_msofbtCalloutRule"],

0xF119: ["DFF_msofbtSelection"]

}

class PPTFile(object):
    """Represents the whole powerpoint file - feed will all bytes."""
    def __init__ (self, chars, params):
        self.chars = chars
        self.size = len(self.chars)
        self.version = None
        self.params = params

        self.header = ole.Header(self.chars, self.params)
        self.pos = self.header.parse()


    def __printSep (self, c='-', w=68, prefix=''):
        print(prefix + c*w)


    def printStreamInfo (self):
        self.__printSep('=', 68)
        print("PPT File Format Dumper by Kohei Yoshida & Thorsten Behrens")
        print("  total stream size: %d bytes"%self.size)
        self.__printSep('=', 68)
        print('')


    def printHeader (self):
        self.header.output()


    def __getDirectoryObj (self):
        obj = self.header.getDirectory()
        if obj is None:
            return None
        obj.parseDirEntries()
        return obj


    def printDirectory (self):
        obj = self.__getDirectoryObj()
        if obj is None:
            return
        obj.output()


    def getDirectoryNames (self):
        obj = self.__getDirectoryObj()
        if obj is None:
            return
        return obj.getDirectoryNames()


    def getDirectoryStreamByName (self, name):
        obj = self.__getDirectoryObj()
        bytes = []
        if obj is not None:
            bytes = obj.getRawStreamByName(name)
        strm = PPTDirStream(bytes, self.params)
        return strm


class PPTDirStream(object):
    """Represents one single powerpoint file subdirectory, like e.g. \"PowerPoint Document\"."""
    def __init__ (self, bytes, params, prefix='', recordInfo=None):
        self.bytes = bytes
        self.size = len(self.bytes)
        self.pos = 0
        self.prefix = prefix
        self.params = params
        self.properties = {"recordInfo": recordInfo}


    def readBytes (self, size=1):
        if self.size - self.pos < size:
            raise EndOfStream
        bytes = self.bytes[self.pos:self.pos+size]
        self.pos += size
        return bytes


    def readUnsignedInt (self, size=1):
        return globals.getUnsignedInt(self.readBytes(size))


    def __print (self, text):
        print(self.prefix + text)


    def __printSep (self, c='-', w=68, prefix=''):
        self.__print(prefix + c*w)


    def readRecords (self):
        try:
            # read until data is exhausted (min record size: 8 bytes)
            while self.pos+8 < self.size:
                print("")
                self.readRecord()
            return True
        except EndOfStream:
            return False


    def printRecordHeader (self, startPos, recordInstance, recordVersion, recordType, size):
        self.__printSep('=')
        if recordType in recData:
            self.__print("[%s]"%recData[recordType][0])
        else:
            self.__print("[anon record]")
        self.__print("(type: %4.4Xh inst: %4.4Xh, vers: %4.4Xh, start: %d, size: %d)"%
              (recordType, recordInstance, recordVersion, startPos, size))
        self.__printSep('=')


    def printRecordDump (self, bytes, recordType):
        size = len(bytes)
        self.__printSep('-', 61, "%4.4Xh: "%recordType)
        for i in xrange(0, size):
            if (i+1) % 16 == 1:
                output(self.prefix + "%4.4Xh: "%recordType)
            output("%2.2X "%ord(bytes[i]))
            if (i+1) % 16 == 0 and i != size-1:
                print("")
        if size > 0:
            print("")
            self.__printSep('-', 61, "%4.4Xh: "%recordType)


    def readRecord (self):
        startPos = self.pos
        recordInstance = self.readUnsignedInt(2)
        recordVersion = (recordInstance & 0x000F)
        recordInstance = recordInstance // 16
        recordType = self.readUnsignedInt(2)
        size = self.readUnsignedInt(4)

        self.printRecordHeader(startPos, recordInstance, recordVersion, recordType, size)
        bytes = self.readBytes(size)

        recordInfo = None
        if recordType in recData and len(recData[recordType]) >= 2:
            recordInfo = recData[recordType]

        if recordVersion == 0x0F:
            # substream? recurse into that
            subSubStrm = PPTDirStream(bytes, self.params, self.prefix+" ", recordInfo)
            subSubStrm.readRecords()
        elif recordInfo is not None:
            handler = recordInfo[1](recordType, recordInstance, size, bytes, self.properties, self.prefix)
            print("")
            # call special record handler, if any
            if handler is not None:
                handler.output()
            self.printRecordDump(bytes, recordType)
        elif size > 0:
            print("")
            self.printRecordDump(bytes, recordType)
