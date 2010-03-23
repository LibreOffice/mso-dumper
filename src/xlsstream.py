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

import sys
import ole, globals, xlsrecord
from globals import output

class EndOfStream(Exception): pass

    # opcode: [canonical name, description, handler (optional)]

recData = {
    0x0006: ["FORMULA", "Cell Formula", xlsrecord.Formula],
    0x000A: ["EOF", "End of File"],
    0x000C: ["CALCCOUNT", "Iteration Count"],
    0x000D: ["CALCMODE", "Calculation Mode"],
    0x000E: ["PRECISION", "Precision"],
    0x000F: ["REFMODE", "Reference Mode"],
    0x0010: ["DELTA", "Iteration Increment"],
    0x0011: ["ITERATION", "Iteration Mode"],
    0x0012: ["PROTECT", "Protection Flag"],
    0x0013: ["PASSWORD", "Protection Password"],
    0x0014: ["HEADER", "Print Header on Each Page"],
    0x0015: ["FOOTER", "Print Footer on Each Page"],
    0x0016: ["EXTERNCOUNT", "Number of External References"],
    0x0017: ["EXTERNSHEET", "External Reference", xlsrecord.ExternSheet],
    0x0018: ["NAME", "Internal Defined Name", xlsrecord.Name],
    0x0019: ["WINDOWPROTECT", "Windows Are Protected"],
    0x0021: ["ARRAY", "Array-Entered Formula", xlsrecord.Array], # undocumented, but identical to 0x0221 ?
    0x001A: ["VERTICALPAGEBREAKS", "Explicit Column Page Breaks"],
    0x001B: ["HORIZONTALPAGEBREAKS", "Explicit Row Page Breaks"],
    0x001C: ["NOTE", "Comment Associated with a Cell"],
    0x001D: ["SELECTION", "Current Selection"],
    0x0022: ["DATEMODE", "Base Date for Displaying Date Values"],
    0x0023: ["EXTERNNAME", "Externally Defined Name", xlsrecord.ExternName],
    0x0026: ["LEFTMARGIN", "Left Margin Measurement"],
    0x0027: ["RIGHTMARGIN", "Right Margin Measurement"],
    0x0028: ["TOPMARGIN", "Top Margin Measurement"],
    0x0029: ["BOTTOMMARGIN", "Bottom Margin Measurement"],
    0x002A: ["PRINTHEADERS", "Print Row/Column Labels"],
    0x002B: ["PRINTGRIDLINES", "Print Gridlines Flag"],
    0x002F: ["FILEPASS", "File Is Password-Protected", xlsrecord.FilePass],
    0x0031: ["FONT", "Font and Character Formatting", xlsrecord.Font],
    0x003C: ["CONTINUE", "Continues Long Records"],
    0x003D: ["WINDOW1", "Window Information"],
    0x0040: ["BACKUP", "Save Backup Version of the File"],
    0x0041: ["PANE", "Number of Panes and Their Position"],
    0x0042: ["CODEPAGE/CODENAME", "Default Code Page/VBE Object Name"],
    0x004D: ["PLS", "Environment-Specific Print Record"],
    0x0050: ["DCON", "Data Consolidation Information"],
    0x0051: ["DCONREF", "Data Consolidation References"],
    0x0052: ["DCONNAME", "Data Consolidation Named References"],
    0x0055: ["DEFCOLWIDTH", "Default Width for Columns", xlsrecord.DefColWidth],
    0x0059: ["XCT", "CRN Record Count", xlsrecord.Xct],
    0x005A: ["CRN", "Nonresident Operands", xlsrecord.Crn],
    0x005B: ["FILESHARING", "File-Sharing Information"],
    0x005C: ["WRITEACCESS", "Write Access User Name"],
    0x005D: ["OBJ", "Describes a Graphic Object", xlsrecord.Obj],
    0x005E: ["UNCALCED", "Recalculation Status"],
    0x005F: ["SAVERECALC", "Recalculate Before Save"],
    0x0060: ["TEMPLATE", "Workbook Is a Template"],
    0x0063: ["OBJPROTECT", "Objects Are Protected"],
    0x007D: ["COLINFO", "Column Formatting Information", xlsrecord.ColInfo],
    0x007F: ["IMDATA", "Image Data"],
    0x0080: ["GUTS", "Size of Row and Column Gutters"],
    0x0081: ["WSBOOL", "Additional Workspace Information"],
    0x0082: ["GRIDSET", "State Change of Gridlines Option"],
    0x0083: ["HCENTER", "Center Between Horizontal Margins"],
    0x0084: ["VCENTER", "Center Between Vertical Margins"],
    0x0085: ["BOUNDSHEET", "Sheet Information", xlsrecord.BoundSheet],
    0x0086: ["WRITEPROT", "Workbook Is Write-Protected"],
    0x0087: ["ADDIN", "Workbook Is an Add-in Macro"],
    0x0088: ["EDG", "Edition Globals"],
    0x0089: ["PUB", "Publisher"],
    0x008C: ["COUNTRY", "Default Country and WIN.INI Country"],
    0x008D: ["HIDEOBJ", "Object Display Options"],
    0x0090: ["SORT", "Sorting Options"],
    0x0091: ["SUB", "Subscriber"],
    0x0092: ["PALETTE", "Color Palette Definition"],
    0x0094: ["LHRECORD", ".WK? File Conversion Information"],
    0x0095: ["LHNGRAPH", "Named Graph Information"],
    0x0096: ["SOUND", "Sound Note"],
    0x0098: ["LPR", "Sheet Was Printed Using LINE.PRINT()"],
    0x0099: ["STANDARDWIDTH", "Standard Column Width"],
    0x009A: ["FNGROUPNAME", "Function Group Name"],
    0x009B: ["FILTERMODE", "Sheet Contains Filtered List", xlsrecord.FilterMode],
    0x009C: ["FNGROUPCOUNT", "Built-in Function Group Count"],
    0x009D: ["AUTOFILTERINFO", "Drop-Down Arrow Count", xlsrecord.AutofilterInfo],
    0x009E: ["AUTOFILTER", "AutoFilter Data", xlsrecord.Autofilter],
    0x00A0: ["SCL", "Window Zoom Magnification"],
    0x00A1: ["SETUP", "Page Setup"],
    0x00A9: ["COORDLIST", "Polygon Object Vertex Coordinates"],
    0x00AB: ["GCW", "Global Column-Width Flags"],
    0x00AE: ["SCENMAN", "Scenario Output Data"],
    0x00AF: ["SCENARIO", "Scenario Data"],
    0x00B0: ["SXVIEW", "View Definition", xlsrecord.SXView],
    0x00B1: ["SXVD", "View Fields", xlsrecord.SXViewFields],
    0x00B2: ["SXVI", "View Item", xlsrecord.SXViewItem],
    0x00B4: ["SXIVD", "Row/Column Field IDs"],
    0x00B5: ["SXLI", "Line Item Array"],
    0x00B6: ["SXPI", "Page Item"],
    0x00B8: ["DOCROUTE", "Routing Slip Information"],
    0x00B9: ["RECIPNAME", "Recipient Name"],
    0x00BC: ["SHRFMLA", "Shared Formula"],
    0x00BD: ["MULRK", "Multiple RK Cells", xlsrecord.MulRK],
    0x00BE: ["MULBLANK", "Multiple Blank Cells"],
    0x00C1: ["MMS", "ADDMENU/DELMENU Record Group Count"],
    0x00C2: ["ADDMENU", "Menu Addition"],
    0x00C3: ["DELMENU", "Menu Deletion"],
    0x00C5: ["SXDI", "Data Item", xlsrecord.SXDataItem],
    0x00C6: ["SXDB", "PivotTable Cache Data", xlsrecord.SXDb],
    0x00C7: ["SXFIELD", "Pivot Field", xlsrecord.SXField],
    0x00C8: ["SXINDEXLIST", "Indices to Source Data"],       
    0x00C9: ["SXDOUBLE", "Double Value", xlsrecord.SXDouble],                    
    0x00CA: ["SXBOOLEAN", "Boolean Value", xlsrecord.SXBoolean],                  
    0x00CB: ["SXERROR", "Error Code", xlsrecord.SXError],                       
    0x00CC: ["SXINTEGER", "Integer Value", xlsrecord.SXInteger],                  
    0x00CD: ["SXSTRING", "String", xlsrecord.SXString],
    0x00CE: ["SXDATETIME", "Date & Time Special Format"],    
    0x00CF: ["SXEMPTY", "Empty Value"],                      
    0x00D0: ["SXTBL", "Multiple Consolidation Source Info"],
    0x00D1: ["SXTBRGIITM", "Page Item Name Count"],
    0x00D2: ["SXTBPG", "Page Item Indexes"],
    0x00D3: ["OBPROJ", "Visual Basic Project"],
    0x00D5: ["SXIDSTM", "Stream ID", xlsrecord.SXStreamID],
    0x00D6: ["RSTRING", "Cell with Character Formatting"],
    0x00D7: ["DBCELL", "Stream Offsets", xlsrecord.DBCell],
    0x00DA: ["BOOKBOOL", "Workbook Option Flag"],
    0x00DC: ["PARAMQRY", "Query Parameters"],
    0x00DC: ["SXEXT", "External Source Information"],
    0x00DD: ["SCENPROTECT", "Scenario Protection"],
    0x00DE: ["OLESIZE", "Size of OLE Object"],
    0x00DF: ["UDDESC", "Description String for Chart Autoformat"],
    0x00E0: ["XF", "Extended Format", xlsrecord.XF],
    0x00E1: ["INTERFACEHDR", "Beginning of User Interface Records"],
    0x00E2: ["INTERFACEEND", "End of User Interface Records"],
    0x00E3: ["SXVS", "View Source", xlsrecord.SXViewSource],
    0x00EA: ["TABIDCONF", "Sheet Tab ID of Conflict History"],
    0x00EB: ["MSODRAWINGGROUP", "Microsoft Office Drawing Group", xlsrecord.MSODrawingGroup],
    0x00EC: ["MSODRAWING", "Microsoft Office Drawing", xlsrecord.MSODrawing],
    0x00ED: ["MSODRAWINGSELECTION", "Microsoft Office Drawing Selection"],
    0x00EF: ["PHONETIC", "Asian Phonetic Settings", xlsrecord.PhoneticInfo],
    0x00F0: ["SXRULE", "PivotTable Rule Data"],
    0x00F1: ["SXEX", "PivotTable View Extended Information"],
    0x00F2: ["SXFILT", "PivotTable Rule Filter"],
    0x00F6: ["SXNAME", "PivotTable Name"],
    0x00F7: ["SXSELECT", "PivotTable Selection Information"],
    0x00F8: ["SXPAIR", "PivotTable Name Pair"],
    0x00F9: ["SXFMLA", "PivotTable Parsed Expression"],
    0x00FB: ["SXFORMAT", "PivotTable Format Record"],
    0x00FC: ["SST", "Shared String Table", xlsrecord.SST],
    0x00FD: ["LABELSST", "Cell Value", xlsrecord.LabelSST],
    0x00FF: ["EXTSST", "Extended Shared String Table"],
    0x0100: ["SXVDEX", "Extended PivotTable View Fields", xlsrecord.SXViewFieldsEx],
    0x0103: ["SXFORMULA", "PivotTable Formula Record"],
    0x0122: ["SXDBEX", "PivotTable Cache Data", xlsrecord.SXDbEx],
    0x013D: ["TABID", "Sheet Tab Index Array"],
    0x0160: ["USESELFS", "Natural Language Formulas Flag"],
    0x0161: ["DSF", "Double Stream File"],
    0x01A9: ["USERBVIEW", "Workbook Custom View Settings"],
    0x01AA: ["USERSVIEWBEGIN", "Custom View Settings"],
    0x01AB: ["USERSVIEWEND", "End of Custom View Records"],
    0x01AD: ["QSI", "External Data Range"],
    0x01AE: ["SUPBOOK", "Supporting Workbook", xlsrecord.SupBook],
    0x01AF: ["PROT4REV", "Shared Workbook Protection Flag"],
    0x01B0: ["CONDFMT", "Conditional Formatting Range Information"],
    0x01B1: ["CF", "Conditional Formatting Conditions"],
    0x01B2: ["DVAL", "Data Validation Information"],
    0x01B5: ["DCONBIN", "Data Consolidation Information"],
    0x01B6: ["TXO", "Text Object"],
    0x01B7: ["REFRESHALL", "Refresh Flag", xlsrecord.RefreshAll],
    0x01B8: ["HLINK", "Hyperlink", xlsrecord.Hyperlink],
    0x01BB: ["SXFDBTYPE", "SQL Datatype Identifier"],
    0x01BC: ["PROT4REVPASS", "Shared Workbook Protection Password"],
    0x01BE: ["DV", "Data Validation Criteria"],
    0x01C0: ["EXCEL9FILE", "Excel 9 File"],
    0X01C1: ["RECALCID", "Recalc Information"],
    0x0200: ["DIMENSIONS", "Cell Table Size", xlsrecord.Dimensions],
    0x0201: ["BLANK", "Blank Cell", xlsrecord.Blank],
    0x0203: ["NUMBER", "Floating-Point Cell Value", xlsrecord.Number],
    0x0204: ["LABEL", "Cell Value", xlsrecord.Label],
    0x0205: ["BOOLERR", "Cell Value"],
    0x0207: ["STRING", "String Value of a Formula", xlsrecord.String],
    0x0208: ["ROW", "Describes a Row", xlsrecord.Row],
    0x020B: ["INDEX", "Index Record"],
    0x0218: ["NAME", "Defined Name"],
    0x0221: ["ARRAY", "Array-Entered Formula", xlsrecord.Array],
    0x0223: ["EXTERNNAME", "Externally Referenced Name"],
    0x0225: ["DEFAULTROWHEIGHT", "Default Row Height", xlsrecord.DefRowHeight],
    0x0231: ["FONT", "Font Description", xlsrecord.Font],
    0x0236: ["TABLE", "Data Table"],
    0x023E: ["WINDOW2", "Sheet Window Information"],
    0x027E: ["RK", "Cell with Encoded Integer or Floating-Point", xlsrecord.RK],
    0x0293: ["STYLE", "Style Information"],
    0x041E: ["FORMAT", "Number Format"],
    0x0802: ["QSISXTAG", "Pivot Table and Query Table Extensions", xlsrecord.PivotQueryTableEx],
    0x0809: ["BOF", "Beginning of File", xlsrecord.BOF],
    0x0810: ["SXVIEWEX9", "Pivot Table Extensions", xlsrecord.SXViewEx9],
    0x0858: ["CHPIVOTREF", "Pivot Chart Reference"],
    0x0862: ["SHEETLAYOUT", "Tab Color below Sheet Name"],
    0x0863: ["BOOKEXT", "Extra Book Info"],
    0x0864: ["SXADDL", "Pivot Table Additional Info", xlsrecord.SXAddlInfo],
    0x0867: ["FEATHEADR", "Shared Feature Header", xlsrecord.FeatureHeader],
    0x0868: ["RANGEPROTECTION", "Protected Range on Protected Sheet"],
    0x087D: ["XFEXT", "XF Extension"],
    0x088C: ["COMPAT12", "Compatibility Checker 12"],
    0x088E: ["TABLESTYLES", "Table Styles"],
    0x0892: ["STYLEEXT", "Named Cell Style Extension"],
    0x0896: ["THEME", "Theme"],
    0x089A: ["MTRSETTINGS", "Multi-threaded Calculation Settings"],
    0x089C: ["HEADERFOOTER", "Header Footer"],
    0x089B: ["COMPRESSPICTURES", "Automatic Picture Compression Mode"],
    0x08A3: ["FORCEFULLCALCULATION", "Force Full Calculation Mode"],
    0x1001: ["CHUNITS", "?"],
    0x1002: ["CHCHART", "Chart Header Data", xlsrecord.CHChart],
    0x1003: ["CHSERIES", "Chart Series", xlsrecord.CHSeries],
    0x1006: ["CHDATAFORMAT", "?"],
    0x1007: ["CHLINEFORMAT", "Line or Border Formatting of A Chart"],
    0x1009: ["CHMARKERFORMAT", "?"],
    0x100D: ["CHSTRING", "Series Category Name or Title Text in Chart"],
    0x100A: ["CHAREAFORMAT", "Area Formatting Attribute of A Chart"],
    0x100B: ["CHPIEFORMAT", "?"],
    0x100C: ["CHATTACHEDLABEL", "?"],
    0x100D: ["CHSTRING", "?"],
    0x1014: ["CHTYPEGROUP", "?"],
    0x1015: ["CHLEGEND", "?", xlsrecord.CHLegend],
    0x1017: ["CHBAR, CHCOLUMN", "?", xlsrecord.CHBar],
    0x1018: ["CHLINE", "?", xlsrecord.CHLine],
    0x1019: ["CHPIE", "?"],
    0x101A: ["CHAREA", "?"],
    0x101B: ["CHSCATTER", "?"],
    0x001C: ["CHCHARTLINE", "?"],
    0x101D: ["CHAXIS", "Chart Axis", xlsrecord.CHAxis],
    0x101E: ["CHTICK", "?"],
    0x101F: ["CHVALUERANGE", "Chart Axis Value Range", xlsrecord.CHValueRange],
    0x1020: ["CHLABELRANGE", "Chart Axis Label Range", xlsrecord.CHLabelRange],
    0x1021: ["CHAXISLINE", "?"],
    0x1024: ["CHDEFAULTTEXT", "?"],
    0x1025: ["CHTEXT", "?"],
    0x1026: ["CHFONT", "?"],
    0x1027: ["CHOBJECTLINK", "?"],
    0x1032: ["CHFRAME", "Border and Area Formatting of A Chart"],
    0x1033: ["CHBEGIN", "Start of Chart Record Block"],
    0x1034: ["CHEND", "End of Chart Record Block"],
    0x1035: ["CHPLOTFRAME", "Chart Plot Frame"],
    0x103A: ["CHCHART3D", "?"],
    0x103C: ["CHPICFORMAT", "?"],
    0x103D: ["CHDROPBAR", "?"],
    0x103E: ["CHRADARLINE", "?"],
    0x103F: ["CHSURFACE", "?"],
    0x1040: ["CHRADARAREA", "?"],
    0x1041: ["CHAXESSET", "?"],
    0x1044: ["CHPROPERTIES", "?", xlsrecord.CHProperties],
    0x1045: ["CHSERGROUP", "?"],
    0x1048: ["CHPIVOTREF", "?"],
    0x104A: ["CHSERPARENT", "?"],
    0x104B: ["CHSERTRENDLINE", "?"],
    0x104E: ["CHFORMAT", "?"],
    0x104F: ["CHFRAMEPOS", "?"],
    0x1050: ["CHFORMATRUNS", "?"],
    0x1051: ["CHSOURCELINK", "Data Source of A Chart", xlsrecord.CHSourceLink],
    0x105B: ["CHSERERRORBAR", "?"],
    0x105D: ["CHSERIESFORMAT", "?"],
    0x105F: ["CH3DDATAFORMAT", "?"],
    0x1061: ["CHPIEEXT", "?"],
    0x1062: ["CHLABELRANGE2", "?"],
    0x1065: ["CHSIINDEX*", "?"],
    0x1066: ["CHESCHERFORMAT", "?"]
}

recDataRev = {
    0x0137: ["INSERT*", "Change Track Insert"],
    0x0138: ["INFO*", "Change Track Info"],
    0x013B: ["CELLCONTENT*", "Change Track Cell Content", xlsrecord.CTCellContent],
    0x013D: ["SHEETID*", "Change Track Sheet Identifier"],
    0x0140: ["MOVERANGE*", "Change Track Move Range"],
    0x014D: ["INSERTSHEET*", "Change Track Insert Sheet"],
    0x014E: ["BONB*", "Change Track Beginning of Nested Block"],
    0x0150: ["BONB*", "Change Track Beginning of Nested Block"],
    0x014F: ["EONB*", "Change Track End of Nested Block"],
    0x0151: ["EONB*", "Change Track End of Nested Block"]
}

class XLStream(object):

    def __init__ (self, chars, params, strmData):
        self.chars = chars
        self.size = len(self.chars)
        self.pos = 0
        self.version = None

        self.header = None
        self.MSAT = None
        self.SAT = None

        self.params = params
        self.strmData = strmData

    def __printSep (self, c='-', w=68, prefix=''):
        print(prefix + c*w)

    def printStreamInfo (self):
        self.__printSep('=', 68)
        print("Excel File Format Dumper by Kohei Yoshida")
        print("  total stream size: %d bytes"%self.size)
        self.__printSep('=', 68)
        print('')

    def printHeader (self):
        self.__parseHeader()
        self.header.output()
        self.MSAT = self.header.getMSAT()

    def printMSAT (self):
        self.MSAT.output()

    def printSAT (self):
        sat = self.MSAT.getSAT()
        sat.output()

    def printSSAT (self):
        obj = self.header.getSSAT()
        if obj == None:
            return
        obj.output()

    def __parseHeader (self):
        if self.header == None:
            self.header = ole.Header(self.chars, self.params)
            self.pos = self.header.parse()

    def __getDirectoryObj (self):
        self.__parseHeader()
        obj = self.header.getDirectory()
        if obj == None:
            return None
        obj.parseDirEntries()
        return obj

    def printDirectory (self):
        obj = self.__getDirectoryObj()
        if obj == None:
            return
        obj.output()


    def getDirectoryNames (self):
        obj = self.__getDirectoryObj()
        if obj == None:
            return
        return obj.getDirectoryNames()


    def getDirectoryStreamByName (self, name):
        obj = self.__getDirectoryObj()
        bytes = []
        if obj != None:
            bytes = obj.getRawStreamByName(name)
        strm = XLDirStream(bytes, self.params, self.strmData)
        return strm

class DirType:
    Workbook = 0
    RevisionLog = 1
    PivotTableCache = 2

class XLDirStream(object):

    def __init__ (self, bytes, params, strmData):
        self.bytes = bytes
        self.size = len(self.bytes)
        self.pos = 0
        self.type = DirType.Workbook

        self.params = params
        self.strmData = strmData


    def readRaw (self, size=1):
        # assume little endian
        bytes = 0
        for i in xrange(0, size):
            b = ord(self.bytes[self.pos])
            if i == 0:
                bytes = b
            else:
                bytes += b*(256**i)
            self.pos += 1

        return bytes

    def readByteArray (self, size=1):
        bytes = []
        for i in xrange(0, size):
            if self.pos >= self.size:
                raise EndOfStream
            bytes.append(ord(self.bytes[self.pos]))
            self.pos += 1
        return bytes

    def __printSep (self, c='-', w=68, prefix=''):
        print(prefix + c*w)

    def __readRecordBytes (self):
        if self.size - self.pos < 4:
            raise EndOfStream

        pos = self.pos
        header = self.readRaw(2)
        if header == 0x0000:
            raise EndOfStream
        size = self.readRaw(2)
        bytes = self.readByteArray(size)
        return pos, header, size, bytes

    def __getRecordHandler (self, header, size, bytes):
        # record handler that parses the raw bytes and displays more 
        # meaningful information.
        handler = None 
        if recData.has_key(header) and len(recData[header]) >= 3:
            handler = recData[header][2](header, size, bytes, self.strmData)

        if handler != None and self.strmData.encrypted:
            # record handler exists.  Parse the record and display more info 
            # unless the stream is encrypted.
            handler = None

        return handler

    def __postReadRecord (self, header):
        if recData.has_key(header) and recData[header][0] == "FILEPASS":
            # presence of FILEPASS record indicates that the stream is 
            # encrypted.
            self.strmData.encrypted = True

    def fillModel (self, model):
        pos, header, size, bytes = self.__readRecordBytes()
        handler = self.__getRecordHandler(header, size, bytes)
        if handler != None:
            handler.fillModel(model)
        self.__postReadRecord(header)


    def readRecordXML (self):
        pos, header, size, bytes = self.__readRecordBytes()
        handler = self.__getRecordHandler(header, size, bytes)
        print (recData[header][1])
        self.__postReadRecord(header)
        return header

    def readRecord (self):
        pos, header, size, bytes = self.__readRecordBytes()

        # record handler that parses the raw bytes and displays more 
        # meaningful information.
        handler = None 
        
        print("")
        self.__printSep('=', 61, "%4.4Xh: "%header)
        if recData.has_key(header):
            print("%4.4Xh: %s - %s (%4.4Xh)"%
                  (header, recData[header][0], recData[header][1], header))
            if len(recData[header]) >= 3:
                handler = recData[header][2](header, size, bytes, self.strmData)
        elif self.type == DirType.RevisionLog and recDataRev.has_key(header):
            print("%4.4Xh: %s - %s (%4.4Xh)"%
                  (header, recDataRev[header][0], recDataRev[header][1], header))
            if len(recDataRev[header]) >= 3:
                handler = recDataRev[header][2](header, size, bytes, self.strmData)
        else:
            print("%4.4Xh: [unknown record name] (%4.4Xh)"%(header, header))

        if self.params.showStreamPos:
            print("%4.4Xh:   size = %d; pos = %d"%(header, size, pos))
        else:
            print("%4.4Xh:   size = %d"%(header, size))

        self.__printSep('-', 61, "%4.4Xh: "%header)
        for i in xrange(0, size):
            if (i+1) % 16 == 1:
                output("%4.4Xh: "%header)
            output("%2.2X "%bytes[i])
            if (i+1) % 16 == 0 and i != size-1:
                print("")
        if size > 0:
            print("")

        if handler != None and not self.strmData.encrypted:
            # record handler exists.  Parse the record and display more info 
            # unless the stream is encrypted.
            handler.output()

        self.__postReadRecord(header)
        return header
