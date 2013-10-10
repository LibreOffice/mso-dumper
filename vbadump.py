#!/usr/bin/env python2

import sys, os.path, optparse, math
sys.path.append(sys.path[0]+"/src")
import ole, globals, node, olestream, vbahelper

# alot of records follow the id, sizeofrecord pattern
# many of the dir records though seem to be id, sizeofstring,
# stringbuffer, reserved, sizeofstring(unicode), stringbuffer(unicode)
class DefaultBuffReservedBuffReader:
    def __init__ (self, reader ):
        self.reader = reader
    def parse(self):
        # pos before header
        print ("  skipping")
        # buffer
        size = self.reader.readUnsignedInt( 4 )
        self.reader.readBytes(size)        
        # reserved
        self.reader.readBytes(2)        
        # buffer
        size = self.reader.readUnsignedInt( 4 )
        self.reader.readBytes(size)    

class StdReader:
    def __init__ (self, reader ):
        self.reader = reader

class ProjectVersionReader:
    def __init__ (self, reader ):
        self.reader = reader
    def parse(self):
        # reserved
        self.reader.readBytes(4)
        # major
        major = self.reader.readUnsignedInt( 4 )
        minor = self.reader.readUnsignedInt( 2 )
        print("  major: 0x%x"%major)
        print("  minor: 0x%x"%minor)

class CodePageReader(StdReader):
    def parse(self):
        size = self.reader.readUnsignedInt( 4 )
        self.reader.codepage = self.reader.readUnsignedInt( size )

        #need a map of codepage code (int) to codepage alias(s) from
        # #FIXME codepage is hardcoded below
        self.reader.codepageName = "cp1252"
        print("  codepage: %i"%self.reader.codepage)
     
class ProjectNameReader(StdReader):
    def parse(self):
        size = self.reader.readUnsignedInt( 4 )
        bytes = self.reader.readBytes( size )
        print("  ProjectName: %s"%bytes.decode(self.reader.codepageName))
                 
class DocStringRecordReader(StdReader):
    def parse(self):
        size = self.reader.readUnsignedInt( 4 )
        bytes = self.reader.readBytes( size )
        print("  DocString: %s size[0x%x]"%(bytes.decode(self.reader.codepageName),size))
        #reserved
        self.reader.readBytes( 2 )
        #unicode docstring
        size = self.reader.readUnsignedInt( 4 )
        bytes = self.reader.readBytes( size )
        print("  DocString(utf-16): %s size[0x%x]"%(bytes.decode("utf-16").decode(self.reader.codepageName), size ))

class ConstantsRecordReader(StdReader):
    def parse(self):
        size = self.reader.readUnsignedInt( 4 )
        bytes = self.reader.readBytes( size )
        print("  Constants: size[0x%x]"%size)
        if size:
            print("%s"%bytes.decode(self.reader.codepageName))
        #reserved
        self.reader.readBytes( 2 )
        #unicode docstring
        size = self.reader.readUnsignedInt( 4 )
        bytes = self.reader.readBytes( size )
        print("  Constants(Utf-16): size[0x%x]"%size)
        if size:
            print("%s"%bytes.decode("uft-16").decode(self.reader.codepageName))

class ProjectHelpFilePathReader(StdReader):
    def parse(self):
        size = self.reader.readUnsignedInt( 4 )
        bytes = self.reader.readBytes( size )
        print("  HelpFile1: %s size[0x%x]"%(bytes.decode(self.reader.codepageName),size))
        #reserved
        self.reader.readBytes( 2 )
        #unicode docstring
        size = self.reader.readUnsignedInt( 4 )
        bytes = self.reader.readBytes( size )
        print("  HelpFile2: %s size[0x%x]"%(bytes.decode(self.reader.codepageName), size ))

class ProjectHelpFileContextReader(StdReader):
    def parse(self):
        size = self.reader.readUnsignedInt( 4 )
        context = self.reader.readUnsignedInt( size )
        print("  HelpContext: 0x%x"%context)

class ProjectModulesReader(StdReader):
    def parse(self):
        # start of module specific data
        size = self.reader.readUnsignedInt( 4 )
        count = self.reader.readUnsignedInt( size )
        self.reader.CurrentModule = ModuleInfo()
        print("  Num Modules: %i"%count) 

class ModuleNameReader(StdReader):
    def parse(self):
        size = self.reader.readUnsignedInt( 4 )
        nameBytes = self.reader.readBytes( size )
        moduleInfo = self.reader.CurrentModule
        moduleInfo.name = nameBytes.decode( self.reader.codepageName )
        print("  ModuleName: %s"%moduleInfo.name)

class ModuleStreamNameReader(StdReader):
    def parse(self):
        size = self.reader.readUnsignedInt( 4 )
        nameBytes = self.reader.readBytes( size )
        moduleInfo = self.reader.CurrentModule
        moduleInfo.streamname = nameBytes.decode( self.reader.codepageName )
        print("  ModuleStreamName: %s"%moduleInfo.name)
        #reserved
        self.reader.readBytes( 2 )
        size = self.reader.readUnsignedInt( 4 )
        nameUnicodeBytes = self.reader.readBytes( size )
        nameUnicode = nameUnicodeBytes.decode("utf-16").decode( self.reader.codepageName)
        print("  ModuleStreamName(utf-16): %s"%nameUnicode)

class ModuleOffSetReader(StdReader):
    def parse(self):
        size = self.reader.readUnsignedInt( 4 )
        moduleInfo = self.reader.CurrentModule
        moduleInfo.offset = self.reader.readUnsignedInt( size )
     
        print("  Offset: 0x%x"%moduleInfo.offset) 

class ProjectModuleTermReader(StdReader):
    def parse(self):
        size = self.reader.readUnsignedInt( 4 )
        # size must be zero ( assert? )
        moduleInfo = self.reader.CurrentModule
        self.reader.CurrentModule = ModuleInfo()
        # add current module to list
        self.reader.Modules.append( moduleInfo )

class ModuleTypeProceduralReader(StdReader):
    def parse(self):
        size = self.reader.readUnsignedInt( 4 )
        # size must be zero ( assert? )
        print("  Module Type: procedure")

class ModuleTypeOtherReader(StdReader):
    def parse(self):
        size = self.reader.readUnsignedInt( 4 )
        # size must be zero ( assert? )
        print("  Module Type: document, class or design")

class SysKindReader(StdReader):
    def parse(self):
        size = self.reader.readUnsignedInt( 4 )
        val = self.reader.readUnsignedInt( size )
        sysKind = "Unknown" 
        if val == 0:
           sysKind = "16 bit windows" 
        elif val == 1:
           sysKind = "32 bit windows" 
        elif val == 2:
           sysKind = "Macintosh" 
        print("  SysType: %s"%sysKind)

class LcidReader(StdReader):
    def parse(self):
        size = self.reader.readUnsignedInt( 4 )
        val = self.reader.readUnsignedInt( size )
        print("  LCID: 0x%x ( expected 0x409 )"%val)
   
class LcidInvokeReader(StdReader):
    def parse(self):
        size = self.reader.readUnsignedInt( 4 )
        val = self.reader.readUnsignedInt( size )
        print("  LCIDINVOKE: 0x%x ( expected 0x409 )"%val)
   

class LibFlagReader(StdReader):
    def parse(self):
        size = self.reader.readUnsignedInt( 4 )
        val = self.reader.readUnsignedInt( size )
        print("  LIBFLAGS: 0x%x"%val)

# map of record id to array containing description of records and optional
# a handler ( inspired by xlsstream.py )
dirRecordData = {
    #dir stream contains........
    #PROJECTINFORMATION RECORD
    #  which contains any of the following sub records
    0x0001: ["PROJECTSYSKIND", "SysKindRecord",SysKindReader],
    0x0002: ["PROJECTLCID", "LcidRecord",LcidReader],
    0x0003: ["PROJECTCODEPAGE", "CodePageRecord", CodePageReader ],
    0x0004: ["PROJECTNAME", "NameRecord", ProjectNameReader],
    0x0005: ["PROJECTDOCSTRING", "DocStringRecord", DocStringRecordReader ],
    0x0006: ["PROJECTHELPFILEPATH", "HelpFilePathRecord", ProjectHelpFilePathReader],
    0x0007: ["PROJECTHELPCONTEXT", "HelpContextRecord", ProjectHelpFileContextReader],
    0x0008: ["PROJECTLIBFLAGS", "LibFlagsRecord",LibFlagReader],
    0x0009: ["PROJECTVERSION", "VersionRecord",ProjectVersionReader],
    0x0010: ["DIRTERMINATOR", "DirTerminator"],
    0x000C: ["PROJECTCONSTANTS", "ConstantsRecord", ConstantsRecordReader],
    0x0014: ["PROJECTLCIDINVOKE", "LcidInvokeRecord",LcidInvokeReader],
    #PROJECTREFERENCES
    # which contains any of the following sub records
    0x0016: ["REFERENCENAME", "NameRecord", DefaultBuffReservedBuffReader ],
    0x000D: ["REFERENCEREGISTERED", "ReferenceRegistered"],
    0x000E: ["REFERENCEPROJECT", "ReferenceProject"],
    0x002F: ["REFERENCECONTROL", "ReferenceControl"],
    #the following "FAKE-#FIXME record is not really a record but actually
    #is a reserved word ( with fixed value 0x0030 ) in the middle of a
    #REFEREBCECONTROL record
    0x0030: ["FAKE-#FIXME", "Fake record"],
    0x0033: ["REFERENCEORIGINAL", "ReferenceOriginal"],
    #
    0x000F: ["PROJECTMODULES", "ModulesRecord", ProjectModulesReader],
    0x0013: ["PROJECTCOOKIE", "CookieRecord"],
    0x002B: ["PROJECTMODULETERM", "ModuleTerminator", ProjectModuleTermReader],
    0x0019: ["MODULENAME", "ModuleName",ModuleNameReader],
    0x0047: ["MODULENAMEUNICODE", "ModuleNameUnicode"],
    0x001A: ["MODULESTREAMNAME", "ModuleStreamName", ModuleStreamNameReader],
    0x001C: ["MODULEDOCSTRING", "ModuleDocString", DefaultBuffReservedBuffReader],
    0x0031: ["MODULEOFFSET", "ModuleOffSet", ModuleOffSetReader],
    0x001E: ["MODULEHELPCONTEXT", "ModuleHelpContext"],
    0x002C: ["MODULECOOKIE", "ModuleCookie"],
    0x0021: ["MODULETYPE", "ModuleTypeProcedural", ModuleTypeProceduralReader],
    0x0022: ["MODULETYPE", "ModuleTypeDocClassOrDesgn", ModuleTypeOtherReader],
    0x0025: ["MODULEREADONLY", "ModuleReadOnly"],
    0x0028: ["MODULEPRIVATE", "ModulePrivate"],
}    


class ModuleInfo:
    def __init__(self):
        self.name = ""
        self.offset = 0
        self.streamname = "" 
        
class DirStreamReader( globals.ByteStream ): 
    def __init__ (self, bytes ):
        globals.ByteStream.__init__(self, bytes)
        self.Modules = []
        self.CodePage = None
        self.CurrentModule = None

    def parse(self):
        print("============ Dir Stream (inflated) size: 0x%x bytes ============"%len(self.bytes))
        print("")
        print("Offset  [recId]  RecordName")
        print("------- -------  -----------")
        while self.isEndOfRecord() == False:
            pos = self.getCurrentPos()
            recordID = self.readUnsignedInt( 2 )
            name = "Unknown"
            if dirRecordData.has_key( recordID ):
                name = dirRecordData[ recordID ][0]
            # if we have a handler let it deal with the record
            labelWidth = int(math.ceil(math.log(len(self.bytes), 10)))
            fmt = "0x%%%d.%dx: "%(labelWidth, labelWidth)
            sys.stdout.write(fmt%pos)
#            print ("%s [0x%x] "%(name,recordID))
            print '[0x{0:0>4x}] {1}'.format(recordID,name)
            if ( dirRecordData.has_key( recordID ) and len( dirRecordData[ recordID ] ) > 2 ):
                reader = dirRecordData[ recordID ][2]( self )
                reader.parse()
            else:
                print ("  skipping")
                size = self.readUnsignedInt( 4 )
                if size:
                    self.readBytes(size)        

class VBAContainer:
    def __init__( self, root ):
        # we'll take a storage DirNode
        # and try and find the VBA container
        # relative to that. That way we should
        # be able to cater for the normal 'word' or
        # 'excel' compound documents or indeed any arbitrary
        # storage that contains a 'VBA' sub-folder
        self.oleRoot = root
        self.vbaRoot = None

    def __findNodeByHierarchicalName( self, node, name ):
        if ( node.getHierarchicalName() == name ):
            return node
        else:
            for child in node.getChildren():
                result = self.__findNodeByHierarchicalName( child, name )
                if result != None:
                    return result
        return None

    def __findNodeContainingLeafName( self, parentNode, name ):
        for child in parentNode.getChildren():
            if name == child.Entry.Name:
                return parentNode
            node = self.__findNodeContainingLeafName( child, name )
            if node != None:
                return node
        return None

    def findVBARoot(self):
        # find the node that containes 'VBA' storage
        return self.__findNodeContainingLeafName( self.oleRoot, "VBA" )

    def dump( self ):
        self.vbaRoot = self.findVBARoot()
        if self.vbaRoot == None:
            print("Can't find VBA subcontainer")
            exit(1)
        # need to read the dir stream
        dirName = self.vbaRoot.getHierarchicalName() + "VBA/dir"
        dirNode = self.__findNodeByHierarchicalName( self.vbaRoot, dirName )
        if dirNode != None:
            #decompress 
            bytes = dirNode.getStream()
            compressed = vbahelper.CompressedVBAStream( bytes, 0 )
            bytes = compressed.decompress()
            reader = DirStreamReader( bytes )
            reader.parse()

            # dump the PROJECTxxx streams ( need to codepage from dir )
            for child in self.vbaRoot.getChildren():
                # first level children are PROJECT, PROJECTwm & PROJECTlk
                if child.isStorage() == False:
                    bytes = child.getStream()
                    print("")
                    print("============ %s Stream size: 0x%x bytes)============"%(child.getName(), len(bytes)))
                    print("")
                    if child.getName() == "PROJECT":
                        #straight text file
                        print("%s"%bytes.decode(reader.codepageName))
                    else:
                        globals.dumpBytes( bytes, 512) 
            for module in reader.Modules:
                fullStreamName = self.vbaRoot.getHierarchicalName() + "VBA/" + module.streamname
                moduleNode = self.__findNodeByHierarchicalName( self.vbaRoot, fullStreamName )
                bytes = moduleNode.getStream()
                print("============ %s Stream (inflated) size: 0x%x bytes offset: 0x%x ============"%(module.streamname,len(bytes), module.offset) )
                compressed = vbahelper.CompressedVBAStream( bytes, module.offset )
                bytes = compressed.decompress()
                source = bytes.decode( reader.codepageName )
                print("")
                print(source)
                print("")


def main():
    parser = optparse.OptionParser()

    if ( len ( sys.argv ) <= 1 ):
        print("usage: vbadump: file")
        sys.exit(1) 
    # prepare for supporting more options
    options, args = parser.parse_args()

    params = globals.Params()    

    container = ole.OleContainer( args[ 0 ], params )

    container.read()
    root = container.getRoot()
    vba = VBAContainer( root ) 
    vba.dump()

    exit(0)

if __name__ == '__main__':
    main()
