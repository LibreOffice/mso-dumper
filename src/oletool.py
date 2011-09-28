#!/usr/bin/env python
########################################################################
#
#  Copyright (c) 2011 Noel Power
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

import sys, os.path, optparse, struct

sys.path.append(sys.path[0]+"/src")

import ole, globals

class DateTime:
    def __init__(self):
        self.day = 0
        self.month = 0 
        self.year = 0
        self.hour = 0
        self.second = 0

class DirNode:

    def __init__(self, entry, index):
        self.Nodes = []
        self.Entry = entry;
        self.HierachicalName = ''
        self.Index = index

    def isStorage():
        return entry.Type == Directory.Type.RootStorage

class OleContainer:

    def __init__(self,filePath, params ):
        self.filePath = filePath
        self.params = params
        self.chars = ''
        self.file = open(self.filePath, 'rb')
        self.chars = self.file.read()
        self.file.close()
        self.header = ole.Header(self.chars, self.params)
        self.pos = self.header.parse()
        self.outputBytes = None;

    def output (self):
        
        # #FIXME common initialisation code
        self.header.output() 

        obj =  self.header.getDirectory()
        if obj != None:
            obj.parseDirEntries()
            obj.output()

    def __getModifiedTime(self, entry):
        # need parse/decode Entry.TimeModified
        # ( although the documentation indicates that it might not be
        # worth it 'cause they are not universally used
        modified  = DateTime
        modified.day = 0
        modified.month = 0 
        modified.year = 0
        modified.hour = 0
        modified.second = 0
        return modified

    def __addSiblings( self, entries, parent, child ):
        # add left siblings
        nextLeft = child.Entry.DirIDLeft
        if ( nextLeft > 0 ):
            newEntry = DirNode( entries[ nextLeft ], nextLeft )
            newEntry.HierachicalName = parent.HierachicalName + globals.encodeName( newEntry.Entry.Name )
            if  newEntry.Entry.DirIDRoot > 0:
                newEntry.HierachicalName = newEntry.HierachicalName + '/'

            self.__addSiblings( entries, parent, newEntry ) 
            parent.Nodes.insert( 0, newEntry )

        nextRight = child.Entry.DirIDRight
        # add children to the right 
        if ( nextRight > 0 ):
            newEntry = DirNode( entries[ nextRight ], nextRight )
            newEntry.HierachicalName = parent.HierachicalName + globals.encodeName( newEntry.Entry.Name )
            if  newEntry.Entry.DirIDRoot > 0:
                newEntry.HierachicalName = newEntry.HierachicalName + '/'
            self.__addSiblings( entries, parent, newEntry ) 
            parent.Nodes.append( newEntry )

    def __buildTreeImpl(self, entries, parent ):

        if ( parent.Entry.DirIDRoot > 0 ):
            newEntry = DirNode( entries[ parent.Entry.DirIDRoot ], parent.Entry.DirIDRoot )
            newEntry.HierachicalName = parent.HierachicalName + globals.encodeName( newEntry.Entry.Name )
            if ( newEntry.Entry.DirIDRoot > 0 ):
                newEntry.HierachicalName =  newEntry.HierachicalName + '/'

            self.__addSiblings( entries, parent, newEntry )
            parent.Nodes.append( newEntry )
            
        for child in parent.Nodes:
            if child.Entry.DirIDRoot > 0:
                self.__buildTreeImpl( entries, child )

    def __buildTree(self, entries ):
        treeRoot = DirNode( entries[0], 0 ) 
        self.__buildTreeImpl( entries, treeRoot )
        return treeRoot

    def __findNodeByHierachicalName( self, node, name ):
        if node.HierachicalName == name:
            return node
        else:
            for child in node.Nodes:
                result = self.__findNodeByHierachicalName( child, name )
                if result != None:
                    return result 
        return None 

    def __printListReport( self, treeNode, stats ):

        dateInfo = self.__getModifiedTime( treeNode.Entry )

        if len( treeNode.HierachicalName ) > 0 :
            print '{0:8d}  {1:0<2d}-{2:0<2d}-{3:0<2d} {4:0<2d}:{5:0<2d}   {6}'.format(treeNode.Entry.StreamSize, dateInfo.day, dateInfo.month, dateInfo.year, dateInfo.hour, dateInfo.second, treeNode.HierachicalName )
     
        for node in treeNode.Nodes:
            # ignore the root
            self.__printListReport( node, stats )

    def __printHeader(self):
        print ("OLE: %s")%self.filePath
        print (" Length     Date   Time    Name")
        print ("--------    ----   ----    ----")

    def listEntries(self):
        obj =  self.header.getDirectory()
        if obj != None:
            obj.parseDirEntries()
            count = 0
            for entry in obj.entries:
                print("Entry [0x%x] Name %s  Root 0x%x Left 0x%x Right %x")%( count, entry.Name, entry.DirIDRoot, entry.DirIDLeft, entry.DirIDRight )
                count = count + 1
    def list(self):
        obj =  self.header.getDirectory()
        if obj != None:
            obj.parseDirEntries()
            count = 0
            rootNode = self.__buildTree( obj.entries )            

            self.__printHeader()
            self.__printListReport( rootNode, obj.entries )
            # need to print a footer ( total bytes, total files like unzip )

    def extract(self, name):

        obj =  self.header.getDirectory()
        if obj != None:
            obj.parseDirEntries()
     
        root = self.__buildTree( obj.entries )
        node = self.__findNodeByHierachicalName( root, name )
        entry = None
        if node != None:
            entry = node.Entry 
        if  entry == None or entry.DirIDRoot > 0 :
            print "can't extract %s"%name
            return

        bytes = obj.getRawStream( entry )

        file = open(entry.Name, 'wb') 
        file.write( bytes[ 0:entry.StreamSize] )
        file.close

    def deleteEntry(self, directory, node, tree ):
        entry = node.Entry
        if ( entry == None or entry.DirIDRoot > 0 ):
            print "can't extract %s"%entry.Name
            return
        #find the chain associated with the entry
        dirEntryOffset = node.Index * 128
        sectorOffset = int( dirEntryOffset % directory.sectorSize )
        chainIndex = int( dirEntryOffset / directory.sectorSize )
                
        chainSID = directory.sectorIDs[ chainIndex ];
        pos = globals.getSectorPos(chainSID, directory.sectorSize) 
        #point at entry
        pos += sectorOffset 

        print "dirEntryOffset %d, chainIndex %d, sectorOffset %d, chainSID %d, pos %d"%(dirEntryOffset, chainIndex, sectorOffset, chainSID, pos)
        #mark the Entry as empty
        # we should make the ole class use bytearray so we can modify the inmemory model
        # instead of doing a copy here
        self.outputBytes = bytearray( self.header.bytes )
        
        nFreeSecID = struct.pack( '<l', -1 )


        self.outputBytes[pos : pos + 68] = bytearray( 68 );
        #DirIDLeft
        self.outputBytes[pos + 68:pos + 72] = nFreeSecID
        #DirIDRight
        self.outputBytes[pos + 72:pos + 76] = nFreeSecID
        #DirIDRoot
        self.outputBytes[pos + 76:pos + 80] = nFreeSecID
        self.outputBytes[pos + 80:pos + 128] = bytearray( 48 )
   
         
        #we should make the associated SAT or SSAT array entries as emtpy
        
    def delete(self, name, directory ):
        print("attempting to delete %s !!")%name
        #we'll do an inefficient delete e.g. no attempt to reclaim space 
        #in the Directory stream/sectors

        #first find the Entry for name
        root = self.__buildTree( directory.entries )
        node = self.__findNodeByHierachicalName( root, name ) 
        if node != None:
            self.deleteEntry( directory, node, root )
        self.writeDoc("/home/npower/testComp.xls", self.outputBytes)
        print("** attempting to write out compound document")
                      
    def writeDoc( self, filePath, contents ):
        out = open( filePath, 'wb' );

        if len( contents  ):
            print ("using in memory model") 
            out.write( contents )
            return
        out.write( self.header.docId )
        out.write( self.header.uId )
         
        out.write( struct.pack( '<h',self.header.revision ) )
        out.write( struct.pack( '<h',self.header.version ) )

        if self.header.byteOrder ==  ole.ByteOrder.LittleEndian:
            out.write( struct.pack( '<h', 0xFFFE ) )
        else:
            out.write( struct.pack( '>h', 0xFFFE ) )
      
        # sector size (usually 512 bytes)
        out.write( struct.pack( '<h',self.header.secSize ) )
        # short sector size (usually 64 bytes)
        out.write( struct.pack( '<h', self.header.secSizeShort ) )
        # padding
        out.write( bytearray(10) ) 
        # total number of sectors in SAT (equals the number of sector IDs 
        # stored in the MSAT).
        out.write( struct.pack( '<l', self.header.numSecSAT ) )

        out.write( struct.pack( '<l', self.header.secIDFirstDirStrm ) )
        out.write( bytearray(4) ) 
        out.write( struct.pack( '<l', self.header.minStreamSize ) )
        out.write( struct.pack( '<l', self.header.secIDFirstSSAT ) )
        out.write( struct.pack( '<l', self.header.numSecSSAT ) )
        out.write( struct.pack( '<l', self.header.secIDFirstMSAT ) )
        out.write( struct.pack( '<l', self.header.numSecMSAT ) )

        # now for the MSAT
        self.writeMSAT( self.header, out )
        out.flush()
        out.close()
                
    def writeMSAT( self, header, outFile ):
        # write the first 109 id(s) straight into the header
        msatArray  = header.getMSAT().secIDs
        nMSATWritten = 0 
        nTotalSecIDs = len( msatArray )
        nFreeSecID = struct.pack( '<l', -1 )
        for i in xrange( 0, 109 ):
            if i < nTotalSecIDs:
                nMSATWritten += 1
                outFile.write( struct.pack( '<l', msatArray[ i ] ) )
            else:
                outFile.write( nFreeSecID )         
#        if nTotalSecIDs > 109:
            # MSAT extends past the header
#            for i in xrange( 109, nTotalSecIDs );
             
        
def main ():
    parser = optparse.OptionParser()
    parser.add_option("-l", "--list", action="store_true", dest="list", default=False, help="lists ole contents")
    parser.add_option("-x", "--extract", action="store_true", dest="extract", default=False, help="extract file")
    parser.add_option("-D", "--Delete", action="store_true", dest="delete", default=False, help="delete file from document")
    parser.add_option("-d", "--debug", action="store_true", dest="debug", default=False, help="spew debug and dump file format")


    options, args = parser.parse_args()

    params = globals.Params()

    params.list =  options.list
    params.extract =  options.extract
    params.debug =  options.debug
    params.delete =  options.delete

    if len(args) < 1:
        globals.error("takes at least one arguments\n")
        parser.print_help()
        sys.exit(1)

    container =  OleContainer( args[ 0 ], params )

    if params.list == True:
        container.list() 
    if params.extract or params.delete:
       files = args
       files.pop(0)
           
       for file in files:
           if params.extract:
               container.extract( file ) 
           else:
               obj = container.header.getDirectory()
               if obj != None:
                   print "parsing dir entries"
                   obj.parseDirEntries()
               container.delete( file, obj ) 
#        container.listEntries() 
    if params.debug == True:
        container.output()

if __name__ == '__main__':
    main()
