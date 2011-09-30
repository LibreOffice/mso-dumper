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


class OleContainer:

    def __init__(self,filePath, params ):
        self.filePath = filePath
        self.params = params
        self.file = open(self.filePath, 'rb')
        self.header = ole.Header(self.file.read(), self.params)
        self.file.close()
        self.pos = self.header.parse()

    def output (self, directory):
        
        # #FIXME common initialisation code
        self.header.output() 
        directory.output()

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
            newEntry.HierachicalName = globals.encodeName( newEntry.Entry.Name )
            if len(  parent.HierachicalName ):
                newEntry.HierachicalName = parent.HierachicalName + '/' + newEntry.HierachicalName

            self.__addSiblings( entries, parent, newEntry ) 
            parent.Nodes.insert( 0, newEntry )

        nextRight = child.Entry.DirIDRight
        # add children to the right 
        if ( nextRight > 0 ):
            newEntry = DirNode( entries[ nextRight ], nextRight )
            newEntry.HierachicalName = globals.encodeName( newEntry.Entry.Name )
            if len(  parent.HierachicalName ):
                newEntry.HierachicalName = parent.HierachicalName + '/' + newEntry.HierachicalName
            self.__addSiblings( entries, parent, newEntry ) 
            parent.Nodes.append( newEntry )

    def __buildTreeImpl(self, entries, parent ):

        if ( parent.Entry.DirIDRoot > 0 ):
            newEntry = DirNode( entries[ parent.Entry.DirIDRoot ], parent.Entry.DirIDRoot )
            newEntry.HierachicalName = parent.HierachicalName + globals.encodeName( newEntry.Entry.Name )
            if ( newEntry.Entry.DirIDRoot > 0 ):
                newEntry.HierachicalName =  newEntry.HierachicalName 

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
            aflag = "f"
            if treeNode.Entry.isStorage():
                aflag = "d"
            print '{0} {1:8d}  {2:0<2d}-{3:0<2d}-{4:0<2d} {5:0<2d}:{6:0<2d}   {7}'.format( aflag,treeNode.Entry.StreamSize, dateInfo.day, dateInfo.month, dateInfo.year, dateInfo.hour, dateInfo.second, treeNode.HierachicalName )
     
        for node in treeNode.Nodes:
            # ignore the root
            self.__printListReport( node, stats )

    def __printHeader(self):
        print ("OLE: %s")%self.filePath
        print ("   Length     Date   Time    Name")
        print ("  --------    ----   ----    ----")

    def listEntries(self):
        obj =  self.header.getDirectory()
        if obj != None:
            obj.parseDirEntries()
            count = 0
            for entry in obj.entries:
                print("Entry [0x%x] Name %s  Root 0x%x Left 0x%x Right %x")%( count, entry.Name, entry.DirIDRoot, entry.DirIDLeft, entry.DirIDRight )
                count = count + 1
    def list(self, directory):
        if directory != None:
            count = 0
            rootNode = self.__buildTree( directory.entries )            

            self.__printHeader()
            self.__printListReport( rootNode, directory.entries )
            # need to print a footer ( total bytes, total files like unzip )

    def extract(self, name, directory):
        if ( directory == None ):
            print "failed to extract %s"%name
            return
        root = self.__buildTree( directory.entries )
        node = self.__findNodeByHierachicalName( root, name )
        entry = None
        if node != None:
            entry = node.Entry 
        if  entry == None or entry.DirIDRoot > 0 :
            print "can't extract %s"%name
            return

        bytes = directory.getRawStream( entry )

        file = open(entry.Name, 'wb') 
        file.write( bytes[ 0:entry.StreamSize] )
        file.close

    def updateEntry( self, directory, node, filePath ):
        file = open( filePath, 'rb' )
        bytes = file.read();
        entry = node.Entry
        
        theSAT = self.header.getSAT()
        sectorSize = self.header.getSectorSize()
        # do we need to change the filestream location ?
        if entry.StreamLocation == ole.StreamLocation.SSAT:
            print ("updateEntry using SSAT")
            theSAT = self.header.getSSAT()
            sectorSize = self.header.getShortSectorSize()
        elif entry.StreamLocation == ole.StreamLocation.SAT:
            print ("updateEntry using SAT")         

        #ok, find out how many sectors to store the data from the file
        streamSize = len(bytes)
        sectorOffset = streamSize % sectorSize
        sectorIndex = int(  streamSize / sectorSize )            

        #compare with the existing number of sectors
        chain = theSAT.getSectorIDChain(entry.StreamSectorID) 
        oldNumChainEntries = len( chain )
        print "chain contains ", chain
        print "size of stream to update %d size of existing entry %d size of sectors %d"%(streamSize, entry.StreamSize, len(chain) * sectorSize )
        print "new stream will take %d sectors to store %d bytes"%(sectorIndex+1,(sectorIndex+1)* sectorSize )
        newNumChainEntries = sectorIndex + 1
        if newNumChainEntries > oldNumChainEntries:
            neededEntries =  newNumChainEntries - oldNumChainEntries
            newChain = self.getFreeSATChainEntries( theSAT, neededEntries, 0 )
            print "received %d of %d chain entries needed"%(len(newChain), neededEntries ) 
            if len(newChain) < neededEntries:
                # create enough sectors to increase SAT table to accomadate new entries
                numEntriesPerSector = int( theSAT.sectorSize / 4 )
                numExtraEntriesNeeded = neededEntries - len( newChain )
                numNeededSectors = int ( ( numExtraEntriesNeeded + ( numExtraEntriesNeeded % numEntriesPerSector ) ) / numEntriesPerSector )
                self.expandSAT( numNeededSectors, entry.StreamLocation == ole.StreamLocation.SAT )
                

    def getFreeSATChainEntries( self, theSAT, numNeeded, searchFrom ):
        freeSATChain = []
        maxEntries = ( len( theSAT.sectorIDs ) * theSAT.sectorSize ) / 4
        for i in xrange( 0, len( theSAT.array ) ):
            if numNeeded == 0:
                break
            nSecID = theSAT.array[i];
            if nSecID == -1: 
                #have we exceeded the allocated size of the SAT table ?
                if i < maxEntries:
                    freeSATChain.append( i )
                    numNeeded -= 1
                else:
                    break;
        return freeSATChain

    def appendFreeChain( self, chain1, chain2, theSAT ):
        #chain2 is a list of free indexes 
        #change the list in memory as well as internal structures
        lastIndex =  chain1[ len( chain1 ) - 1 ]
        if len( chain2 ):
            #modify memory
            pos = self.memPosOfSATChainIndex( lastIndex, theSAT ) 
            for entry in chain2:
                self.header.bytes[ pos : pos + 4 ] = struct.pack( '<l', entry )
                pos = self.memPosOfSATChainIndex( entry, theSAT )
            pos = self.memPosOfSATChainIndex(  chain2[ len( chain2 ) - 1 ], theSAT )
            self.header.bytes[ pos : pos + 4 ] = struct.pack( '<l', -2 )

            #modify model
            for entry in chain2:
                print "array[%d] assingin value %d"%( lastIndex, entry )
                theSAT.array[ lastIndex ] = entry
                lastIndex = entry
                theSAT.sectorIDs.append( entry )
        theSAT.array[ chain2[ len( chain2 ) - 1 ] ] = -2

    def memPosOfSATChainIndex( self, index, theSAT ):
        #find the position in memory of the index into the SAT table
        #each index takes up 4 bytes so sat[ entry ] would be at 
        #position ( 4 * index )
        entryPos = 4 * index
        #calculate the offset into a sector
        sectorOffset = entryPos %  theSAT.sectorSize
        sectorIndex = int(  entryPos /  theSAT.sectorSize )
        sectorSize = theSAT.sectorSize
        #now point to the offset in the sector this array position lives
        #in
        pos = 512 + ( theSAT.sectorIDs[ sectorIndex ] * theSAT.sectorSize ) + sectorOffset
        return pos
                
    def freeSATChainEntries( self, chain, theSAT, mem ):
        #we need to calculate the position in the file stream that each array
        #position in the stream corrosponds to
        nFreeSecID = struct.pack( '<l', -1 )
        for entry in chain:
            pos = self.memPosOfSATChainIndex( entry, theSAT )
            mem[pos:pos + 4] = nFreeSecID

    def findEntryMemPos( self, node, directory ):
        #find the chain associated with the entry
        dirEntryOffset = node.Index * 128
        sectorOffset = int( dirEntryOffset % directory.sectorSize )
        chainIndex = int( dirEntryOffset / directory.sectorSize )
                
        chainSID = directory.sectorIDs[ chainIndex ];
        pos = globals.getSectorPos(chainSID, directory.sectorSize) 
        #point at entry
        pos += sectorOffset 
        return pos

    def expandSAT( self, numSectors, isSAT ):
        #theSAT could be the SSAT or the SAT, doesn't matter we 
        #still need a new sector to extend whichever SAT table
        print "Attempt to expand SAT by %d sectors"%numSectors
        # are there any available entries in the SAT ?
        newChain = self.getFreeSATChainEntries( self.header.getSAT(), numSectors, 0 ) 

        if len( newChain ) == 0:
            print "Error, we need don't support extending the MSAT yet"
            return 
        print "received %d of %d chain entries needed"%(len(newChain), numSectors ) 
        print "-> ", newChain 
     
        #new sectors implies we need to fill them with -1(s) ( or even create the sector if it doesn't exist ) 
        oldChain = [] 
        if isSAT:
            #modify SAT chain to incorporate new secID
            oldChain = self.header.getSAT().getSectorIDChain( self.header.getFirstSectorID(ole.BlockType.SAT) )
            print "old SAT chain -> ", oldChain 
        else:
            oldChain = self.header.getSAT().getSectorIDChain( self.header.getFirstSectorID(ole.BlockType.SSAT) )
            print "old SSAT chain -> ", oldChain 

        self.appendFreeChain( oldChain, newChain, self.header.getSAT() )
        
        if isSAT:
            oldChain = self.header.getSAT().getSectorIDChain( self.header.getFirstSectorID(ole.BlockType.SAT) )
            print "new SAT chain -> ", oldChain 
        else:
            oldChain = self.header.getSAT().getSectorIDChain( self.header.getFirstSectorID(ole.BlockType.SSAT) )
            print "new SSAT chain -> ", oldChain 

    def deleteEntry(self, directory, node, tree ):
        entry = node.Entry
        if ( entry.Type == None ):
            print "can't extract %s"%entry.Name
            return
        print("** attempting to delete %s"%entry.Name)
        #point at entry
        pos = self.findEntryMemPos( node, directory )

        #mark the Entry as empty
        nFreeSecID = struct.pack( '<l', -1 )

        self.header.bytes[pos : pos + 68] = bytearray( 68 );
        #DirIDLeft
        self.header.bytes[pos + 68:pos + 72] = nFreeSecID
        #DirIDRight
        self.header.bytes[pos + 72:pos + 76] = nFreeSecID
        #DirIDRoot
        self.header.bytes[pos + 76:pos + 80] = nFreeSecID
        self.header.bytes[pos + 80:pos + 128] = bytearray( 48 )

        #  make the associated SAT or SSAT array entries as emtpy
        theSAT = self.header.getSAT()

        if entry.StreamLocation == ole.StreamLocation.SSAT:
            print ("using SSAT")
            theSAT = self.header.getSSAT()
        elif entry.StreamLocation == ole.StreamLocation.SSAT:
            print ("using SAT")
        
        chain = theSAT.getSectorIDChain(entry.StreamSectorID) 
        self.freeSATChainEntries( chain, theSAT, self.header.bytes )               
 
        # #FIXME what about references ( e.g. parent of this entry should we 
        # blank the associated left/right id(s) ) - leave for now  

        #if this node has children then I suppose we need to delete them too
        for child in node.Nodes: 
            self.deleteEntry( directory, child, tree )

    def add(self, filePath, directory):
        if os.path.isabs( filePath ) != True:
            filePath = os.path.abspath( filePath )
        print "Add %s"%filePath

        if os.path.isdir( filePath ):
           print "can't yet handle adding storage for dir %s"%filePath
           return;
        if os.path.isfile( filePath ) != True:
           print "%s is not a file, bailing out"%filePath
           return
        nodePath = os.path.relpath( filePath )
        fileLeafName =  os.path.basename(nodePath)
        print "node path is %s"%nodePath
        print "basename is %s"%fileLeafName
        if nodePath[0] == ".":
            print "warning path %s is not relative to current path, using %s instead"%(nodePath,fileLeafName)
            dirPath = ''
        else:
            dirPath = os.path.dirname(nodePath)

        print "dirname is %s"%dirPath

        root= self.__buildTree( directory.entries )            
        node = self.__findNodeByHierachicalName( root, nodePath )

        if  node != None:
            #update a file
            print "updating %s"%fileLeafName
            self.updateEntry( directory, node, filePath )
        else:
            #new file or storage
            node = self.__findNodeByHierachicalName( root, dirPath )
            if node == None: 
                print "error: %s does not exist"%dirPath
                return
            print "adding a new file to %s"%dirPath

    def delete(self, name, directory ):
        #we'll do an inefficient delete e.g. no attempt to reclaim space 
        #in the Directory stream/sectors

        #first find the Entry for name
        root = self.__buildTree( directory.entries )
        node = self.__findNodeByHierachicalName( root, name ) 
        if node != None:
            self.deleteEntry( directory, node, root )
        self.writeDoc("/home/npower/testComp.xls", self.header.bytes)
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
    parser.add_option("-a", "--add", action="store_true", dest="add", default=False, help="adds or updates entry")
    parser.add_option("-d", "--debug", action="store_true", dest="debug", default=False, help="spew debug and dump file format")


    options, args = parser.parse_args()

    params = globals.Params()

    params.list =  options.list
    params.extract =  options.extract
    params.debug =  options.debug
    params.delete =  options.delete
    params.add =  options.add

    if len(args) < 1:
        globals.error("takes at least one arguments\n")
        parser.print_help()
        sys.exit(1)

    container =  OleContainer( args[ 0 ], params )
    directory = container.header.getDirectory()
    if directory != None:
        directory.parseDirEntries()

    if params.list == True:
        container.list( directory ) 
    if params.extract or params.delete or params.add:
       files = args
       files.pop(0)
           
       for file in files:
           if params.extract:
               container.extract( file, directory ) 
           elif params.delete:
               container.delete( file, directory ) 
           else:
               container.add( file, directory )
    if params.debug == True:
#        container.listEntries() 
        container.output( directory )

if __name__ == '__main__':
    main()
