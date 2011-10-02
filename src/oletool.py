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
            newEntry.HierachicalName = newEntry.Entry.Name
            if len(  parent.HierachicalName ):
                newEntry.HierachicalName = parent.HierachicalName + '/' + newEntry.HierachicalName

            self.__addSiblings( entries, parent, newEntry ) 
            parent.Nodes.insert( 0, newEntry )

        nextRight = child.Entry.DirIDRight
        # add children to the right 
        if ( nextRight > 0 ):
            newEntry = DirNode( entries[ nextRight ], nextRight )
            newEntry.HierachicalName = newEntry.Entry.Name
            if len(  parent.HierachicalName ):
                newEntry.HierachicalName = parent.HierachicalName + '/' + newEntry.HierachicalName
            self.__addSiblings( entries, parent, newEntry ) 
            parent.Nodes.append( newEntry )

    def __buildTreeImpl(self, entries, parent ):

        if ( parent.Entry.DirIDRoot > 0 ):
            newEntry = DirNode( entries[ parent.Entry.DirIDRoot ], parent.Entry.DirIDRoot )
            if len(  parent.HierachicalName ):
                newEntry.HierachicalName = parent.HierachicalName + '/' + newEntry.HierachicalName
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

    def updateEntry( self, directory, entry, filePath ):
        file = open( filePath, 'rb' )
        bytes = file.read();
        print "Entry is ",entry.Name 
        theSAT = self.header.getSAT()
        sectorSize = self.header.getSectorSize()
        entry.StreamSize = len(bytes)
        print "stream size %d min streamsize %d"%( entry.StreamSize, self.header.minStreamSize )

        entryID = entry.StreamSectorID

        oldChain = []
        if ( entry.StreamLocation == ole.StreamLocation.SSAT ):
            theSAT = self.header.getSSAT()
        else:
            theSAT = self.header.getSAT()
        if ( entryID > -1 ):
            oldChain = theSAT.getSectorIDChain(entryID) 
            theSAT.freeChainEntries( oldChain )

        if entry.StreamSize < self.header.minStreamSize:
            print "going to use ssat"
            theSAT = self.header.getSSAT()
            entry.StreamLocation = ole.StreamLocation.SSAT
            sectorSize = self.header.getShortSectorSize()
        else:
            theSAT = self.header.getSAT()
            entry.StreamLocation = ole.StreamLocation.SAT

        #ok, find out how many sectors to store the data from the file
        sectorOffset = entry.StreamSize % sectorSize
        sectorIndex = int(  entry.StreamSize / sectorSize )            
        newNumChainEntries = sectorIndex + 1

        if ( entry.StreamLocation == ole.StreamLocation.SSAT ):
            oldNumChainEntries = len( oldChain )
            print "[ssat] new stream will take %d sectors to store %d bytes"%(sectorIndex+1,(sectorIndex+1)* sectorSize )
            if newNumChainEntries > oldNumChainEntries:
                neededEntries =  newNumChainEntries - oldNumChainEntries
                newChain = self.header.getOrAllocateFreeSSATChainEntries( neededEntries )
                
            neededEntries =  newNumChainEntries - oldNumChainEntries
            if (  neededEntries != len(newChain) ):
                raise Exception("no space available")
            #find the highest index to see if we need to expand the root
            #storage
            maxIndex = 0
            for i in xrange(0,len(newChain) ):
                current = newChain[ i ]
                if i == 0:
                    maxIndex = current
                else:
                    if current > maxIndex:
                        maxIndex = current 
            #numer of sectors the maxIndex represents is + 1 ( e,g, and index of 0
            #means at least one sector
            maxIndex += 1
            if directory.ensureRootStorageCapacity( maxIndex * self.header.getShortSectorSize() ):
                directory.writeToShortSectors( newChain, bytes )
            else:
                raise Exception("no space available")
            entry.StreamSectorID = newChain[ 0 ]
        else:
            chain = self.header.getSAT().getFreeChainEntries( newNumChainEntries, 0 )
            # populate and terminate the chain
            lastIndex =  chain[ len( chain ) - 1 ]
            self.header.getSAT().array[ lastIndex ] = -2
            for i in xrange(0,len(chain) ):
                if i > 0:
                    self.header.getSAT().array[ chain[ i-1 ] ] = chain[ i ]
            
            # updating normal sectors
            directory.writeToStdSectors( chain, bytes )
            entry.StreamSectorID = chain[ 0 ]

    def makeEntryEmpty( self, entry ):
        # #FIXME can we clone an Entry somehow <sigh> python knowledge fail
        # clear entry
        entry.Name = ''
        entry.CharBufferSize = 0
        entry.Type = ole.Directory.Type.Empty
        entry.NodeColor = ole.Directory.NodeColor.Red
        entry.DirIDLeft = -1
        entry.DirIDRight = -1
        entry.DirIDRoot = -1
        entry.UniqueID = bytearray(16) 
        entry.UserFlags =  bytearray(4)
        entry.TimeCreated =  bytearray(8)
        entry.TimeModified =  bytearray(8)
        entry.StreamSectorID = 0
        entry.StreamSize = 0

    def deleteEntry(self, directory, node, tree ):
        entry = node.Entry
        if ( entry.Type == None ):
            print "can't extract %s"%entry.Name
            return

        theSAT = self.header.getSAT()
        if entry.StreamLocation == ole.StreamLocation.SSAT:
            print ("using SSAT")
            theSAT = self.header.getSSAT()
        elif entry.StreamLocation == ole.StreamLocation.SSAT:
            print ("using SAT")
        
        chain = theSAT.getSectorIDChain(entry.StreamSectorID) 
        self.makeEntryEmpty( entry )

        theSAT.freeChainEntries( chain )
 
        # #FIXME what about references ( e.g. parent of this entry should we 
        # blank the associated left/right id(s) ) - leave for now  

        #if this node has children then I suppose we need to delete them too
        for child in node.Nodes: 
            self.deleteEntry( directory, child, tree )
    def compareNodes( self, lhs, rhs ):
        if len( lhs.Name ) > len( rhs.Name ):
            return 1
        elif len( lhs.Name ) > len ( rhs.Name ):
            return 0
        else: #equal
            if lhs.Name > rhs.Name:
                return  1
            else:
                return  0

    def insertSiblingInTree ( self, directory, dirID, node, otherNode ) :
        if self.compareNodes( node, otherNode ):
            if otherNode.DirIDRight > 0:
                if ( self.compareNodes( node, directory.entries[ otherNode.DirIDRight ] ) ):
                    self.insertSiblingInTree( directory, dirID, node, directory.entries[ otherNode.DirIDRight ] )
                else:
                    self.insertSiblingInTree( directory, dirID, node, directory.entries[ otherNode.DirIDLeft ] )
            else:           
                otherNode.DirIDRight = dirID
        else:
            if otherNode.DirIDLeft > 0:
                if ( self.compareNodes( node, directory.entries[ otherNode.DirIDLeft ] ) ):
                    self.insertSiblingInTree( directory, dirID, node, directory.entries[ otherNode.DirIDLeft ] )
                else:
                    self.insertSiblingInTree( directory, dirID, node, directory.entries[ otherNode.DirIDRight ] )
            else:           
                otherNode.DirIDLeft = dirID

    def insertDirEntry( self, directory, root, childName ):
        # have we space for a new entry
        if len( directory.entries ) + 1 > ( ( len( directory.sectorIDs ) * directory.sectorSize ) ) / 128:
            # allocate a new SAT sector to increase the size of Directory
            oldChain = directory.header.getSAT().getSectorIDChain( directory.header.getFirstSectorID(ole.BlockType.Directory) )
            newChainEntries = directory.header.getSAT().getFreeChainEntries( 1,0 )
            directory.header.createSATSectors( newChainEntries, False )

            if len( newChainEntries ) == 0:
                # looks like we need to expand the MSAT to expand the SAT array
                raise Exception("expanding MSAT not supported yet")
            directory.addSector(newChainEntries[0])
            # add newChain to old chain
            lastIndex =  oldChain[ len( oldChain ) - 1 ]
            for satID in newChainEntries:
                directory.header.getSAT().array[ lastIndex ] = satID
                lastIndex = satID
            # terminate the chain
            directory.header.getSAT().array[ newChainEntries[ len( newChainEntries ) - 1 ] ] = -2
            
        # create new DirEntry
        lastEntry = directory.entries[ len( directory.entries ) - 1]
        entry = ole.Directory.Entry()
        self.makeEntryEmpty( entry )
        directory.entries.append( entry )
        dirID = len( directory.entries) - 1 
        # if the old last entry was empty use that other wise use
        # the new one
        if lastEntry.Type == ole.Directory.Type.Empty:
            dirID = len( directory.entries ) - 2
            entry = lastEntry

        entry.Name = childName
        self.insertSiblingInTree ( directory, dirID, entry, directory.entries[ root.Entry.DirIDRoot ] ) 
        # fill with empty entries to fill the sector
        mod = len(directory.entries ) % 4
        if mod:
            for i in xrange(0,4-mod):
                empty = ole.Directory.Entry()
                self.makeEntryEmpty( empty )
                directory.entries.append( empty )
                 
        return entry

    def add(self, filePath, directory):
        if os.path.isabs( filePath ) != True:
            filePath = os.path.abspath( filePath )
        print "Add %s"%filePath

        if os.path.isdir( filePath ):
           print "can't yet handle adding storage for dir %s"%filePath
           return;

        storageTestCreate = False

        if os.path.isfile( filePath ) != True and storageTestCreate !=True:
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
            self.updateEntry( directory, node.Entry, filePath )
        else:
            #storage
            node = self.__findNodeByHierachicalName( root, dirPath )
            print "adding a new file to %s"%dirPath
            entry = self.insertDirEntry( directory, node, fileLeafName )
            if  storageTestCreate:
                entry.NodeColor = ole.Directory.NodeColor.Black
                entry.Type = directory.Type.UserStorage
            else:
                #FIXME how to allocate the NodeColor
                entry.NodeColor = node.Entry.NodeColor
                entry.NodeColor = node.Entry.NodeColor
                entry.Type = directory.Type.UserStream
                self.updateEntry( directory, entry, filePath )
            
        self.header.write() 
        directory.write() 
        self.writeDoc("/home/npower/testComp.xls", self.header.bytes)

    def delete(self, name, directory ):
        #we'll do an inefficient delete e.g. no attempt to reclaim space 
        #in the Directory stream/sectors

        #first find the Entry for name
        root = self.__buildTree( directory.entries )
        node = self.__findNodeByHierachicalName( root, name ) 
        if node != None:
            self.deleteEntry( directory, node, root )

        sizeBefore = len( self.header.bytes )
        #self.header.write() 
        directory.write() 
        sizeAfter = len( self.header.bytes )
        self.writeDoc("/home/npower/testComp.xls", self.header.bytes)
                      
    def writeDoc( self, filePath, contents ):
        out = open( filePath, 'wb' );

        if len( contents  ):
            out.write( contents )
        else:
            print ("failed to write document") 
        out.flush()
        out.close()
                
        
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
