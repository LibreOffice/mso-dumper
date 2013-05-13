#!/usr/bin/env python2
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

import sys, os.path, optparse

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

    def __init__(self, entry):
        self.Nodes = []
        self.Entry = entry;
        self.HierachicalName = ''

    def isStorage():
        return entry.Type == Directory.Type.RootStorage

class OleContainer:

    def __init__(self,filePath, params ):
        self.filePath = filePath
        self.header = None
        self.params = params
        self.pos = None
        
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

    def __parseFile (self):
        file = open(self.filePath, 'rb')
        self.chars = file.read()
        file.close()    

    def __addSiblings( self, entries, parent, child ):
        # add left siblings
        nextLeft = child.Entry.DirIDLeft
        if ( nextLeft > 0 ):
            newEntry = DirNode( entries[ nextLeft ] )
            newEntry.HierachicalName = parent.HierachicalName + globals.encodeName( newEntry.Entry.Name )
            if  newEntry.Entry.DirIDRoot > 0:
                newEntry.HierachicalName = newEntry.HierachicalName + '/'

            self.__addSiblings( entries, parent, newEntry ) 
            parent.Nodes.insert( 0, newEntry )

        nextRight = child.Entry.DirIDRight
        # add children to the right 
        if ( nextRight > 0 ):
            newEntry = DirNode( entries[ nextRight ] )
            newEntry.HierachicalName = parent.HierachicalName + globals.encodeName( newEntry.Entry.Name )
            if  newEntry.Entry.DirIDRoot > 0:
                newEntry.HierachicalName = newEntry.HierachicalName + '/'
            self.__addSiblings( entries, parent, newEntry ) 
            parent.Nodes.append( newEntry )

    def __buildTreeImpl(self, entries, parent ):

        if ( parent.Entry.DirIDRoot > 0 ):
            newEntry = DirNode( entries[ parent.Entry.DirIDRoot ] )
            newEntry.HierachicalName = parent.HierachicalName + globals.encodeName( newEntry.Entry.Name )
            if ( newEntry.Entry.DirIDRoot > 0 ):
                newEntry.HierachicalName =  newEntry.HierachicalName + '/'

            self.__addSiblings( entries, parent, newEntry )
            parent.Nodes.append( newEntry )
            
        for child in parent.Nodes:
            if child.Entry.DirIDRoot > 0:
                self.__buildTreeImpl( entries, child )

    def __buildTree(self, entries ):
        treeRoot = DirNode( entries[0] ) 
        self.__buildTreeImpl( entries, treeRoot )
        return treeRoot

    def __findEntryByHierachicalName( self, node, name ):
        if node.HierachicalName == name:
            return node.Entry
        else:
            for child in node.Nodes:
                result = self.__findEntryByHierachicalName( child, name )
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
        self.__parseFile()
        #if self.header == None:
        #    self.header = ole.Header(self.chars, self.params)
        #    self.pos = self.header.parse()
        self.header = ole.Header(self.chars, self.params)
        self.pos = self.header.parse()
        obj =  self.header.getDirectory()
        if obj != None:
            obj.parseDirEntries()
            count = 0
            for entry in obj.entries:
                print("Entry [0x%x] Name %s  Root 0x%x Left 0x%x Right %x")%( count, entry.Name, entry.DirIDRoot, entry.DirIDLeft, entry.DirIDRight )
                count = count + 1
    def list(self):
        # need to share the inititialisation and parse stuff between the different options
        self.__parseFile()
        if self.header == None:
            self.header = ole.Header(self.chars, self.params)
            self.pos = self.header.parse()
        obj =  self.header.getDirectory()
        if obj != None:
            obj.parseDirEntries()
            count = 0
            rootNode = self.__buildTree( obj.entries )            

            self.__printHeader()
            self.__printListReport( rootNode, obj.entries )
            # need to print a footer ( total bytes, total files like unzip )

    def extract(self, name):
        if  self.header == None:
            self.__parseFile()
            self.header = ole.Header(self.chars, self.params)
            self.pos = self.header.parse()

        obj =  self.header.getDirectory()
        if obj != None:
            obj.parseDirEntries()
     
        root = self.__buildTree( obj.entries )
        entry = self.__findEntryByHierachicalName( root, name )

        if  entry == None or entry.DirIDRoot > 0 :
            print "can't extract %s"%name
            return

        bytes = obj.getRawStream( entry )

        file = open(entry.Name, 'wb') 
        file.write( bytes )
        file.close

def main ():
    parser = optparse.OptionParser()
    parser.add_option("-l", "--list", action="store_true", dest="list", default=False, help="lists ole contents")
    parser.add_option("-x", "--extract", action="store_true", dest="extract", default=False, help="extract file")


    options, args = parser.parse_args()

    params = globals.Params()

    params.list =  options.list
    params.extract =  options.extract

    if len(args) < 1:
        globals.error("takes at least one arguments\n")
        parser.print_help()
        sys.exit(1)

    container =  OleContainer( args[ 0 ], params )

    if params.list == True:
        container.list() 
    if params.extract:
       files = args
       files.pop(0)
           
       for file in files:
           container.extract( file ) 
#        container.listEntries() 

if __name__ == '__main__':
    main()
