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
import globals
from globals import getSignedInt
# ----------------------------------------------------------------------------
# Reference: The Microsoft Compound Document File Format by Daniel Rentz
# http://sc.openoffice.org/compdocfileformat.pdf
# ----------------------------------------------------------------------------

from globals import output
import struct

class TypeValue:
        minusOne = struct.pack( '<l', -1 )

class NoRootStorage(Exception): pass

class ByteOrder:
    LittleEndian = 0
    BigEndian    = 1
    Unknown      = 2

class BlockType:
    MSAT      = 0
    SAT       = 1
    SSAT      = 2
    Directory = 3

class StreamLocation:
    SAT  = 0
    SSAT = 1

class Header(object):

    @staticmethod
    def byteOrder (chars):
        b1, b2 = chars[0], chars[1]
        if b1 == 0xFE and b2 == 0xFF:
            return ByteOrder.LittleEndian
        elif b1 == 0xFF and b2 == 0xFE:
            return ByteOrder.BigEndian
        else:
            return ByteOrder.Unknown


    def __init__ (self, bytes, params):
        self.bytes = bytearray(bytes)
        self.MSAT = None

        self.docId = None
        self.uId = None
        self.revision = 0
        self.version = 0
        self.byteOrder = ByteOrder.Unknown
        self.minStreamSize = 0

        self.numSecMSAT = 0
        self.numSecSSAT = 0
        self.numSecSAT = 0

        self.secIDFirstMSAT = -2
        self.secIDFirstDirStrm = -2
        self.secIDFirstSSAT = -2

        self.secSize = 512
        self.secSizeShort = 64

        self.params = params

    def getSectorSize (self):
        return 2**self.secSize


    def getShortSectorSize (self):
        return 2**self.secSizeShort


    def getFirstSectorID (self, blockType):
        if blockType == BlockType.MSAT:
            return self.secIDFirstMSAT
        elif blockType == BlockType.SSAT:
            return self.secIDFirstSSAT
        elif blockType == BlockType.Directory:
            return self.secIDFirstDirStrm
        return -2


    def output (self):

        def printRawBytes (bytes):
            bytes = bytes.decode('windows-1252')
            for b in bytes:
                output("%2.2X "%ord(b))
            output("\n")

        def printSep (c, w, prefix=''):
            print(prefix + c*w)

        printSep('=', globals.OutputWidth)
        print("Compound Document Header")
        printSep('-', globals.OutputWidth)

        if self.params.debug:
            globals.dumpBytes(self.bytes[0:512].decode('windows-1252'))
            printSep('-', globals.OutputWidth)

        # document ID and unique ID
        output("Document ID: ")
        printRawBytes(self.docId)
        output("Unique ID: ")
        printRawBytes(self.uId)

        # revision and version
        print("Revision: %d  Version: %d"%(self.revision, self.version))

        # byte order
        output("Byte order: ")
        if self.byteOrder == ByteOrder.LittleEndian:
            print("little endian")
        elif self.byteOrder == ByteOrder.BigEndian:
            print("big endian")
        else:
            print("unknown")

        # sector size (usually 512 bytes)
        print("Sector size: %d (%d)"%(2**self.secSize, self.secSize))

        # short sector size (usually 64 bytes)
        print("Short sector size: %d (%d)"%(2**self.secSizeShort, self.secSizeShort))

        # total number of sectors in SAT (equals the number of sector IDs 
        # stored in the MSAT).
        print("Total number of sectors used in SAT: %d"%self.numSecSAT)

        print("Sector ID of the first sector of the directory stream: %d"%
              self.secIDFirstDirStrm)

        print("Minimum stream size: %d"%self.minStreamSize)

        if self.secIDFirstSSAT == -2:
            print("Sector ID of the first SSAT sector: [none]")
        else:
            print("Sector ID of the first SSAT sector: %d"%self.secIDFirstSSAT)

        print("Total number of sectors used in SSAT: %d"%self.numSecSSAT)

        if self.secIDFirstMSAT == -2:
            # There is no more sector ID stored outside the header.
            print("Sector ID of the first MSAT sector: [end of chain]")
        else:
            # There is more sector IDs than 109 IDs stored in the header.
            print("Sector ID of the first MSAT sector: %d"%(self.secIDFirstMSAT))

        print("Total number of sectors used to store additional MSAT: %d"%self.numSecMSAT)


    def parse (self):

        # document ID and unique ID
        self.docId = self.bytes[0:8]
        self.uId = self.bytes[8:24]

        # revision and version
        self.revision = getSignedInt(self.bytes[24:26])
        self.version = getSignedInt(self.bytes[26:28])

        # byte order
        self.byteOrder = Header.byteOrder(self.bytes[28:30])

        # sector size (usually 512 bytes)
        self.secSize = getSignedInt(self.bytes[30:32])

        # short sector size (usually 64 bytes)
        self.secSizeShort = getSignedInt(self.bytes[32:34])

        # total number of sectors in SAT (equals the number of sector IDs 
        # stored in the MSAT).
        self.numSecSAT = getSignedInt(self.bytes[44:48])

        self.secIDFirstDirStrm = getSignedInt(self.bytes[48:52])
        self.minStreamSize = getSignedInt(self.bytes[56:60])
        self.secIDFirstSSAT = getSignedInt(self.bytes[60:64])
        self.numSecSSAT = getSignedInt(self.bytes[64:68])
        self.secIDFirstMSAT = getSignedInt(self.bytes[68:72])
        self.numSecMSAT = getSignedInt(self.bytes[72:76])

        # master sector allocation table
        self.MSAT = MSAT(2**self.secSize, self.bytes, self.params)

        # First part of MSAT consisting of an array of up to 109 sector IDs.
        # Each sector ID is 4 bytes in length.
        for i in xrange(0, 109):
            pos = 76 + i*4
            id = getSignedInt(self.bytes[pos:pos+4])
            if id == -1:
                break

            self.MSAT.appendSectorID(id)

        if self.secIDFirstMSAT != -2:
            # additional sectors are used to store more SAT sector IDs.
            secID = self.secIDFirstMSAT
            size = self.getSectorSize()
            inLoop = True
            while inLoop:
                pos = 512 + secID*size
                bytes = self.bytes[pos:pos+size]
                n = int(size/4)
                for i in xrange(0, n):
                    pos = i*4
                    id = getSignedInt(bytes[pos:pos+4])
                    if id < 0:
                        inLoop = False
                        break
                    elif i == n-1:
                        # last sector ID - points to the next MSAT sector.
                        secID = id
                        break
                    else:
                        self.MSAT.appendSectorID(id)

        return 512 

    def write (self):
        self.bytes[0:8] = self.docId
        self.bytes[8:24] = self.uId
         
        self.bytes[24:26] = struct.pack( '<h',self.revision )
        self.bytes[26:28] = struct.pack( '<h',self.version )

        byteOrder = struct.pack( '>h', 0xFFFE )
        if self.byteOrder ==  ByteOrder.LittleEndian:
            byteOrder = struct.pack( '<h', 0xFFFE )
        self.bytes[28:30] = byteOrder  

        # sector size (usually 512 bytes)
        self.bytes[30:32] = struct.pack( '<h',self.secSize )
        # short sector size (usually 64 bytes)
        self.bytes[32:34] = struct.pack( '<h', self.secSizeShort )
        # padding
        self.bytes[34:44] = bytearray(10)
        # total number of sectors in SAT (equals the number of sector IDs 
        # stored in the MSAT).
        self.bytes[44:48] = struct.pack( '<l', self.numSecSAT )

        self.bytes[48:52] = struct.pack( '<l', self.secIDFirstDirStrm )
        self.bytes[52:56] = bytearray(4)
        self.bytes[56:60] = struct.pack( '<l', self.minStreamSize )
	self.bytes[60:64] = struct.pack( '<l', self.secIDFirstSSAT )
        self.bytes[64:68] = struct.pack( '<l', self.numSecSSAT )
        self.bytes[68:72] = struct.pack( '<l', self.secIDFirstMSAT )
        self.bytes[72:76] = struct.pack( '<l', self.numSecMSAT )
        # write the MSAT, SAT & SSAT
        self.writeMSAT()
        self.getSAT().write()
        self.getSSAT().write()
   
    def writeMSAT (self): 
        # part of MSAT in header
        numIds = len( self.MSAT.secIDs )
        TypeValue.minusOne = struct.pack( '<l', -1 )
        for i in xrange(0, 109):
            pos = 76 + i*4
            if i < numIds:
                self.bytes[pos:pos+4] = struct.pack( '<l', self.MSAT.secIDs[ i ] )
            else:
                self.bytes[pos:pos+4] = TypeValue.minusOne
        # are additional sectors are used to store more SAT sector IDs.
        if self.secIDFirstMSAT != -2:
            secID = self.secIDFirstMSAT
            size = self.getSectorSize()
            # we have to assume if the MSAT has been grown that the memory has
            # been kept up to date, e.g. that the nextMSAT entry has been updated
            # in the previous MSAT sector
            entriesPerSector = int ( size / 4  )  - 1
            previousSectorIndex = 0
            for i in xrange( 0, len( self.secIDs ) ):
                sectorIndex = int( i/entriesPerSector )
                sectorOffSet = ( i % entriesPerSector ) * 4
                pos = 512 + secID*size + sectorOffSet
                if previousSectorIndex != sectorIndex:
                    secID = getSignedInt(self.bytes[pos:pos+4])
                    pos = 512 + secID*size + sectorOffSet
                self.bytes[pos:pos+4] = self.secIDs[ i ]
                previousSectorIndex = sectorIndex

    def getMSAT (self):
        return self.MSAT


    def getSAT (self):
        return self.MSAT.getSAT()


    def getSSAT (self):
        ssatID = self.getFirstSectorID(BlockType.SSAT)
        if ssatID < 0:
            return None
        chain = self.getSAT().getSectorIDChain(ssatID)
        if len(chain) == 0:
            return None
        obj = SSAT(2**self.secSize, self.bytes, self.params)
        for secID in chain:
            obj.addSector(secID)
        obj.buildArray()
        return obj

    def expandSSAT( self, numExtraEntriesNeeded ):
        # create enough sectors to increase SAT table to accomadate new entries
        numEntriesPerSector = int( self.getSSAT().sectorSize / 4 )
        numNeededSectors = int ( numExtraEntriesNeeded / numEntriesPerSector ) + 1              
        # need to get free sector(s) from SAT to expand the SSAT
        newSATSectors = self.getSAT().getFreeChainEntries( numNeededSectors, 0 )
        if len( newSATSectors ) == 0:
            # looks like we need to expand the MSAT to expand the SAT array
            raise Exception("expanding MSAT not supported yet")
        self.createSATSectors( newSATSectors ) 
        # add the sectors into the SSAT
        for sectorID in newSATSectors:
            self.getSSAT().sectorIDs.append( sectorID )
        # expand SSAT array with the contens of the new SATSectors
        self.getSSAT().appendArray( newSATSectors )
        # need to update the SectorIDChain for the SSAT
        oldSSATChain = self.getSAT().getSectorIDChain( self.getFirstSectorID(BlockType.SSAT) )
        lastIndex =  oldSSATChain[ len( oldSSATChain ) - 1 ]
        for entry in newSATSectors:
            self.getSAT().array[ lastIndex ] = entry
            lastIndex = entry
        # terminate the chain
        self.getSAT().array[ newSATSectors[ len( newSATSectors ) - 1 ] ] = -2

    def getOrAllocateFreeSATChainEntries( self, numNeeded ):
        chain = []
        #see how many free chains we can get
        chain = self.getSAT().getFreeChainEntries( numNeeded, 0 )
        currentFree = len(chain)
        print "*** here needed %d got %d"%(numNeeded, currentFree )
        if currentFree < numNeeded:
            #need to allocate a number of sectors to expand the SAT table
            numSATEntriesPerSector = ( self.getSAT().getSectorSize() / 4 )
            numNeededSectors =  int(numNeeded/numSATEntriesPerSector) + 1
            print "need",numNeeded,"/",numSATEntriesPerSector,"=",numNeededSectors,"from MSAT"
            #try and allocate sector to use from the SAT
            MSATSectors = self.getSAT().getFreeChainEntries( numNeededSectors, 0 )
            if len(MSATSectors) != numNeededSectors:
                print "Error: haven't implemented expanding the SAT to allow more sectors to be allocated for the MSAT"
                return chain
            self.createSATSectors( MSATSectors )

            #is there room in the MSAT table header part
            if len( self.getMSAT().secIDs ) + len( MSATSectors ) < 109:
                for sector in MSATSectors:
                    self.getMSAT().appendSectorID( sector )
                    self.getSAT().appendArray( MSATSectors )
            else:
                print "*** extending the MSAT not supported yes"
            #try again
            chain = self.getSAT().getFreeChainEntries( numNeeded, 0 )
                     
             
        return chain
    def getOrAllocateFreeSSATChainEntries (self, numNeeded):
        chain = []
        #see how many free chains we can get
        chain = self.getSSAT().getFreeChainEntries( numNeeded, 0 )
        currentFree = len(chain)
        print "*** here needed %d got %d"%(numNeeded, currentFree )
        if currentFree < numNeeded:
            #need to allocate a number of sectors to expand the SSAT table
            numExtraEntriesNeeded =  numNeeded - currentFree
            self.expandSSAT( numExtraEntriesNeeded )
            chain = self.getSSAT().getFreeChainEntries( numNeeded, 0 )
        return chain

    def createSATSectors( self, sectorIDs, initWithFreeId = True ):
        for secID in sectorIDs:
            pos = globals.getSectorPos(secID, self.getSAT().sectorSize)
            if pos > len( self.bytes ):
                raise Exception("illegal sector, non contiguous")
            if pos == len( self.bytes ):
                # sector doesn't yet exist in memory allocate it
                sector = bytearray( self.getSAT().sectorSize )
                self.bytes += sector
            if  initWithFreeId:
                for i in xrange( 0, self.getSAT().sectorSize / 4 ):
                    beginpos = pos + (i * 4 )
                    self.bytes[ beginpos : beginpos + 4 ] = TypeValue.minusOne
         
    def getDirectory (self):
        dirID = self.getFirstSectorID(BlockType.Directory)
        if dirID < 0:
            return None
        chain = self.getSAT().getSectorIDChain(dirID)
        if len(chain) == 0:
            return None
        obj = Directory(self, self.params)
        for secID in chain:
            obj.addSector(secID)
        return obj


    def dummy ():
        pass




class MSAT(object):
    """Master Sector Allocation Table (MSAT)

This class represents the master sector allocation table (MSAT) that stores 
sector IDs that point to all the sectors that are used by the sector 
allocation table (SAT).  The actual SAT are to be constructed by combining 
all the sectors pointed by the sector IDs in order of occurrence.
"""
    def __init__ (self, sectorSize, bytes, params):
        self.sectorSize = sectorSize
        self.secIDs = []
        self.bytes = bytes
        self.__SAT = None

        self.params = params

    def appendSectorID (self, id):
        self.secIDs.append(id)

    def output (self):
        print('')
        print("="*globals.OutputWidth)
        print("Master Sector Allocation Table (MSAT)")
        print("-"*globals.OutputWidth)

        for id in self.secIDs:
            print("sector ID: %5d   (pos: %7d)"%(id, 512+id*self.sectorSize))

    def getSATSectorPosList (self):
        list = []
        for id in self.secIDs:
            pos = 512 + id*self.sectorSize
            list.append([id, pos])
        return list

    def getSAT (self):
        if self.__SAT != None:
            return self.__SAT

        obj = SAT(self.sectorSize, self.bytes, self.params)
        print "**** creating  SAT "
        for id in self.secIDs:
            print "**** adding %d to SAT sectors"%id
            obj.addSector(id)
        obj.buildArray()
        self.__SAT = obj
        return self.__SAT

class SAT(object):
    """Sector Allocation Table (SAT)
"""
    def __init__ (self, sectorSize, bytes, params):
        self.sectorSize = sectorSize
        self.sectorIDs = []
        self.bytes = bytes
        self.array = []
        self.params = params


    def getFreeChainEntries (self, numNeeded, startIndex ):
        freeSATChain = []
        maxEntries = len( self.array )
        for i in xrange( startIndex, maxEntries ):
            if numNeeded == 0:
                break
            nSecID = self.array[i];
            if nSecID == -1: 
                #have we exceeded the allocated size of the SAT table ?
                if i < maxEntries:
                    freeSATChain.append( i )
                    numNeeded -= 1
                else:
                    break;
        return freeSATChain        

    def getSectorSize (self):
        return self.sectorSize


    def addSector (self, id):
        self.sectorIDs.append(id)


    def buildArray (self):
        if len(self.array) > 0:
            # array already built.
            return

        numItems = int(self.sectorSize/4)
        self.array = []
        self.appendArray( self.sectorIDs )

    def appendArray( self, sectorIDs ):
        numItems = self.sectorSize / 4
        for secID in sectorIDs:
            pos = 512 + secID*self.sectorSize
            for i in xrange(0, numItems):
                beginPos = pos + i*4
                id = getSignedInt(self.bytes[beginPos:beginPos+4])
                self.array.append(id)

    def write (self):
        #writes the contents of the SAT array to memory sectors
        for index in xrange(0, len( self.array )):
            entryPos = 4 * index
            #calculate the offset into a sector
            sectorOffset = entryPos %  self.sectorSize
            sectorIndex = int(  entryPos /  self.sectorSize )
            sectorSize = self.sectorSize
            pos = 512 + ( self.sectorIDs[ sectorIndex ] * self.sectorSize ) + sectorOffset
            print "index %d, sectorOffset %d, sectorIndex %d, len(self.sectorIDs) %d sectorID %d pos %d value %d"%(index , sectorOffset , sectorIndex , len(self.sectorIDs),self.sectorIDs[ sectorIndex ], pos, self.array[ index ] )
            self.bytes[pos:pos+4] =  struct.pack( '<l', self.array[ index ] )

    def freeChainEntries (self, chain):
        for index in chain:
            self.array[ index ] = -1

    def outputRawBytes (self):
        bytes = ""
        for secID in self.sectorIDs:
            pos = 512 + secID*self.sectorSize
            bytes += self.bytes[pos:pos+self.sectorSize]
        globals.dumpBytes(bytes, 512)


    def outputArrayStats (self):
        sectorTotal = len(self.array)
        sectorP  = 0       # >= 0
        sectorM1 = 0       # -1
        sectorM2 = 0       # -2
        sectorM3 = 0       # -3
        sectorM4 = 0       # -4
        sectorMElse = 0    # < -4
        sectorLiveTotal = 0
        for i in xrange(0, len(self.array)):
            item = self.array[i]
            if item >= 0:
                sectorP += 1
            elif item == -1:
                sectorM1 += 1
            elif item == -2:
                sectorM2 += 1
            elif item == -3:
                sectorM3 += 1
            elif item == -4:
                sectorM4 += 1
            elif item < -4:
                sectorMElse += 1
            else:
                sectorLiveTotal += 1
        print("total sector count:          %4d"%sectorTotal)
        print("* live sector count:         %4d"%sectorP)
        print("* end-of-chain sector count: %4d"%sectorM2)  # end-of-chain is also live

        print("* free sector count:         %4d"%sectorM1)
        print("* SAT sector count:          %4d"%sectorM3)
        print("* MSAT sector count:         %4d"%sectorM4)
        print("* other sector count:        %4d"%sectorMElse)


    def output (self):
        print('')
        print("="*globals.OutputWidth)
        print("Sector Allocation Table (SAT)")
        print("-"*globals.OutputWidth)
        if self.params.debug:
            self.outputRawBytes()
            print("-"*globals.OutputWidth)
            for i in xrange(0, len(self.array)):
                print("%5d: %5d"%(i, self.array[i]))
            print("-"*globals.OutputWidth)

        self.outputArrayStats()


    def getSectorIDChain (self, initID):
        if initID < 0:
            return []
        chain = [initID]
        nextID = self.array[initID]
        while nextID != -2:
            chain.append(nextID)
            nextID = self.array[nextID]
        return chain


class SSAT(SAT):
    """Short Sector Allocation Table (SSAT)

SSAT contains an array of sector ID chains of all short streams, as oppposed 
to SAT which contains an array of sector ID chains of all standard streams.
The sector IDs included in the SSAT point to the short sectors in the short
stream container stream.

The first sector ID of SSAT is in the header, and the IDs of the remaining 
sectors are contained in the SAT as a sector ID chain.
"""

    def output (self):
        print('')
        print("="*globals.OutputWidth)
        print("Short Sector Allocation Table (SSAT)")
        print("-"*globals.OutputWidth)
        if self.params.debug:
            self.outputRawBytes()
            print("-"*globals.OutputWidth)
            for i in xrange(0, len(self.array)):
                item = self.array[i]
                output("%3d : %3d\n"%(i, item))

        self.outputArrayStats()


class Directory(object):
    """Directory Entries

This stream contains a list of directory entries that are stored within the
entire file stream.
"""

    class Type:
        Empty = 0
        UserStorage = 1
        UserStream = 2
        LockBytes = 3
        Property = 4
        RootStorage = 5

    class NodeColor:
        Red = 0
        Black = 1
        Unknown = 99
        
    class Entry:
        def __init__ (self):
            self.Name = ''
            self.CharBufferSize = 0
            self.Type = Directory.Type.Empty
            self.NodeColor = Directory.NodeColor.Unknown
            self.DirIDLeft = -1
            self.DirIDRight = -1
            self.DirIDRoot = -1
            self.UniqueID = None
            self.UserFlags = None
            self.TimeCreated = None
            self.TimeModified = None
            self.StreamSectorID = -2
            self.StreamSize = 0
            self.bytes = []

        def isStorage (self):
            return self.Type == Directory.Type.RootStorage or \
                self.Type == Directory.Type.UserStorage

    def __init__ (self, header, params):
        self.sectorSize = header.getSectorSize()
        self.bytes = header.bytes
        self.minStreamSize = header.minStreamSize
        self.sectorIDs = []
        self.entries = []
        self.SAT = header.getSAT()
        self.SSAT = header.getSSAT()
        self.header = header
        self.RootStorage = None
        self.RootStorageBytes = ""
        self.params = params

    def __buildRootStorageBytes (self):
        if self.RootStorage == None:
            # no root storage exists.
            return

        firstSecID = self.RootStorage.StreamSectorID
        chain = self.header.getSAT().getSectorIDChain(firstSecID)
        for secID in chain:
            pos = 512 + secID*self.sectorSize
            self.RootStorageBytes += self.header.bytes[pos:pos+self.sectorSize]

    def __getRawStream (self, entry):
        chain = []
        if entry.StreamLocation == StreamLocation.SAT:
            chain = self.header.getSAT().getSectorIDChain(entry.StreamSectorID)
        elif entry.StreamLocation == StreamLocation.SSAT:
            chain = self.header.getSSAT().getSectorIDChain(entry.StreamSectorID)

        if entry.StreamLocation == StreamLocation.SSAT:
            # Get the root storage stream.
            if self.RootStorage == None:
                raise NoRootStorage

            bytes = ""
            self.__buildRootStorageBytes()
            size = self.header.getShortSectorSize()
            for id in chain:
                pos = id*size
                bytes += self.RootStorageBytes[pos:pos+size]
            return bytes

        offset = 512
        size = self.header.getSectorSize()
        bytes = ""
        for id in chain:
            pos = offset + id*size
            bytes += self.header.bytes[pos:pos+size]

        return bytes

    def writeToStdSectors (self, targetChain, bytes ):
        srcPos = 0
        secSize = self.header.getSectorSize()
        numSectors = len(targetChain)
        print "writing out %d sectors"%(numSectors)
        for i in xrange(0, numSectors ):
            srcPos = ( i * secSize )
            targetPos = 512 + ( targetChain[ i ] * secSize )
            #bytearray behaves strangely if we try and assign a slice
            #not as big as we are addressing     
            if i == numSectors - 1:
                endPos = len(bytes) - ( secSize * ( numSectors  - 1 ))
                print "sector %d writing %d bytes as pos %d srcpos %d"%(i, endPos , targetPos, srcPos )
                self.bytes[ targetPos : targetPos + endPos ] = bytes[ srcPos : srcPos + endPos ]
            else:
                print "sector %d writing %d bytes as pos %d srcpos %d"%(i, secSize , targetPos, srcPos )
                self.bytes[ targetPos : targetPos + secSize ] = bytes[ srcPos : srcPos + secSize ]
            
    def writeToShortSectors (self, targetChain, bytes ):
        #targetChain is a short sector chain
        bytesSize = len(bytes)
        rootChain = self.header.getSAT().getSectorIDChain( self.RootStorage.StreamSectorID )
        shortSectorSize = self.header.getShortSectorSize()
        sectorSize = self.header.getSectorSize()
        for entry in targetChain:
            shortSecPos = entry * shortSectorSize
            #calculate the offset into a sector
            rootSectorOffset = shortSecPos %  sectorSize
            rootSectorIndex = int(  shortSecPos /  sectorSize )            
            targetPos = 512 + ( rootChain[ rootSectorIndex ] * sectorSize ) + rootSectorOffset
            # FIXME bytes ( from the source file ) probably isn't a full short
            # sector in size, we need to make sure we only write what we have
            # e.g. target[x:y] = src[ x:y ] might not give the expected result
            # otherwise, e.g. can get truncated etc.
            if ( bytesSize - shortSecPos < shortSectorSize ):
                endPos = bytesSize - shortSecPos
                self.header.bytes[targetPos : targetPos +  endPos ] = bytes[ shortSecPos : shortSecPos + endPos ]
            else:
                self.header.bytes[targetPos : targetPos +  shortSecPos ] = bytes[ shortSecPos : shortSecPos + shortSectorSize ]

    def ensureRootStorageCapacity(self, size ):
        chain = self.header.getSAT().getSectorIDChain( self.RootStorage.StreamSectorID )
        capacity = len(chain) *  self.header.getSAT().getSectorSize()
        sectorsNeeded = int( size / self.header.getSAT().getSectorSize() ) + 1
        if ( size >= capacity ):
            # get a new SAT sector
            newSATSectors = self.header.getSAT().getFreeChainEntries( sectorsNeeded - len(chain) , 0 )
            self.header.createSATSectors( newSATSectors, False ) 
            #rechain RootStorage SATChain
            lastIndex =  chain[ len( chain ) - 1 ]
            for entry in newSATSectors:
                self.header.getSAT().array[ lastIndex ] = entry
                lastIndex = entry
            # terminate the chain
            self.header.getSAT().array[ newSATSectors[ len( newSATSectors ) - 1 ] ] = -2           
            newchain = self.header.getSAT().getSectorIDChain( self.RootStorage.StreamSectorID )
            print "ensureRootStorageCapacity ole chain -", newchain
             
            return True            
        return True

    def getRawStream (self, entry):
        bytes = self.__getRawStream(entry)
        return bytes

    def getRawStreamByName (self, name):
        bytes = []
        for entry in self.entries:
            if entry.Name == name:
                bytes = self.__getRawStream(entry)
                break
        return bytes


    def addSector (self, id):
        self.sectorIDs.append(id)


    def output (self, debug=False):
        print('')
        print("="*globals.OutputWidth)
        print("Directory")

        if debug:
            print("-"*globals.OutputWidth)
            print("sector(s) used:")
            for secID in self.sectorIDs:
                print("  sector %d"%secID)

            print("")
            for secID in self.sectorIDs:
                print("-"*globals.OutputWidth)
                print("  Raw Hex Dump (sector %d)"%secID)
                print("-"*globals.OutputWidth)
                pos = globals.getSectorPos(secID, self.sectorSize)
                globals.dumpBytes(self.bytes[pos:pos+self.sectorSize], 128)
        dirID = 0
        for entry in self.entries:
            print("DirId[%d]"%dirID)
            dirID += 1
            self.__outputEntry(entry, debug)

    def __outputEntry (self, entry, debug):
        print("-"*globals.OutputWidth)
        if len(entry.Name) > 0:
            name = entry.Name
            if ord(name[0]) <= 5:
                name = "<%2.2Xh>%s"%(ord(name[0]), name[1:])
            print("name: %s   (name buffer size: %d bytes)"%(name, entry.CharBufferSize))
        else:
            print("name: [empty]   (name buffer size: %d bytes)"%entry.CharBufferSize)

        if self.params.debug:
            print("-"*globals.OutputWidth)
            globals.dumpBytes(entry.bytes.decode('windows-1252'))
            print("-"*globals.OutputWidth)

        output("type: ")
        if entry.Type == Directory.Type.Empty:
            print("empty")
        elif entry.Type == Directory.Type.LockBytes:
            print("lock bytes")
        elif entry.Type == Directory.Type.Property:
            print("property")
        elif entry.Type == Directory.Type.RootStorage:
            print("root storage")
        elif entry.Type == Directory.Type.UserStorage:
            print("user storage")
        elif entry.Type == Directory.Type.UserStream:
            print("user stream")
        else:
            print("[unknown type]")

        output("node color: ")
        if entry.NodeColor == Directory.NodeColor.Red:
            print("red")
        elif entry.NodeColor == Directory.NodeColor.Black:
            print("black")
        elif entry.NodeColor == Directory.NodeColor.Unknown:
            print("[unknown color]")

        print("linked dir entries: left: %d; right: %d; root: %d"%
              (entry.DirIDLeft, entry.DirIDRight, entry.DirIDRoot))

        self.__outputRaw("unique ID",  entry.UniqueID)
        self.__outputRaw("user flags", entry.UserFlags)
        self.__outputRaw("time created", entry.TimeCreated)
        self.__outputRaw("time last modified", entry.TimeModified)

        output("stream info: ")
        if entry.StreamSectorID < 0 or entry.StreamSize == 0:
            print("[empty stream]")
        else:
            strmLoc = "SAT"
            if entry.StreamLocation == StreamLocation.SSAT:
                strmLoc = "SSAT"
            print("(first sector ID: %d; size: %d; location: %s)"%
                  (entry.StreamSectorID, entry.StreamSize, strmLoc))
    
            satObj = None
            secSize = 0
            if entry.StreamLocation == StreamLocation.SAT:
                satObj = self.SAT
                secSize = self.header.getSectorSize()
            elif entry.StreamLocation == StreamLocation.SSAT:
                satObj = self.SSAT
                secSize = self.header.getShortSectorSize()
            if satObj != None:
                chain = satObj.getSectorIDChain(entry.StreamSectorID)
                print("sector count: %d"%len(chain))
                print("total sector size: %d"%(len(chain)*secSize))
                if self.params.showSectorChain:
                    self.__outputSectorChain(chain)


    def __outputSectorChain (self, chain):
        line = "sector chain: "
        lineLen = len(line)
        for id in chain:
            frag = "%d, "%id
            fragLen = len(frag)
            if lineLen + fragLen > 68:
                print(line)
                line = frag
                lineLen = fragLen
            else:
                line += frag
                lineLen += fragLen
        if line[-2:] == ", ":
            line = line[:-2]
            lineLen -= 2
        if lineLen > 0:
            print(line)


    def __outputRaw (self, name, bytes):
        bytes = bytes.decode('windows-1252')
        if bytes == None:
            return

        output("%s: "%name)
        for byte in bytes:
            output("%2.2X "%ord(byte))
        print("")

    def getDirectoryEntries (self):
        return self.entries

    def getDirectoryNames (self):
        names = []
        for entry in self.entries:
            names.append(entry.Name)
        return names


    def parseDirEntries (self):
        if len(self.entries):
            # directory entries already built
            return

        # combine all sectors first.
        bytes = ""
        for secID in self.sectorIDs:
            pos = globals.getSectorPos(secID, self.sectorSize)
            bytes += self.bytes[pos:pos+self.sectorSize]
        self.entries = []

        # each directory entry is exactly 128 bytes.
        numEntries = int(len(bytes)/128)
        if numEntries == 0:
            return
        for i in xrange(0, numEntries):
            pos = i*128
            self.entries.append(self.parseDirEntry(bytes[pos:pos+128]))

    def write(self):
        #write the entries directly to the sectors
        entriesPerSector = int( self.sectorSize / 128 )
        for i in xrange(0, len( self.entries )):
            sectorIndex = int( i/entriesPerSector )
            pos = globals.getSectorPos(  self.sectorIDs[ sectorIndex ], self.sectorSize)
            pos = pos + ( ( i % entriesPerSector ) * 128 )
            self.writeDirEntry( pos, self.entries[i] )

    def parseDirEntry (self, bytes):
        entry = Directory.Entry()
        entry.bytes = bytes

        entry.CharBufferSize = getSignedInt(bytes[64:66])
        name = bytes[0:entry.CharBufferSize].decode("utf-16")
        entry.Name = name[0:( entry.CharBufferSize -1 )/2 ]
        entry.Type = getSignedInt(bytes[66:67])
        entry.NodeColor = getSignedInt(bytes[67:68])

        entry.DirIDLeft  = getSignedInt(bytes[68:72])
        entry.DirIDRight = getSignedInt(bytes[72:76])
        entry.DirIDRoot  = getSignedInt(bytes[76:80])

        entry.UniqueID     = bytes[80:96]
        entry.UserFlags    = bytes[96:100]
        entry.TimeCreated  = bytes[100:108]
        entry.TimeModified = bytes[108:116]

        entry.StreamSectorID = getSignedInt(bytes[116:120])
        entry.StreamSize     = getSignedInt(bytes[120:124])
        entry.StreamLocation = StreamLocation.SAT
        if entry.Type != Directory.Type.RootStorage and \
            entry.StreamSize < self.header.minStreamSize:
            entry.StreamLocation = StreamLocation.SSAT

        if entry.Type == Directory.Type.RootStorage and entry.StreamSectorID >= 0:
            # This is an existing root storage.
            self.RootStorage = entry

        return entry
    
    def writeDirEntry(self, pos, entry):
        nameBytes = ( len( entry.Name) + 1 ) * 2
        sizeBefore = len( self.header.bytes )
        #self.bytes[pos + 0: pos + nameBytes + 1 ] = entry.Name.encode( 'utf-16' )
        for i in xrange(0,len(entry.Name) ):
            cPos = pos + i * 2
            self.bytes[cPos : cPos +2 ] = struct.pack( '<h', ord( entry.Name[i] ) )
        self.bytes[pos + 64: pos + 66] = struct.pack( '<h', ( len( entry.Name) + 1 ) *2 )
        self.bytes[pos + 66: pos + 67] = struct.pack( '<b', entry.Type )
        self.bytes[pos + 67: pos + 68]= struct.pack( '<b',entry.NodeColor )

        self.bytes[pos + 68: pos + 72] = struct.pack( '<l',entry.DirIDLeft )
        self.bytes[pos + 72: pos + 76] = struct.pack( '<l', entry.DirIDRight )
        self.bytes[pos + 76: pos + 80] = struct.pack('<l', entry.DirIDRoot )

        self.bytes[pos + 80:pos + 96] = entry.UniqueID
        self.bytes[pos + 96:pos + 100] = entry.UserFlags
        self.bytes[pos + 100:pos + 108] = entry.TimeCreated
        self.bytes[pos + 108:pos + 116] = entry.TimeModified

        self.bytes[pos + 116: pos + 120] = struct.pack('<l', entry.StreamSectorID )
        self.bytes[pos + 120: pos + 124] = struct.pack('<l', entry.StreamSize )
        sizeAfter = len( self.header.bytes  )
        if sizeBefore != sizeAfter:
            raise  Exception("Record size changed")
