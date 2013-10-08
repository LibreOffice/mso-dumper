import  sys, struct

class CompressedVBAStream(object):
    
    CHUNKSIZE = 4096
    def __init__ (self,chars, offset):
#        print("init - decompress from %s"%filepath)
        self.chars = chars
        self.mnOffset = offset

    def __decompressRawChunk (self):
#        print("decompressRawChunk")
        for i in xrange(0,self.CHUNKSIZE):
            self.DecompressedChunk[ self.DecompressedCurrent + i ] =  self.chars[self.CompressedCurrent + i ]
        self.CompressedCurrent += self.CHUNKSIZE
        self.DecompressedCurrent += self.CHUNKSIZE

    def __copyTokenHelp(self):
        difference = self.DecompressedCurrent - self.DecompressedChunkStart
        bitCount = 0
        while( ( 1 << bitCount ) < difference ):
            bitCount +=1

        if bitCount < 4:
            bitCount = 4;

        lengthMask = 0xFFFF >> bitCount
        offSetMask = ~lengthMask
#        maximumLength = ( 0xFFFF >> bitCOunt ) + 3
        return lengthMask, offSetMask, bitCount

    def __unPackCopyToken (self, copyToken ):
       lengthMask, offSetMask, bitCount = self.__copyTokenHelp() 
       length = ( copyToken & lengthMask ) + 3
       temp1 = copyToken & offSetMask
       temp2 = 16 - bitCount
       offSet = ( temp1 >> temp2 ) + 1
       return offSet, length

    def __byteCopy( self, srcOffSet, dstOffSet, length ):
 
        destSize = len( self.DecompressedChunk )
        srcCurrent = srcOffSet
        dstCurrent = dstOffSet 
        for i in xrange( 0, length ):
            self.DecompressedChunk[ dstCurrent ] = self.DecompressedChunk[ srcCurrent ]
            srcCurrent +=1
            dstCurrent +=1
                
    def __decompressToken (self, index, flagByte):
        flagBit = ( ( flagByte >> index ) & 1 )
        if flagBit == 0:
            #self.DecompressedChunk += struct.unpack("b", self.chars[self.CompressedCurrent ])[0]       
            self.DecompressedChunk[self.DecompressedCurrent] = self.chars[self.CompressedCurrent ]
            self.CompressedCurrent += 1
            self.DecompressedCurrent += 1
        else:
            copyToken = struct.unpack_from("<H", self.chars, self.CompressedCurrent)[0] 
            offSet, length = self.__unPackCopyToken( copyToken )
            copySource =  self.DecompressedCurrent - offSet
            self.__byteCopy( copySource, self.DecompressedCurrent, length )
            self.DecompressedCurrent += length
            self.CompressedCurrent += 2

    def __decompressTokenSequence (self):
#        print(" __decompressTokenSequence at CompressedCurrent %i CompressedEnd %i"%( self.CompressedCurrent, self.CompressedEnd ) )
        flagByte = struct.unpack("b", self.chars[self.CompressedCurrent ])[0]
        self.CompressedCurrent += 1
        if  self.CompressedCurrent < self.CompressedEnd:
            for i in xrange(0,8):
                if  self.CompressedCurrent < self.CompressedEnd:
                    self.__decompressToken(i,flagByte)
 
    def decompressCompressedChunk (self):
#        print("decompressCompressedChunk")
        
        header = struct.unpack_from("<H",self.chars, self.CompressedChunkStart)[0]
        #extract size from header
        size = header & 0xFFF 
#        print("raw size %i"%size)
        size = size + 3
        #extract chunk sig from header
        sigflag = header >> 12
        sig = sigflag & 0x7
        #extract chunk flag from sig 
        compressedChunkFlag = (( sigflag & 0x8 ) ==  0x8)
#        print("chunksize = %i"%size)
#        print("sig = %x"%sig)
#        print("compress = %i"%compressedChunkFlag)
        self.DecompressedChunk = bytearray(self.CHUNKSIZE);  # one chunk ( need to cater for more than one chunk of course )
        self.DecompressedChunkStart = 0
        self.DecompressedCurrent = 0
        self.CompressedEnd = self.CompressedRecordEnd
        if ( ( self.CompressedChunkStart + size ) < self.CompressedRecordEnd ):
            self.CompressedEnd = ( self.CompressedChunkStart + size )
        self.CompressedCurrent = self.CompressedChunkStart + 2
        if compressedChunkFlag == 1:
            while self.CompressedCurrent < self.CompressedEnd:
                self.__decompressTokenSequence()
        else:
            self.__decompressRawChunk()
        if self.DecompressedCurrent:
             truncChunk = self.DecompressedChunk[0:self.DecompressedCurrent]
             self.DecompressedContainer.extend( truncChunk )

    def decompress (self):
        self.DecompressedContainer = bytearray();
        self.CompressedCurrent = self.mnOffset
        self.CompressedRecordEnd = len(self.chars )
        self.DeCompressedBufferEnd = 0
        self.DecompressedChunkStart = self.CompressedCurrent 
        val = struct.unpack("b", self.chars[self.CompressedCurrent])[0]
        if val == 1:
#            print("valid CompressedContainer.SignatureByte")
            self.CompressedCurrent += 1
            while self.CompressedCurrent < self.CompressedRecordEnd:
                self.CompressedChunkStart = self.CompressedCurrent
#                print("about to call decompressChunk for start %i"%self.CompressedChunkStart)
                self.decompressCompressedChunk()
            return self.DecompressedContainer
        else:
            raise Exception("error decompressing container invalid signature byte %i"% val)
         
        return None

