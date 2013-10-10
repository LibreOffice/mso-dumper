import  sys, struct
class VBAStreamBase:
    CHUNKSIZE = 4096
    def __init__(self, chars, offset):
        self.mnOffset = offset
        self.chars = chars

class UnCompressedVBAStream(VBAStreamBase):
    def __packCopyToken(self, offset, length ):
        lengthMask, offSetMask, bitCount, maximumLength = self.__copyTokenHelp()
        temp1 = offset - 1
        temp2 = 16 - bitCount
        temp3 = length - 3
        copyToken = (temp1 << temp2) | temp3
        return copyToken

    def __compressRawChunk(self):
        self.CompressedCurrent = self.CompressedChunkStart + 2
        self.DecompressedCurrent  = self.DecompressedChunkStart
        PadCount = self.CHUNKSIZE
        LastByte = self.DecompressedChunkStart + PadCount
        if self.DecompressedBufferEnd < LastByte:
           LastByte =  self.DecompressedBufferEnd

        for index in xrange( self.DecompressedChunkStart,  LastByte ):
            self.CompressedContainer[ self.CompressedCurrent ] = self.chars[ index ]
            self.CompressedCurrent = self.CompressedCurrent + 1
            self.DecompressedCurrent = self.DecompressedCurrent + 1
            PadCount = PadCount - 1

        for index in xrange( 0, PadCount ):
            self.CompressedContainer[ self.CompressedCurrent ] = 0x0;   
            self.CompressedCurrent = self.CompressedCurrent + 1

    # #FIXME move to base class
    def __copyTokenHelp(self):
        difference = self.DecompressedCurrent - self.DecompressedChunkStart
        bitCount = 0
        while( ( 1 << bitCount ) < difference ):
            bitCount +=1

        if bitCount < 4:
            bitCount = 4;

        lengthMask = 0xFFFF >> bitCount
        offSetMask = ~lengthMask
        maximumLength = ( 0xFFFF >> bitCount ) + 3
        return lengthMask, offSetMask, bitCount, maximumLength

    def __matching( self, DecompressedEnd ):
        Candidate = self.DecompressedCurrent - 1
        BestLength = 0
        while Candidate >= self.DecompressedChunkStart:
            C = Candidate
            D = self.DecompressedCurrent
            Len = 0
            while  D < DecompressedEnd and ( self.chars[ D ] == self.chars[ C ] ):
                Len = Len + 1
                C = C + 1
                D = D + 1
            if Len > BestLength:
                BestLength = Len
                BestCandidate = Candidate
            Candidate = Candidate - 1
        if BestLength >=  3:
            lengthMask, offSetMask, bitCount, maximumLength = self.__copyTokenHelp()
            Length = BestLength
            if ( maximumLength < BestLength ):
                Length = maximumLength 
            Offset = self.DecompressedCurrent - BestCandidate
        else:
            Length = 0
            Offset = 0
        return Offset, Length

    def __compressToken( self, CompressedEnd, DecompressedEnd, index, Flags ):
        Offset = 0
        Offset, Length = self.__matching( DecompressedEnd )
        if Offset:
            if (self.CompressedCurrent + 1) < CompressedEnd:
                copyToken = self.__packCopyToken( Offset, Length )
                struct.pack_into("<H", self.CompressedContainer, self.CompressedCurrent, copyToken )

                temp1 = ( 1 << index )
                temp2 = Flags & ~temp1
                Flags = temp2 | temp1

                self.CompressedCurrent = self.CompressedCurrent + 2
                self.DecompressedCurrent = self.DecompressedCurrent + Length
            else:
                self.CompressedCurrent = CompressedEnd
        else:
            if self.CompressedCurrent < CompressedEnd:
                self.CompressedContainer[ self.CompressedCurrent ] = self.chars[ self.DecompressedCurrent ]
                self.CompressedCurrent = self.CompressedCurrent + 1
                self.DecompressedCurrent = self.DecompressedCurrent + 1
            else:
                self.CompressedCurrent = CompressedEnd
        return Flags

    def __compressTokenSequence(self, CompressedEnd, DecompressedEnd ):
        FlagByteIndex = self.CompressedCurrent
        TokenFlags = 0
        self.CompressedCurrent = self.CompressedCurrent + 1
        for index in xrange(0,8): 
            if ( ( self.DecompressedCurrent < DecompressedEnd )
                and (self.CompressedCurrent < CompressedEnd) ):

                TokenFlags = self.__compressToken( CompressedEnd, DecompressedEnd, index, TokenFlags )
        self.CompressedContainer[ FlagByteIndex ] = TokenFlags
    def __CompressDecompressedChunk(self):
        self.CompressedContainer.extend(  bytearray(self.CHUNKSIZE + 2) )
        CompressedEnd = self.CompressedChunkStart + 4098
        self.CompressedCurrent = self.CompressedChunkStart + 2
        DecompressedEnd = self.DecompressedBufferEnd
        if  (self.DecompressedChunkStart + self.CHUNKSIZE) <  self.DecompressedBufferEnd:
            DecompressedEnd = (self.DecompressedChunkStart + self.CHUNKSIZE)

        while (self.DecompressedCurrent < DecompressedEnd) and (self.CompressedCurrent < CompressedEnd):
                self.__compressTokenSequence( CompressedEnd, DecompressedEnd)

        if self.DecompressedCurrent < DecompressedEnd:
            self.__compressRawChunk( DecompressedEnd - 1 )
            CompressedFlag = 0
        else:
            CompressedFlag = 1
        Size = self.CompressedCurrent - self.CompressedChunkStart
        Header = 0x0000
        #Pack CompressedChunkSize with Size and Header
        temp1=Header & 0xF000
        temp2 = Size - 3
        Header = temp1 | temp2
        #Pack CompressedChunkFlag with CompressedFlag and Header
        temp1 = Header & 0x7FFF
        temp2 = CompressedFlag << 15
        Header = temp1 | temp2
        #CALL Pack CompressedChunkSignature with Header
        temp1 = Header & 0x8FFF
        Header = temp1 | 0x3000
        #SET the CompressedChunkHeader located at CompressedChunkStart TO Header
        struct.pack_into("<H", self.CompressedContainer, self.CompressedChunkStart, Header )

        # trim buffer to size
        if ( self.CompressedCurrent ):
            self.CompressedContainer = self.CompressedContainer[ 0:self.CompressedCurrent ]

    def compress(self):
        self.CompressedContainer = bytearray()
        self.CompressedCurrent = 0
        self.CompressedChunkStart = 0
        self.DecompressedCurrent = 0
        self.DecompressedBufferEnd = len(self.chars)
        self.DecompressedChunkStart = 0
        SignatureByte = 0x01
        self.CompressedContainer.append( SignatureByte )
        self.CompressedCurrent = self.CompressedCurrent + 1
        while self.DecompressedCurrent < self.DecompressedBufferEnd:
            self.CompressedChunkStart = self.CompressedCurrent
            self.DecompressedChunkStart = self.DecompressedCurrent
            self.__CompressDecompressedChunk()
        return self.CompressedContainer

class CompressedVBAStream(VBAStreamBase):
    
    def __decompressRawChunk (self):
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
        flagByte = struct.unpack("b", self.chars[self.CompressedCurrent ])[0]
        self.CompressedCurrent += 1
        if  self.CompressedCurrent < self.CompressedEnd:
            for i in xrange(0,8):
                if  self.CompressedCurrent < self.CompressedEnd:
                    self.__decompressToken(i,flagByte)
 
    def decompressCompressedChunk (self):
        
        header = struct.unpack_from("<H",self.chars, self.CompressedChunkStart)[0]
        #extract size from header
        size = header & 0xFFF 
        size = size + 3
        #extract chunk sig from header
        sigflag = header >> 12
        sig = sigflag & 0x7
        #extract chunk flag from sig 
        compressedChunkFlag = (( sigflag & 0x8 ) ==  0x8)
        self.DecompressedChunk = bytearray(self.CHUNKSIZE);
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
        self.DecompressedChunkStart = 0
        val = struct.unpack("b", self.chars[self.CompressedCurrent])[0]
        if val == 1:
            self.CompressedCurrent += 1
            while self.CompressedCurrent < self.CompressedRecordEnd:
                self.CompressedChunkStart = self.CompressedCurrent
                self.decompressCompressedChunk()
            return self.DecompressedContainer
        else:
            raise Exception("error decompressing container invalid signature byte %i"% val)
         
        return None

