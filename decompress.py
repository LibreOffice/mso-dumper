#!/usr/bin/env python2

import sys, os.path, optparse, struct


#The following state is maintained for a CompressedBuffer:
#CompressedRecordEnd: The location of the byte after the last byte in the CompressedContainer.
#CompressedCurrent: The location of the next byte in the CompressedContainer to be read by
#                   decompression or to be written by compression.
#The following state is maintained for the current CompressedChunk:
#CompressedChunkStart: The location of the first byte of the CompressedChunk within the
#                      CompressedContainer .
#The following state is maintained for a DecompressedBuffer:
#DecompressedCurrent: The location of the next byte in the DecompressedBuffer to be written by
#                     decompression or to be read by compression.
#DecompressedBufferEnd: The location of the byte after the last byte in the DecompressedBuffer.
#The following state is maintained for the current DecompressedChunk:
#DecompressedChunkStart: The location of the first byte of the DecompressedC hunk within the
#                        DecompressedBuffer.

class DecompressVBA(object):
    
    def __init__ (self,filepath):
        print("init - decompress from %s"%filepath)
        self.filepath = filepath
        file = open(self.filepath,'rb')
        self.chars = file.read()

    def __decompressRawChunk (self):
        chunkSize = 4096
        for i in xrange(0,chunkSize):
            self.DecompressedContainer += struct.unpack("b", self.chars[self.CompressedCurrent + i ])[0]       
        self.CompressedCurrent += chunkSize
        self.DecompressedCurrent += chunkSize

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
 
        destSize = len( self.DecompressedContainer )
        srcCurrent = srcOffSet
        dstCurrent = dstOffSet 
        for i in xrange( 0, length ):
            # check if we need to append to the Decompressed container
            print("destSize %i dstCurrent %i srcCurrent %i"%(destSize,dstCurrent,srcCurrent))
            if destSize > dstCurrent:
                self.DecompressedContainer += struct.unpack("b", self.DecompressedContainer[ srcCurrent ] )
            else: 
                self.DecompressedContainer[ dstCurrent ] = self.DecompressedContainer[ srcCurrent ]
            srcCurrent +=1
            dstCurrent +=1
                
    def __decompressToken (self, index, flagByte):
        flagBit = ( ( flagByte >> index ) & 1 )
        if flagBit == 0:
            #self.DecompressedContainer += struct.unpack("b", self.chars[self.CompressedCurrent ])[0]       
            self.DecompressedContainer +=  self.chars[self.CompressedCurrent ]
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
                self.__decompressToken(i,flagByte)
 
    def decompressCompressedChunk (self):
        print("decompressCompressedChunk")
        
        header = struct.unpack_from("<H",self.chars, self.CompressedChunkStart)[0]
        #extract size from header
        size = header & 0xFFF 
        #extract chunk sig from header
        sigflag = header >> 12
        sig = sigflag & 0x7
        #extract chunk flag from sig 
        compressedChunkFlag = (( sigflag & 0x8 ) ==  0x8)
        #print("size = %i"%size)
        #print("sig = %x"%sig)
        #print("compress = %i"%compressedChunkFlag)
        self.DecompressedChunkStart = self.DecompressedCurrent
        self.CompressedEnd = self.CompressedRecordEnd
        if ( ( self.CompressedChunkStart + size ) < self.CompressedRecordEnd ):
            self.CompressedEnd = ( self.CompressedChunkStart + size )
        self.DecompressedCurrent = self.CompressedChunkStart + 2
        if compressedChunkFlag:
            while self.CompressedCurrent < self.CompressedEnd:
                self.__decompressTokenSequence()
        else:
            self.__decompressRawChunk()
    
    def read (self):
        self.DecompressedContainer = ""
        self.CompressedCurrent = 0
        self.DecompressedCurrent = 0
        self.CompressedRecordEnd = len(self.chars )
        val = struct.unpack("b", self.chars[self.CompressedCurrent])[0]
        if val == 1:
            print("valid CompressedContainer.SignatureByte")
            self.CompressedCurrent += 1
            while self.CompressedCurrent < self.CompressedRecordEnd:
                self.CompressedChunkStart = self.CompressedCurrent
                self.decompressCompressedChunk()
        else:
            print("error decompressing containter invalid signature byte %i"% val)
         
        
def main():
    if ( len ( sys.argv ) <= 1 ):
        print("usage: decompress: file")
        sys.exit(1) 
    dumper = DecompressVBA( sys.argv[1] )
    dumper.read()
    exit(0)

if __name__ == '__main__':
    main()
