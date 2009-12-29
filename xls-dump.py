#!/usr/bin/env python

import sys, os.path, optparse
sys.path.append(sys.path[0]+"/src")
import ole, xlsstream, globals

from globals import error

class XLDumper(object):

    def __init__ (self, filepath, params):
        self.filepath = filepath
        self.params = params

    def __printDirHeader (self, dirname, byteLen):
        dirname = globals.decodeName(dirname)
        print("")
        print("="*68)
        print("%s (size: %d bytes)"%(dirname, byteLen))
        print("-"*68)

    def dump (self):
        file = open(self.filepath, 'rb')
        strmData = globals.StreamData()
        strm = xlsstream.XLStream(file.read(), self.params, strmData)
        file.close()
        strm.printStreamInfo()
        strm.printHeader()
        strm.printMSAT()
        strm.printSAT()
        strm.printSSAT()
        strm.printDirectory()
        dirnames = strm.getDirectoryNames()
        for dirname in dirnames:
            if len(dirname) == 0 or dirname == 'Root Entry':
                continue

            dirstrm = strm.getDirectoryStreamByName(dirname)
            self.__printDirHeader(dirname, len(dirstrm.bytes))
            if dirname == "Workbook":
                success = True
                while success: 
                    success = self.__readSubStream(dirstrm)

            elif dirname == "Revision Log":
                dirstrm.type = xlsstream.DirType.RevisionLog
                self.__readSubStream(dirstrm)
            elif dirname == '_SX_DB_CUR':
                dirstrm.type = xlsstream.DirType.PivotTableCache
                self.__readSubStream(dirstrm)
            elif strmData.isPivotCacheStream(dirname):
                dirstrm.type = xlsstream.DirType.PivotTableCache
                self.__readSubStream(dirstrm)
            else:
                globals.dumpBytes(dirstrm.bytes, 512)

    def __readSubStream (self, strm):
        try:
            # read bytes from BOF to EOF.
            header = 0x0000
            while header != 0x000A:
                header = strm.readRecord()
            return True
        except xlsstream.EndOfStream:
            return False


def main ():
    parser = optparse.OptionParser()
    parser.add_option("-d", "--debug", action="store_true", dest="debug", default=False,
        help="turn on debug mode")
    parser.add_option("--show-sector-chain", action="store_true", dest="show_sector_chain", default=False,
        help="show sector chain information at the start of the output.")
    parser.add_option("--show-stream-pos", action="store_true", dest="show_stream_pos", default=False,
        help="show the position of each record relative to the stream.")
    options, args = parser.parse_args()
    params = globals.Params()
    params.debug = options.debug
    params.showSectorChain = options.show_sector_chain
    params.showStreamPos = options.show_stream_pos
    if len(args) < 1:
        globals.error("takes at least one argument\n")
        parser.print_help()
        return

    dumper = XLDumper(args[0], params)
    dumper.dump()

if __name__ == '__main__':
    main()
