#!/usr/bin/env python

import sys, os.path, optparse
sys.path.append(sys.path[0]+"/src")
import ole, xlsstream, globals

from globals import error

class XLDumper(object):

    def __init__ (self, filepath, params):
        self.filepath = filepath
        self.params = params
        self.strm = None
        self.strmData = None

    def __printDirHeader (self, dirname, byteLen):
        dirname = globals.decodeName(dirname)
        print("")
        print("="*68)
        print("%s (size: %d bytes)"%(dirname, byteLen))
        print("-"*68)

    def __parseFile (self):
        file = open(self.filepath, 'rb')
        self.strmData = globals.StreamData()
        self.strm = xlsstream.XLStream(file.read(), self.params, self.strmData)
        file.close()

    def dumpXML (self):
        self.__parseFile()
        dirnames = self.strm.getDirectoryNames()
        for dirname in dirnames:
            print (dirname)

    def dump (self):
        self.__parseFile()
        self.strm.printStreamInfo()
        self.strm.printHeader()
        self.strm.printMSAT()
        self.strm.printSAT()
        self.strm.printSSAT()
        self.strm.printDirectory()
        dirnames = self.strm.getDirectoryNames()
        for dirname in dirnames:
            if len(dirname) == 0 or dirname == 'Root Entry':
                continue

            dirstrm = self.strm.getDirectoryStreamByName(dirname)
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
            elif self.strmData.isPivotCacheStream(dirname):
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
    parser.add_option("--dump-xml", action="store_true", dest="dump_xml", default=False,
        help="dump content in XML format.")
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
    if options.dump_xml:
        dumper.dumpXML()
    else:
        dumper.dump()

if __name__ == '__main__':
    main()
