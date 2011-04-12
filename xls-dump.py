#!/usr/bin/env python
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

import sys, os.path, optparse
sys.path.append(sys.path[0]+"/src")
import ole, xlsstream, globals, node, xlsmodel, olestream

from globals import error

def isOleStream (dirname):
    """Determine whether or not a stream is an OLE stream.

Accodring to the spec, an OLE stream is always named '\1Ole'."""

    name = [0x01, 0x4F, 0x6C, 0x65] # 0x01, 'O', 'l', 'e'
    if len(dirname) != len(name):
        return False

    for i in xrange(0, len(dirname)):
        if ord(dirname[i]) != name[i]:
            return False

    return True

class XLDumper(object):

    def __init__ (self, filepath, params):
        self.filepath = filepath
        self.params = params
        self.strm = None
        self.strmData = None

    def __printDirHeader (self, direntry, byteLen):
        dirname = direntry.Name
        dirname = globals.encodeName(dirname)
        print("")
        print("="*68)
        if direntry.isStorage():
            print("%s (storage)"%dirname)
        else:
            print("%s (stream, size: %d bytes)"%(dirname, byteLen))
        print("-"*68)

    def __parseFile (self):
        file = open(self.filepath, 'rb')
        self.strmData = xlsstream.StreamData()
        self.strm = xlsstream.XLStream(file.read(), self.params, self.strmData)
        file.close()

    def dumpXML (self):
        self.__parseFile()
        dirnames = self.strm.getDirectoryNames()
        for dirname in dirnames:
            if dirname != "Workbook":
                # for now, we only dump the Workbook directory stream.
                continue

            dirstrm = self.strm.getDirectoryStreamByName(dirname)
            self.__readSubStreamXML(dirstrm)

    def dumpCanonicalXML (self):
        self.__parseFile()
        dirnames = self.strm.getDirectoryNames()
        docroot = node.Root()
        root = docroot.appendElement('xls-dump')

        for dirname in dirnames:
            if dirname != "Workbook":
                # for now, we only dump the Workbook directory stream.
                continue

            dirstrm = self.strm.getDirectoryStreamByName(dirname)
            wbmodel = self.__buildWorkbookModel(dirstrm)
            wbmodel.encrypted = self.strmData.encrypted
            root.appendChild(wbmodel.createDOM())
        
        node.prettyPrint(sys.stdout, docroot)

    def dump (self):
        self.__parseFile()
        self.strm.printStreamInfo()
        self.strm.printHeader()
        self.strm.printMSAT()
        self.strm.printSAT()
        self.strm.printSSAT()
        self.strm.printDirectory()
        dirEntries = self.strm.getDirectoryEntries()
        for entry in dirEntries:
            dirname = entry.Name
            if len(dirname) == 0:
                continue

            dirstrm = self.strm.getDirectoryStream(entry)
            self.__printDirHeader(entry, len(dirstrm.bytes))
            if entry.isStorage():
                continue

            elif dirname == "Workbook":
                success = True
                while success: 
                    success = self.__readSubStream(dirstrm)

            elif dirname == "Revision Log":
                dirstrm.type = xlsstream.DirType.RevisionLog
                self.__readSubStream(dirstrm)

            elif self.strmData.isPivotCacheStream(dirname):
                dirstrm.type = xlsstream.DirType.PivotTableCache
                self.__readSubStream(dirstrm)
            elif isOleStream(dirname):
                self.__readOleStream(dirstrm)
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

    def __readOleStream (self, dirstrm):
        strm = olestream.OLEStream(dirstrm.bytes)
        strm.read()

    def __readSubStreamXML (self, strm):
        try:
            while True:
                strm.readRecordXML()
        except xlsstream.EndOfStream:
            pass

    def __buildWorkbookModel (self, strm):
        model = xlsmodel.Workbook()
        try:
            while True:
                strm.fillModel(model)
        except xlsstream.EndOfStream:
            pass

        return model

def main ():
    parser = optparse.OptionParser()
    parser.add_option("-d", "--debug", action="store_true", dest="debug", default=False,
        help="Turn on debug mode")
    parser.add_option("--show-sector-chain", action="store_true", dest="show_sector_chain", default=False,
        help="Show sector chain information at the start of the output.")
    parser.add_option("--show-stream-pos", action="store_true", dest="show_stream_pos", default=False,
        help="Show the position of each record relative to the stream.")
    parser.add_option("--dump-mode", dest="dump_mode", default="flat", metavar="MODE",
        help="Specify the dump mode.  Possible values are: 'flat', 'xml', or 'canonical-xml'.  The default value is 'flat'.")
    options, args = parser.parse_args()
    params = globals.Params()
    params.debug = options.debug
    params.showSectorChain = options.show_sector_chain
    params.showStreamPos = options.show_stream_pos

    if len(args) < 1:
        globals.error("takes at least one argument\n")
        parser.print_help()
        sys.exit(1)

    dumper = XLDumper(args[0], params)
    if options.dump_mode == 'flat':
        dumper.dump()
    elif options.dump_mode == 'xml':
        dumper.dumpXML()
    elif options.dump_mode == 'canonical-xml' or options.dump_mode == 'cxml':
        dumper.dumpCanonicalXML()
    else:
        error("unknown dump mode: '%s'\n"%options.dump_mode)
        parser.print_help()
        sys.exit(1)

if __name__ == '__main__':
    main()
