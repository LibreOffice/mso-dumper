#!/usr/bin/env python2
########################################################################
#
#  Copyright (c) 2010 Kohei Yoshida, Thorsten Behrens
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

import sys, os.path, getopt
sys.path.append(sys.path[0]+"/src")
import ole, pptstream, globals, olestream

from globals import error

def usage (exname):
    exname = os.path.basename(exname)
    msg = """Usage: %s [options] [ppt file]

Options:
  --help        displays this help message.
"""%exname
    print msg


class PPTDumper(object):

    def __init__ (self, filepath, params):
        self.filepath = filepath
        self.params = params

    def __printDirHeader (self, dirname, byteLen):
        dirname = globals.encodeName(dirname)
        print("")
        print("="*68)
        print("%s (size: %d bytes)"%(dirname, byteLen))
        print("-"*68)

    def dump (self):
        file = open(self.filepath, 'rb')
        strm = pptstream.PPTFile(file.read(), self.params)
        file.close()
        strm.printStreamInfo()
        strm.printHeader()
        strm.printDirectory()
        dirnames = strm.getDirectoryNames()
        result = True
        for dirname in dirnames:
            if len(dirname) == 0 or dirname == 'Root Entry':
                continue

            dirstrm = strm.getDirectoryStreamByName(dirname)
            self.__printDirHeader(dirname, len(dirstrm.bytes))
            if  dirname == "PowerPoint Document":
                if not self.__readSubStream(dirstrm):
                    result = False
            elif  dirname == "Current User":
                if not self.__readSubStream(dirstrm):
                    result = False
            elif  dirname == "\x05DocumentSummaryInformation":
                strm = olestream.PropertySetStream(dirstrm.bytes)
                strm.read()
            else:
                globals.dumpBytes(dirstrm.bytes, 512)
        return result

    def __readSubStream (self, strm):
        # read all records in substream
        return strm.readRecords()


def main (args):
    exname, args = args[0], args[1:]
    if len(args) < 1:
        print("takes at least one argument")
        usage(exname)
        return

    params = globals.Params()
    try:
        opts, args = getopt.getopt(args, "h", ["help", "debug", "show-sector-chain"])
        for opt, arg in opts:
            if opt in ['-h', '--help']:
                usage(exname)
                return
            elif opt in ['--debug']:
                params.debug = True
            elif opt in ['--show-sector-chain']:
                params.showSectorChain = True
            else:
                error("unknown option %s\n"%opt)
                usage()

    except getopt.GetoptError:
        error("error parsing input options\n")
        usage(exname)
        return

    dumper = PPTDumper(args[0], params)
    if not dumper.dump():
        error("FAILURE\n")


if __name__ == '__main__':
    main(sys.argv)
