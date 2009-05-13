#!/usr/bin/env python

########################################################################
#
#    OpenOffice.org - a multi-platform office productivity suite
#
#    Author:
#      Kohei Yoshida  <kyoshida@novell.com>
#      Thorsten Behrens <tbehrens@novell.com>
#
#   The Contents of this file are made available subject to the terms
#   of GNU Lesser General Public License Version 2.1 and any later
#   version.
#
########################################################################

import sys, os.path, getopt
sys.path.append(sys.path[0]+"/src")
import ole, pptstream, globals

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
        dirname = globals.decodeName(dirname)
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
