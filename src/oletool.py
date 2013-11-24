#!/usr/bin/env python2
########################################################################
#
#  Copyright (c) 2013 Noel Power
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


def main ():
    parser = optparse.OptionParser()
    parser.add_option("-l", "--list", action="store_true", dest="list", default=False, help="lists ole contents")
    parser.add_option("-x", "--extract", action="store_true", dest="extract", default=False, help="extract file")


    options, args = parser.parse_args()

    params = globals.Params()

    params.list =  options.list
    params.extract =  options.extract

    if len(args) < 1:
        globals.error("takes at least one argument\n")
        parser.print_help()
        sys.exit(1)

    container = ole.OleContainer( args[ 0 ], params )

    if params.list == True:
        container.list()
    if params.extract:
       files = args
       files.pop(0)

       for file in files:
           container.extract( file )

if __name__ == '__main__':
    main()
