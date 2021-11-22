#!/usr/bin/env python3
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
#

from msodumper import wmfrecord, globals
import sys
import optparse

class WMFDumper:
    def __init__(self, filepath):
        self.filepath = filepath

    def dump(self):
        file = open(self.filepath, 'rb')
        strm = wmfrecord.WMFStream(file.read())
        file.close()
        print('<?xml version="1.0"?>')
        strm.dump()


def main():
    parser = optparse.OptionParser()
    options, args = parser.parse_args()

    if len(args) < 1:
        globals.error("takes at least one argument\n")
        parser.print_help()
        sys.exit(1)

    dumper = WMFDumper(args[0])
    dumper.dump()


if __name__ == '__main__':
    main()

# vim:set filetype=python shiftwidth=4 softtabstop=4 expandtab:
