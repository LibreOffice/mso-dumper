#!/usr/bin/env python3
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
#

from msodumper import swlaycacherecord
import sys


# Dumps the Star Writer binary layout cache format.
class SwLayCacheDumper:
    def __init__(self, filepath):
        self.filepath = filepath

    def dump(self):
        file = open(self.filepath, 'rb')
        strm = swlaycacherecord.SwLayCacheStream(file.read())
        file.close()
        print('<?xml version="1.0"?>')
        strm.dump()


def main(args):
    dumper = SwLayCacheDumper(args[1])
    dumper.dump()


if __name__ == '__main__':
    main(sys.argv)

# vim:set filetype=python shiftwidth=4 softtabstop=4 expandtab:
