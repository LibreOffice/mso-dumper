#!/usr/bin/env python3
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
#

from msodumper import ole1previewrecord
import sys


# Dumps the OLE1 preview: see [MS-OLEDS] 2.2.5.
class Ole1PreviewDumper:
    def __init__(self, filepath):
        self.filepath = filepath

    def dump(self):
        file = open(self.filepath, 'rb')
        strm = ole1previewrecord.Ole1PreviewStream(file.read())
        file.close()
        print('<?xml version="1.0"?>')
        strm.dump()


def main(args):
    dumper = Ole1PreviewDumper(args[1])
    dumper.dump()


if __name__ == '__main__':
    main(sys.argv)

# vim:set filetype=python shiftwidth=4 softtabstop=4 expandtab:
