#!/usr/bin/env python3
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
#

from msodumper import globals, vsdstream
import sys
sys = reload(sys)
sys.setdefaultencoding("utf-8")


class VSDDumper:
    def __init__(self, filepath, params):
        self.filepath = filepath
        self.params = params

    def dump(self):
        file = open(self.filepath, 'rb')
        strm = vsdstream.createVSDFile(file.read(), self.params)
        file.close()
        dirnames = strm.getDirectoryNames()
        print('<?xml version="1.0"?>\n<streams ole-type="%s">' % strm.getName())
        if strm.error:
            print('<error what="%s"/>' % strm.error)
        for dirname in dirnames:
            strm.getDirectoryStreamByName(dirname).dump()
        print('</streams>')


def main(args):
    params = globals.Params()
    dumper = VSDDumper(args[1], params)
    dumper.dump()


if __name__ == '__main__':
    main(sys.argv)

# vim:set filetype=python shiftwidth=4 softtabstop=4 expandtab:
