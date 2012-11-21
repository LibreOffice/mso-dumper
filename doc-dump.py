#!/usr/bin/env python
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
#

import sys
sys = reload(sys)
sys.setdefaultencoding("utf-8")
sys.path.append(sys.path[0]+"/src")
import globals
import docstream

class DOCDumper:
    def __init__(self, filepath, params):
        self.filepath = filepath
        self.params = params

    def dump(self):
        file = open(self.filepath, 'rb')
        strm = docstream.DOCFile(file.read(), self.params)
        file.close()
        dirnames = strm.getDirectoryNames()
        print '<?xml version="1.0"?>\n<streams>'
        for dirname in dirnames:
            if len(dirname) == 0 or dirname in ['Root Entry']:
                continue
            strm.getDirectoryStreamByName(dirname).dump()
        print '</streams>'

def main(args):
    exname, args = args[0], args[1:]
    params = globals.Params()
    dumper = DOCDumper(args[0], params)
    dumper.dump()

if __name__ == '__main__':
    main(sys.argv)

# vim:set filetype=python shiftwidth=4 softtabstop=4 expandtab:
