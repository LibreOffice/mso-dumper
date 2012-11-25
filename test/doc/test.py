#!/usr/bin/env python
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
#

import sys
sys.path.append(sys.path[0]+"/../..")
sys.path.append(sys.path[0]+"/../../src")
doc_dumper = __import__('doc-dump')
from xml.etree import ElementTree
import unittest
import os

class Test(unittest.TestCase):
    def dump(self, name):
        try:
            os.unlink("%s.doc.xml" % name)
        except OSError:
            pass
        sock = open("%s.doc.xml" % name, "w")
        saved = sys.stdout
        sys.stdout = sock
        doc_dumper.main(["doc-dumper", "%s.doc" % name])
        sys.stdout = saved
        sock.close()
        tree = ElementTree.parse('%s.doc.xml' % name)
        return tree.getroot()

    def test_charprops(self):
        root = self.dump('charprops')

        runs = root.findall('./stream[@name="WordDocument"]/fib/fibRgFcLcbBlob/lcbPlcfBteChpx/plcBteChpx/aFC/aPnBteChpx/chpxFkp/rgfc')
        self.assertEqual(2, len(runs))

        self.assertEqual('Hello ', runs[0].findall('./transformed')[0].attrib['value'])
        self.assertEqual(0, len(runs[0].findall("./chpx/prl/sprm[@name='sprmCFBold']")))

        self.assertEqual('world!\\x0D', runs[1].findall('./transformed')[0].attrib['value'])
        self.assertEqual(1, len(runs[1].findall("./chpx/prl/sprm[@name='sprmCFBold']")))

if __name__ == '__main__':
    unittest.main()

# vim:set filetype=python shiftwidth=4 softtabstop=4 expandtab:
