#!/usr/bin/env python3
# -*- encoding: UTF-8 -*-
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
#

import sys
sys.path.append(sys.path[0] + "/../..")
wmf_dumper = __import__('wmf-dump')
from xml.etree import ElementTree
import unittest
import os


class Test(unittest.TestCase):
    def dump(self, name):
        try:
            os.unlink("%s.wmf.xml" % name)
        except OSError:
            pass
        sock = open("%s.wmf.xml" % name, "w")
        saved = sys.stdout
        sys.stdout = sock
        dumper = wmf_dumper.WMFDumper("%s.wmf" % name)
        dumper.dump()
        sys.stdout = saved
        sock.close()
        tree = ElementTree.parse('%s.wmf.xml' % name)
        self.root = tree.getroot()
        # Make sure everything is dumped - so it can't happen that dump(a) == dump(b), but a != b.
        self.assertEqual(0, len(self.root.findall('todo')))

    def test_pass(self):
        """This test just makes sure that all files in the 'pass' directory are
        dumped without problems."""

        for dirname, dirnames, filenames in os.walk('pass'):
            for filename in filenames:
                if filename.endswith(".wmf"):
                    self.dump(os.path.join(dirname, filename).replace('.wmf', ''))

if __name__ == '__main__':
    unittest.main()

# vim:set filetype=python shiftwidth=4 softtabstop=4 expandtab:
