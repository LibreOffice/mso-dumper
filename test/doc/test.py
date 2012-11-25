#!/usr/bin/env python
# -*- encoding: UTF-8 -*-
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
        self.root = tree.getroot()

    def getFontId(self, name):
        fonts = self.root.findall('stream[@name="WordDocument"]/fib/fibRgFcLcbBlob/lcbSttbfFfn/sttbfFfn/cchData')
        for i in fonts:
            if len (i.findall('ffn/xszFfn[@value="%s"]' % name)) == 1:
                return int(i.attrib['index'])

    def getRuns(self):
        return self.root.findall('stream[@name="WordDocument"]/fib/fibRgFcLcbBlob/lcbPlcfBteChpx/plcBteChpx/aFC/aPnBteChpx/chpxFkp/rgfc')

    def getParas(self):
        return self.root.findall('stream[@name="WordDocument"]/fib/fibRgFcLcbBlob/lcbPlcfBtePapx/plcBtePapx/aFC/aPnBtePapx/papxFkp/rgfc')

    def getHex(self, num):
        return int(num, 16)

    def test_hello(self):
        self.dump('hello')

        runs = self.getRuns()
        self.assertEqual(1, len(runs))

        self.assertEqual('Hello world!\\x0D', runs[0].findall('transformed')[0].attrib['value'])

    def test_unicode(self):
        self.dump('unicode')

        uni = self.root.findall('stream[@name="WordDocument"]/fib/fibRgFcLcbBlob/lcbClx/clx/pcdt/plcPcd/aCP/transformed')
        self.assertEqual('Hello world! éáőű\\x0D', uni[0].attrib['value'])

    def test_charprops(self):
        self.dump('charprops')

        runs = self.getRuns()
        self.assertEqual(2, len(runs))

        self.assertEqual('Hello ', runs[0].findall('transformed')[0].attrib['value'])
        self.assertEqual(0, len(runs[0].findall('chpx/prl/sprm[@name="sprmCFBold"]')))

        self.assertEqual('world!\\x0D', runs[1].findall('transformed')[0].attrib['value'])
        self.assertEqual(1, len(runs[1].findall('chpx/prl/sprm[@name="sprmCFBold"]')))

    def test_fonts(self):
        self.dump('fonts')
        runs = self.getRuns()
        self.assertEqual(2, len(runs))

        self.assertEqual('This is Times New Roman ', runs[0].findall('transformed')[0].attrib['value'])
        self.assertEqual(self.getFontId("Times New Roman"), self.getHex(runs[0].findall('chpx/prl/sprm[@name="sprmCRgFtc0"]')[0].attrib['operand']))

        self.assertEqual('this is arial.\\x0D', runs[1].findall('transformed')[0].attrib['value'])
        self.assertEqual(self.getFontId("Arial"), self.getHex(runs[1].findall('chpx/prl/sprm[@name="sprmCRgFtc0"]')[0].attrib['operand']))

    def test_parprops(self):
        self.dump('parprops')
        paras = self.getParas()
        self.assertEqual(2, len(paras))

        self.assertEqual('Second para.\\x0D', paras[1].findall('transformed')[0].attrib['value'])
        self.assertEqual('0x1', paras[1].findall('bxPap/papxInFkp/grpPrlAndIstd/prl/sprm[@name="sprmPJc"]')[0].attrib['operand'])

    def test_parstyles(self):
        self.dump('parstyles')
        styles = self.root.findall('stream[@name="WordDocument"]/fib/fibRgFcLcbBlob/lcbStshf/stsh/rglpstd')
        self.assertEqual('Center', styles[15].findall('std/xstz/xst/rgtchar')[0].attrib['value'])

        paras = self.getParas()
        self.assertEqual(15, self.getHex(paras[1].findall('bxPap/papxInFkp/grpPrlAndIstd/istd')[0].attrib['value']))

    def test_comment(self):
        self.dump('comment')
        comments = self.root.findall('stream[@name="WordDocument"]/fib/fibRgFcLcbBlob/lcbPlcfandTxt/plcfandTxt/aCP')
        self.assertEqual(2, len(comments))
        
        self.assertEqual('This is a comment.\\x0D', comments[0].findall('transformed')[0].attrib['value'])
        self.assertEqual('This is also commented.\\x0D', comments[1].findall('transformed')[0].attrib['value'])

        commentStarts = self.root.findall('stream[@name="WordDocument"]/fib/fibRgFcLcbBlob/lcbPlcfAtnBkf/plcfBkf/aCP')
        commentEnds = self.root.findall('stream[@name="WordDocument"]/fib/fibRgFcLcbBlob/lcbPlcfAtnBkl/plcfBkl/aCP')

        # The first comment covers Hello\x05, the second covers This\x05.
        self.assertEqual('H', commentStarts[0].findall('transformed')[0].attrib['value'])
        self.assertEqual('\\x05', commentEnds[0].findall('transformed')[0].attrib['value'])
        self.assertEqual('T', commentStarts[1].findall('transformed')[0].attrib['value'])
        self.assertEqual('\\x05', commentEnds[1].findall('transformed')[0].attrib['value'])

if __name__ == '__main__':
    unittest.main()

# vim:set filetype=python shiftwidth=4 softtabstop=4 expandtab:
