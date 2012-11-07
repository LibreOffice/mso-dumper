#!/usr/bin/env python

import globals

class DOCDirStream:
    """Represents one single word file subdirectory, like e.g. 'WordDocument'."""

    def __init__(self, bytes, params = None, name = None, mainStream = None, doc = None):
        self.bytes = bytes
        self.size = len(self.bytes)
        self.pos = 0
        self.params = params
        self.name = name
        self.mainStream = mainStream
        self.doc = doc
    
    def printAndSet(self, key, value, hexdump = True, end = True):
        setattr(self, key, value)
        if hexdump:
            value = hex(value)
        if end:
            print '<%s value="%s"/>' % (key, value)
        else:
            print '<%s value="%s">' % (key, value)

    def getBit(self, byte, bitNumber):
        return (byte & (1 << bitNumber)) >> bitNumber

    def dump(self):
        print '<stream name="%s" size="%s"/>' % (globals.encodeName(self.name), self.size)

# vim:set filetype=python shiftwidth=4 softtabstop=4 expandtab:
