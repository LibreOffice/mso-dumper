#!/usr/bin/env python2

import sys, os.path, optparse
sys.path.append(sys.path[0]+"/src")
import ole, globals, node, olestream

class VBADumper(object):
    
    def __init__ (self,filepath):
        print("init - file to dump from %s"%filepath)
        self.filepath = filepath
    def __parseHeader(self,chars,params):    
        self.header = ole.Header(chars, params) 
        self.header.parse()
    def __getDirectoryObj(self):
        obj = self.header.getDirectory()
        if  obj != None:
            obj.parseDirEntries()
        return obj
    def parse(self):
        params = globals.Params()
        file = open(self.filepath, 'rb')
        self.__parseHeader(file.read(), params )
        self.obj = self.__getDirectoryObj()
#        obj.output()
         
def main():
    if ( len ( sys.argv ) <= 1 ):
        print("usage: vbadump: file")
        sys.exit(1) 
    dumper = VBADumper( sys.argv[1] )
    dumper.parse()
    exit(0)

if __name__ == '__main__':
    main()
