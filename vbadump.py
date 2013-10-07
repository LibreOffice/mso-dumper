#!/usr/bin/env python2

import sys, os.path, optparse
sys.path.append(sys.path[0]+"/src")
import ole, globals, node, olestream

class VBADumper(object):
    def __init__ (self,filepath):
        print("init - file to dump from %s"%filepath)

def main():
    if ( len ( sys.argv ) <= 1 ):
        print("usage: vbadump: file")
        sys.exit(1) 
    dumper = VBADumper( sys.argv[1] )
    exit(0)

if __name__ == '__main__':
    main()
