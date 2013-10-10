#!/usr/bin/env python2

import sys, os.path, optparse
sys.path.append(sys.path[0]+"/src")
import vbahelper

def main():
    if ( len ( sys.argv ) <= 1 ):
        print("usage: compress: file [offset]")
        sys.exit(1) 

    offset = 0
    if len ( sys.argv ) > 2:
        offset = int(sys.argv[2])

    file = open(sys.argv[1],'rb')
    chars = file.read()
    file.close()

    decompressed = vbahelper.UnCompressedVBAStream( chars, offset )
    compressed = decompressed.compress()

    out = open( sys.argv[1] + ".compressed", 'wb' );
    out.write( compressed );
    out.flush()
    out.close()

    exit(0)

if __name__ == '__main__':
    main()
