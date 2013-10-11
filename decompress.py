#!/usr/bin/env python2

import sys, os.path, optparse
sys.path.append(sys.path[0]+"/src")
import vbahelper

def main():

    offset = 0
    if len ( sys.argv ) > 1:
        offset = int(sys.argv[1])

    chars = sys.stdin.read()

    compressed = vbahelper.CompressedVBAStream( chars, offset )
    decompressed = compressed.decompress()
    sys.stdout.write(decompressed)

    exit(0)

if __name__ == '__main__':
    main()
