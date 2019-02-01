#!/usr/bin/env python3
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# PowerPoint generates SmartArt streams, like ppt/diagrams/data1.xml containing
# UUIDs for each <dgm:pt> element. This makes it hard to reason about them,
# referring to numerical identifiers is just easier in notes. This script
# replaces UUIDs with integers, which is the style used in the OOXML spec as
# well.

import re
import sys


def main():
    buf = sys.stdin.read()

    counter = 0
    while True:
        match = re.search("\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}", buf)
        if not match:
            break
        uuid = match.group()
        buf = buf.replace(uuid, str(counter))
        counter += 1

    sys.stdout.write(buf)


if __name__ == "__main__":
    main()

# vim:set shiftwidth=4 softtabstop=4 expandtab:
