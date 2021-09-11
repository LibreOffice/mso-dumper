#!/usr/bin/env python3
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
#
# This is a simple script to convert C enum from documentation to the
# Dict that is used in the Python code
#
# Example input:
# ---------------
# typedef enum
# {
# BLACKONWHITE = 0x0001,
# WHITEONBLACK = 0x0002,
# COLORONCOLOR = 0x0003,
# HALFTONE = 0x0004
# } StretchMode;
# ---------------
#
# Example output:
# ---------------
# StretchMode = {
#     0x0001: "BLACKONWHITE",
#     0x0002: "WHITEONBLACK",
#     0x0003: "COLORONCOLOR",
#     0x0004: "HALFTONE",
# }
# ---------------

import sys

result = ""
for line in sys.stdin:
    word = line.split()
    if len(word) == 0 or word[0] == "{" or word[0] == "enum":
        continue
    elif word[0] == "}":
        name = word[1].replace(";", "")
        result = name + " = {\n" + result
    elif len(word) == 3:
        enum = word[2].replace(",", "")
        result += ("    " + enum + ": \"" + word[0] + "\",\n")
result += "}"
print(result)
