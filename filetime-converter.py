#!/usr/bin/env python

"""
    Name:   filetime_converter
    Author: Jeremy Scott
    Email:  dev@jeremyscott.org
    Purpose: Python script used to pass the Outlook Thread-Index value through the algorithm
             to decode and parse the header contents and return the WIN32 FILETIME, GUID, and
             any Child message FILETIME values.

    License: BSD

    Copyright (c) 2014, Jeremy Scott
    All rights reserved.
    Redistribution and use in source and binary forms, with or without modification, are permitted
    provided that the following conditions are met:

    Redistributions of source code must retain the above copyright notice, this list of conditions
    and the following disclaimer. Redistributions in binary form must reproduce the above copyright
    notice, this list of conditions and the following disclaimer in the documentation and/or other
    materials provided with the distribution.

    THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR
    IMPLIED WARRANTIES,INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND
    FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR
    CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
    DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
    LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF
    LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
    ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH
    DAMAGE.
"""

import sys
import argparse
from datetime import datetime,timedelta

__author__="Jeremy Scott"
__copyright__ = "Copyright (c) 2014, Jeremy Scott - jeremyscott.org"
__license__ = "BSD"
__version__ = "0.1"
__maintainer__ = __author__
__email__ = "dev [a] jeremyscott.org"
__status__ = "Development"

def convertFiletime(value):
    '''
    Definition used to parse PR_CONVERSATION_INDEX value and return human readable value.
    '''
    hex_data = value.decode('base64')
    hex_chars = map(hex,map(ord,hex_data))
    hex_string = "".join(c[2:4].zfill(2) for c in hex_chars)
    ft_value = hex_string[:12] + '0000'
    guid = hex_string[12:44]
    time_offset = int(ft_value,16) / 10.
    filetime = datetime(1601,1,1) + timedelta(microseconds=time_offset)

    print "Decoded:\t" + hex_string
    print "FILETIME:\t" + str(filetime)
    print "GUID:\t\t" + guid[:8] + "-" + guid[8:12] + "-" + guid[12:16] + "-" + guid[16:20] + "-" + guid[20:]

    if hex_string > 44:
        child_blocks = hex_string[44:]
        n=10
        children = [child_blocks[i:i+n] for i in range(0, len(child_blocks), n)]
        count = 0
        for child in children:
            scale = 16
            num_of_bits = 40
            binary = bin(int(child, scale))[2:].zfill(num_of_bits)
            time_diff = '0'*15 + binary[1:32] + '0'*18
            c_time_offset = int(time_diff, 2) / 10.
            filetime = filetime + timedelta(microseconds=c_time_offset)
            print "\tChild Message[" + str(count+1) + "]:  " + str(filetime)
            count += 1
    else:
        pass

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("value", nargs=1, help='Thread-Index value')
    parser.add_argument('--version', action='version', version='filetime_converter v' + __version__)
    args = parser.parse_args()

    value = str(args.value)
    convertFiletime(value)

if __name__ == '__main__':
    main()
