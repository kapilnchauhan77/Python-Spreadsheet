#############################################################################
##
# Copyright (C) 2012 Hans-Peter Jansen <hpj@urpla.net>.
# Copyright (C) 2011 Nokia Corporation and/or its subsidiary(-ies).
# All rights reserved.
# Contact: Nokia Corporation (qt-info@nokia.com)
##
# This file is part of the examples of PyQt.
##
# $QT_BEGIN_LICENSE:LGPL$
# GNU Lesser General Public License Usage
# This file may be used under the terms of the GNU Lesser General Public
# License version 2.1 as published by the Free Software Foundation and
# appearing in the file LICENSE.LGPL included in the packaging of this
# file. Please review the following information to ensure the GNU Lesser
# General Public License version 2.1 requirements will be met:
# http:#www.gnu.org/licenses/old-licenses/lgpl-2.1.html.
##
# In addition, as a special exception, Nokia gives you certain additional
# rights. These rights are described in the Nokia Qt LGPL Exception
# version 1.1, included in the file LGPL_EXCEPTION.txt in this package.
##
# GNU General Public License Usage
# Alternatively, this file may be used under the terms of the GNU General
# Public License version 3.0 as published by the Free Software Foundation
# and appearing in the file LICENSE.GPL included in the packaging of this
# file. Please review the following information to ensure the GNU General
# Public License version 3.0 requirements will be met:
# http:#www.gnu.org/copyleft/gpl.html.
##
# Other Usage
# Alternatively, this file may be used in accordance with the terms and
# conditions contained in a signed written agreement between you and Nokia.
# $QT_END_LICENSE$
##
#############################################################################


def something(a):
    if (65 + a) > 90:
        b = a + 1
        while b > 26:
            b -= 26
        c = (b) % 66

        if chr(64 + c) == "Z":
            return (((a + 1) // 26))*(chr(64 + c))
        else:
            return (((a + 1) // 26) + 1)*(chr(64 + c))
    else:
        return chr(65 + a)


def decode_pos(pos):
    try:
        row = int(pos[1:]) - 1
        col = ord(pos[0]) - ord('A')
    except ValueError:
        row = -1
        col = -1

    return row, col


def encode_pos(row, col):
    return something((col//65)+(col % 65)) + str(row + 1)
