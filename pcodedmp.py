#!/usr/bin/env python

from __future__ import print_function
from oletools.olevba import VBA_Parser, decompress_stream
from oletools.ezhexviewer import hexdump3
from struct import *
import argparse
import sys
import os

__author__ = 'Vesselin Bontchev <vbontchev@yahoo.com>'
__license__ = 'GPL'
__VERSION__ = '1.01'

def getWord(buffer, offset, endian):
    return unpack_from(endian + 'H', buffer, offset)[0]

def getDWord(buffer, offset, endian):
    return unpack_from(endian + 'L', buffer, offset)[0]

def skipStructure(buffer, offset, endian, isLengthDW, elementSize, checkForMinusOne):
    if (isLengthDW):
        length = getDWord(buffer, offset, endian)
        offset += 4
    else:
        length = getWord(buffer, offset, endian)
        offset += 2
    if (not checkForMinusOne or (length != 0xFFFF)):
        offset += length * elementSize
    return offset

def getVar(buffer, offset, endian, isDWord):
    if(isDWord):
        value = getDWord(buffer, offset, endian)
        offset += 4
    else:
        value = getWord(buffer, offset, endian)
        offset += 2
    return offset, value

def getTypeAndLength(buffer, offset, endian):
    if (endian == '>'):
        return ord(buffer[offset]), ord(buffer[offset + 1])
    else:
        return ord(buffer[offset + 1]), ord(buffer[offset])

def processPROJECT(vbaParser, projectPath, disasmonly):
    projectData = vbaParser.ole_file.openstream(projectPath).read()
    if (not disasmonly):
        print('-' * 79)
        print('PROJECT dump:')
        print(projectData)
    return projectData

def processDir(vbaParser, dirPath, verbose, disasmonly):
    tags = {
	 1 : 'PROJ_SYSKIND',	# 0 - Win16, 1 - Win32, 2 - Mac, 3 - Win64
	 2 : 'PROJ_LCID',
	 3 : 'PROJ_CODEPAGE',
	 4 : 'PROJ_NAME',
	 5 : 'PROJ_DOCSTRING',
	 6 : 'PROJ_HELPFILE',
	 7 : 'PROJ_HELPCONTEXT',
	 8 : 'PROJ_LIBFLAGS',
	 9 : 'PROJ_VERSION',
	10 : 'PROJ_GUID',
	11 : 'PROJ_PROPERTIES',
	12 : 'PROJ_CONSTANTS',
	13 : 'PROJ_LIBID_REGISTERED',
	14 : 'PROJ_LIBID_PROJ',
	15 : 'PROJ_MODULECOUNT',
	16 : 'PROJ_EOF',
	17 : 'PROJ_TYPELIB_VERSION',
	18 : 'PROJ_COMPAT_EXE',
	19 : 'PROJ_COOKIE',
	20 : 'PROJ_LCIDINVOKE',
	21 : 'PROJ_COMMAND_LINE',
	22 : 'PROJ_REFNAME_PROJ',

	25 : 'MOD_NAME',
	26 : 'MOD_STREAM',

	28 : 'MOD_DOCSTRING',
	29 : 'MOD_HELPFILE',
	30 : 'MOD_HELPCONTEXT',

	32 : 'MOD_PROPERTIES',
	33 : 'MOD_FBASMOD_StdMods',
	34 : 'MOD_FBASMOD_Classes',
	35 : 'MOD_FBASMOD_Creatable',
	36 : 'MOD_FBASMOD_NoDisplay',
	37 : 'MOD_FBASMOD_NoEdit',
	38 : 'MOD_FBASMOD_RefLibs',
	39 : 'MOD_FBASMOD_NonBasic',
	40 : 'MOD_FBASMOD_Private',
	41 : 'MOD_FBASMOD_Internal',
	42 : 'MOD_FBASMOD_AllModTypes',
	43 : 'MOD_END',
	44 : 'MOD_COOKIETYPE',
	45 : 'MOD_BASECLASSNULL',
	46 : 'MOD_BASECLASS',
	47 : 'PROJ_LIBID_TWIDDLED',
	48 : 'PROJ_LIBID_EXTENDED',
	49 : 'MOD_TEXTOFFSET',
	50 : 'MOD_UNICODESTREAM',

	60 : 'PROJ_UNICODE_CONSTANTS',
	61 : 'PROJ_UNICODE_HELPFILE',
	62 : 'PROJ_UNICODE_REFNAME_PROJ',
	63 : 'PROJ_UNICODE_COMMAND_LINE',
	64 : 'PROJ_UNICODE_DOCSTRING',

	71 : 'MOD_UNICODE_NAME',
	72 : 'MOD_UNICODE_DOCSTRING',
	73 : 'MOD_UNICODE_HELPFILE'
    }
    if (not disasmonly):
        print('-' * 79)
        print('dir stream after decompression:')
    dirDataCompressed = vbaParser.ole_file.openstream(dirPath).read()
    dirData = decompress_stream(dirDataCompressed)
    streamSize = len(dirData)
    codeModules = []
    if (not disasmonly):
        print('%d bytes' % streamSize)
        if (verbose):
            print(hexdump3(dirData, length=16))
        print('dir stream parsed:')
    offset = 0
    # The "dir" stream is ALWAYS in little-endian format, even on a Mac
    while offset < streamSize:
        try:
            tag = getWord(dirData, offset, '<')
            wLength = getWord(dirData, offset + 2, '<')
            # The following idiocy is because Microsoft can't stick
            # to their own format specification
            if (tag == 9):
                wLength = 6
            elif (tag == 3):
                wLength = 2
            # End of the idiocy
            if (not tag in tags):
                tagName = 'UNKNOWN'
            else:
                tagName = tags[tag]
            if (not disasmonly):
                print('%08X:  %s' % (offset, tagName), end='')
            offset += 6
            if (wLength):
                if (not disasmonly):
                    print(':')
                    print(hexdump3(dirData[offset:offset + wLength], length=16))
                if (tagName == 'MOD_STREAM'):
                    codeModules.append(dirData[offset:offset + wLength])
                offset += wLength
            elif (not disasmonly):
                print('')
        except:
            break
    return dirData, codeModules

def process_VBA_PROJECT(vbaParser, vbaProjectPath, verbose, disasmonly):
    vbaProjectData = vbaParser.ole_file.openstream(vbaProjectPath).read()
    if (disasmonly):
        return vbaProjectData
    print('-' * 79)
    print('_VBA_PROJECT stream:')
    print('%d bytes' % len(vbaProjectData))
    if (verbose):
        print(hexdump3(vbaProjectData, length=16))
    return vbaProjectData

def getTheIdentifiers(vbaProjectData):
    identifiers = []
    try:
        magic = getWord(vbaProjectData, 0, '<')
        if (magic != 0x61CC):
            return identifiers
        version = getWord(vbaProjectData, 2, '<')
        unicodeRef  = (version >= 0x5B) and (not version in [0x60, 0x62, 0x63]) or (version == 0x4E)
        unicodeName = (version >= 0x59) and (not version in [0x60, 0x62, 0x63]) or (version == 0x4E)
        nonUnicodeName = ((version <= 0x59) and (version != 0x4E)) or (0x5F > version > 0x6B)
        word = getWord(vbaProjectData, 5, '<')
        if (word == 0x000E):
            endian = '>'
        else:
            endian = '<'
        offset = 0x1E
        offset, numRefs = getVar(vbaProjectData, offset, endian, False)
        offset += 2
        for ref in range(numRefs):
            offset, refLength = getVar(vbaProjectData, offset, endian, False)
            if (refLength == 0):
                offset += 6
            else:
                if ((unicodeRef and (refLength < 5)) or ((not unicodeRef) and (refLength < 3))):
                    offset += refLength
                else:
                    if (unicodeRef):
                        c = vbaProjectData[offset + 4]
                    else:
                        c = vbaProjectData[offset + 2]
                    offset += refLength
                    if (c in ['C', 'D']):
                        offset = skipStructure(vbaProjectData, offset, endian, False, 1, False)
            offset += 10
            offset, word = getVar(vbaProjectData, offset, endian, False)
            if (word):
                offset = skipStructure(vbaProjectData, offset, endian, False, 1, False)
                offset, wLength = getVar(vbaProjectData, offset, endian, False)
                if (wLength):
                    offset += 2
                offset += wLength + 30
        # Number of entries in the class/user forms table
        offset = skipStructure(vbaProjectData, offset, endian, False, 2, False)
        # Number of compile-time identifier-value pairs
        offset = skipStructure(vbaProjectData, offset, endian, False, 4, False)
        offset += 2
        # Typeinfo typeID
        offset = skipStructure(vbaProjectData, offset, endian, False, 1, True)
        # Project description
        offset = skipStructure(vbaProjectData, offset, endian, False, 1, True)
        # Project help file name
        offset = skipStructure(vbaProjectData, offset, endian, False, 1, True)
        offset += 0x64
        # Skip the module descriptors
        offset, numProjects = getVar(vbaProjectData, offset, endian, False)
        for project in range(numProjects):
            offset, wLength = getVar(vbaProjectData, offset, endian, False)
            # Code module name
            if (unicodeName):
                offset += wLength
            if (nonUnicodeName):
                if (wLength):
                    offset, wLength = getVar(vbaProjectData, offset, endian, False)
                offset += wLength
            # Stream time
            offset = skipStructure(vbaProjectData, offset, endian, False, 1, False)
            offset = skipStructure(vbaProjectData, offset, endian, False, 1, True)
            offset, streamID = getVar(vbaProjectData, offset, endian, False)
            if (version >= 0x6B):
                offset = skipStructure(vbaProjectData, offset, endian, False, 1, True)
            offset = skipStructure(vbaProjectData, offset, endian, False, 1, True)
            offset += 2
            if (version != 0x51):
                offset += 4
            offset = skipStructure(vbaProjectData, offset, endian, False, 8, False)
            offset += 11
        offset += 6
        offset = skipStructure(vbaProjectData, offset, endian, True, 1, False)
        offset += 6
        offset, w0 = getVar(vbaProjectData, offset, endian, False)
        offset, numIDs = getVar(vbaProjectData, offset, endian, False)
        offset, w1 = getVar(vbaProjectData, offset, endian, False)
        offset += 4
        numJunkIDs = numIDs + w1 - w0
        numIDs = w0 - w1
        # Skip the junk IDs
        for id in range(numJunkIDs):
            offset += 4
            idType, idLength = getTypeAndLength(vbaProjectData, offset, endian)
            offset += 2
            if (idType > 0x7F):
                offset += 6
            offset += idLength
        # Now offset points to the start of the variable names area
        for id in range(numIDs):
            isKwd = False
            idType, idLength = getTypeAndLength(vbaProjectData, offset, endian)
            offset += 2
            if ((idLength == 0) and (idType == 0)):
                offset += 2
                idType, idLength = getTypeAndLength(vbaProjectData, offset, endian)
                offset += 2
                isKwd = True
            if (idType & 0x80):
                offset += 6
            if (idLength):
                identifiers.append(vbaProjectData[offset:offset + idLength])
                offset += idLength
            if (not isKwd):
                offset += 4
    except Exception as e:
        print('Error: %s.' % e, file=sys.stderr)
    return identifiers

#'name', '0x', 'imp_', 'func_', 'var_', 'rec_', 'type_', 'context_'
#    2,     2,      2,       4,      4,      4,       4,         4

# VBA7 opcodes; VBA3, VBA5 and VBA6 will be upconverted to these.
opcodes = {
  0 : { 'mnem' : 'Imp',                   'args' : [],                       'varg' : False },
  1 : { 'mnem' : 'Eqv',                   'args' : [],                       'varg' : False },
  2 : { 'mnem' : 'Xor',                   'args' : [],                       'varg' : False },
  3 : { 'mnem' : 'Or',                    'args' : [],                       'varg' : False },
  4 : { 'mnem' : 'And',                   'args' : [],                       'varg' : False },
  5 : { 'mnem' : 'Eq',                    'args' : [],                       'varg' : False },
  6 : { 'mnem' : 'Ne',                    'args' : [],                       'varg' : False },
  7 : { 'mnem' : 'Le',                    'args' : [],                       'varg' : False },
  8 : { 'mnem' : 'Ge',                    'args' : [],                       'varg' : False },
  9 : { 'mnem' : 'Lt',                    'args' : [],                       'varg' : False },
 10 : { 'mnem' : 'Gt',                    'args' : [],                       'varg' : False },
 11 : { 'mnem' : 'Add',                   'args' : [],                       'varg' : False },
 12 : { 'mnem' : 'Sub',                   'args' : [],                       'varg' : False },
 13 : { 'mnem' : 'Mod',                   'args' : [],                       'varg' : False },
 14 : { 'mnem' : 'IDiv',                  'args' : [],                       'varg' : False },
 15 : { 'mnem' : 'Mul',                   'args' : [],                       'varg' : False },
 16 : { 'mnem' : 'Div',                   'args' : [],                       'varg' : False },
 17 : { 'mnem' : 'Concat',                'args' : [],                       'varg' : False },
 18 : { 'mnem' : 'Like',                  'args' : [],                       'varg' : False },
 19 : { 'mnem' : 'Pwr',                   'args' : [],                       'varg' : False },
 20 : { 'mnem' : 'Is',                    'args' : [],                       'varg' : False },
 21 : { 'mnem' : 'Not',                   'args' : [],                       'varg' : False },
 22 : { 'mnem' : 'UMi',                   'args' : [],                       'varg' : False },
 23 : { 'mnem' : 'FnAbs',                 'args' : [],                       'varg' : False },
 24 : { 'mnem' : 'FnFix',                 'args' : [],                       'varg' : False },
 25 : { 'mnem' : 'FnInt',                 'args' : [],                       'varg' : False },
 26 : { 'mnem' : 'FnSgn',                 'args' : [],                       'varg' : False },
 27 : { 'mnem' : 'FnLen',                 'args' : [],                       'varg' : False },
 28 : { 'mnem' : 'FnLenB',                'args' : [],                       'varg' : False },
 29 : { 'mnem' : 'Paren',                 'args' : [],                       'varg' : False },
 30 : { 'mnem' : 'Sharp',                 'args' : [],                       'varg' : False },
 31 : { 'mnem' : 'LdLHS',                 'args' : [],                       'varg' : False },
 32 : { 'mnem' : 'Ld',                    'args' : ['name'],                 'varg' : False },
 33 : { 'mnem' : 'MemLd',                 'args' : ['name'],                 'varg' : False },
 34 : { 'mnem' : 'DictLd',                'args' : ['name'],                 'varg' : False },
 35 : { 'mnem' : 'IndexLd',               'args' : ['0x'],                   'varg' : False },
 36 : { 'mnem' : 'ArgsLd',                'args' : ['name',   '0x'],         'varg' : False },
 37 : { 'mnem' : 'ArgsMemLd',             'args' : ['name',   '0x'],         'varg' : False },
 38 : { 'mnem' : 'ArgsDictLd',            'args' : ['name',   '0x'],         'varg' : False },
 39 : { 'mnem' : 'St',                    'args' : ['name'],                 'varg' : False },
 40 : { 'mnem' : 'MemSt',                 'args' : ['name'],                 'varg' : False },
 41 : { 'mnem' : 'DictSt',                'args' : ['name'],                 'varg' : False },
 42 : { 'mnem' : 'IndexSt',               'args' : ['name'],                 'varg' : False },
 43 : { 'mnem' : 'ArgsSt',                'args' : ['name',   '0x'],         'varg' : False },
 44 : { 'mnem' : 'ArgsMemSt',             'args' : ['name',   '0x'],         'varg' : False },
 45 : { 'mnem' : 'ArgsDictSt',            'args' : ['name',   '0x'],         'varg' : False },
 46 : { 'mnem' : 'set',                   'args' : ['name'],                 'varg' : False },
 47 : { 'mnem' : 'Memset',                'args' : ['name'],                 'varg' : False },
 48 : { 'mnem' : 'Dictset',               'args' : ['name'],                 'varg' : False },
 49 : { 'mnem' : 'Indexset',              'args' : ['name'],                 'varg' : False },
 50 : { 'mnem' : 'ArgsSet',               'args' : ['name',   '0x'],         'varg' : False },
 51 : { 'mnem' : 'ArgsMemSet',            'args' : ['name',   '0x'],         'varg' : False },
 52 : { 'mnem' : 'ArgsDictSet',           'args' : ['name',   '0x'],         'varg' : False },
 53 : { 'mnem' : 'MemLdWith',             'args' : ['name'],                 'varg' : False },
 54 : { 'mnem' : 'DictLdWith',            'args' : ['name'],                 'varg' : False },
 55 : { 'mnem' : 'ArgsMemLdWith',         'args' : ['name',   '0x'],         'varg' : False },
 56 : { 'mnem' : 'ArgsDictLdWith',        'args' : ['name',   '0x'],         'varg' : False },
 57 : { 'mnem' : 'MemStWith',             'args' : ['name'],                 'varg' : False },
 58 : { 'mnem' : 'DictStWith',            'args' : ['name'],                 'varg' : False },
 59 : { 'mnem' : 'ArgsMemStWith',         'args' : ['name',   '0x'],         'varg' : False },
 60 : { 'mnem' : 'ArgsDictStWith',        'args' : ['name',   '0x'],         'varg' : False },
 61 : { 'mnem' : 'MemSetWith',            'args' : ['name'],                 'varg' : False },
 62 : { 'mnem' : 'DictSetWith',           'args' : ['name'],                 'varg' : False },
 63 : { 'mnem' : 'ArgsMemSetWith',        'args' : ['name',   '0x'],         'varg' : False },
 64 : { 'mnem' : 'ArgsDictSetWith',       'args' : ['name',   '0x'],         'varg' : False },
 65 : { 'mnem' : 'ArgsCall',              'args' : ['name',   '0x'],         'varg' : False },
 66 : { 'mnem' : 'ArgsMemCall',           'args' : ['name',   '0x'],         'varg' : False },
 67 : { 'mnem' : 'ArgsMemCallWith',       'args' : ['name',   '0x'],         'varg' : False },
 68 : { 'mnem' : 'ArgsArray',             'args' : ['name',   '0x'],         'varg' : False },
 69 : { 'mnem' : 'Assert',                'args' : [],                       'varg' : False },
 70 : { 'mnem' : 'Bos',                   'args' : ['0x'],                   'varg' : False },
 71 : { 'mnem' : 'BosImplicit',           'args' : [],                       'varg' : False },
 72 : { 'mnem' : 'Bol',                   'args' : [],                       'varg' : False },
 73 : { 'mnem' : 'LdAddressOf',           'args' : [],                       'varg' : False },
 74 : { 'mnem' : 'MemAddressOf',          'args' : [],                       'varg' : False },
 75 : { 'mnem' : 'Case',                  'args' : [],                       'varg' : False },
 76 : { 'mnem' : 'CaseTo',                'args' : [],                       'varg' : False },
 77 : { 'mnem' : 'CaseGt',                'args' : [],                       'varg' : False },
 78 : { 'mnem' : 'CaseLt',                'args' : [],                       'varg' : False },
 79 : { 'mnem' : 'CaseGe',                'args' : [],                       'varg' : False },
 80 : { 'mnem' : 'CaseLe',                'args' : [],                       'varg' : False },
 81 : { 'mnem' : 'CaseNe',                'args' : [],                       'varg' : False },
 82 : { 'mnem' : 'CaseEq',                'args' : [],                       'varg' : False },
 83 : { 'mnem' : 'CaseElse',              'args' : [],                       'varg' : False },
 84 : { 'mnem' : 'CaseDone',              'args' : [],                       'varg' : False },
 85 : { 'mnem' : 'Circle',                'args' : ['0x'],                   'varg' : False },
 86 : { 'mnem' : 'Close',                 'args' : ['0x'],                   'varg' : False },
 87 : { 'mnem' : 'CloseAll',              'args' : [],                       'varg' : False },
 88 : { 'mnem' : 'Coerce',                'args' : [],                       'varg' : False },
 89 : { 'mnem' : 'CoerceVar',             'args' : [],                       'varg' : False },
 90 : { 'mnem' : 'Context',               'args' : ['context_'],             'varg' : False },
 91 : { 'mnem' : 'Debug',                 'args' : [],                       'varg' : False },
 92 : { 'mnem' : 'DefType',               'args' : ['0x', '0x'],             'varg' : False },
 93 : { 'mnem' : 'Dim',                   'args' : [],                       'varg' : False },
 94 : { 'mnem' : 'DimImplicit',           'args' : [],                       'varg' : False },
 95 : { 'mnem' : 'Do',                    'args' : [],                       'varg' : False },
 96 : { 'mnem' : 'DoEvents',              'args' : [],                       'varg' : False },
 97 : { 'mnem' : 'DoUnitil',              'args' : [],                       'varg' : False },
 98 : { 'mnem' : 'DoWhile',               'args' : [],                       'varg' : False },
 99 : { 'mnem' : 'Else',                  'args' : [],                       'varg' : False },
100 : { 'mnem' : 'ElseBlock',             'args' : [],                       'varg' : False },
101 : { 'mnem' : 'ElseIfBlock',           'args' : [],                       'varg' : False },
102 : { 'mnem' : 'ElseIfTypeBlock',       'args' : [],                       'varg' : False },
103 : { 'mnem' : 'End',                   'args' : [],                       'varg' : False },
104 : { 'mnem' : 'EndContext',            'args' : [],                       'varg' : False },
105 : { 'mnem' : 'EndFunc',               'args' : [],                       'varg' : False },
106 : { 'mnem' : 'EndIf',                 'args' : [],                       'varg' : False },
107 : { 'mnem' : 'EndIfBlock',            'args' : [],                       'varg' : False },
108 : { 'mnem' : 'EndImmediate',          'args' : [],                       'varg' : False },
109 : { 'mnem' : 'EndProp',               'args' : [],                       'varg' : False },
110 : { 'mnem' : 'EndSelect',             'args' : [],                       'varg' : False },
111 : { 'mnem' : 'EndSub',                'args' : [],                       'varg' : False },
112 : { 'mnem' : 'EndType',               'args' : [],                       'varg' : False },
113 : { 'mnem' : 'EndWith',               'args' : [],                       'varg' : False },
114 : { 'mnem' : 'Erase',                 'args' : ['0x'],                   'varg' : False },
115 : { 'mnem' : 'Error',                 'args' : [],                       'varg' : False },
116 : { 'mnem' : 'EventDecl',             'args' : ['0x', '0x', '0x', '0x'], 'varg' : False },
117 : { 'mnem' : 'RaiseEvent',            'args' : ['0x', '0x', '0x', '0x'], 'varg' : False },
118 : { 'mnem' : 'ArgsMemRaiseEvent',     'args' : ['0x', '0x', '0x', '0x'], 'varg' : False },
119 : { 'mnem' : 'ArgsMemRaiseEventWith', 'args' : ['0x', '0x', '0x', '0x'], 'varg' : False },
120 : { 'mnem' : 'ExitDo',                'args' : [],                       'varg' : False },
121 : { 'mnem' : 'ExitFor',               'args' : [],                       'varg' : False },
122 : { 'mnem' : 'ExitFunc',              'args' : [],                       'varg' : False },
123 : { 'mnem' : 'ExitProp',              'args' : [],                       'varg' : False },
124 : { 'mnem' : 'ExitSub',               'args' : [],                       'varg' : False },
125 : { 'mnem' : 'FnCurDir',              'args' : [],                       'varg' : False },
126 : { 'mnem' : 'FnDir',                 'args' : [],                       'varg' : False },
127 : { 'mnem' : 'Empty0',                'args' : [],                       'varg' : False },
128 : { 'mnem' : 'Empty1',                'args' : [],                       'varg' : False },
129 : { 'mnem' : 'FnError',               'args' : [],                       'varg' : False },
130 : { 'mnem' : 'FnFormat',              'args' : [],                       'varg' : False },
131 : { 'mnem' : 'FnFreeFile',            'args' : [],                       'varg' : False },
132 : { 'mnem' : 'FnInStr',               'args' : [],                       'varg' : False },
133 : { 'mnem' : 'FnInStr3',              'args' : [],                       'varg' : False },
134 : { 'mnem' : 'FnInStr4',              'args' : [],                       'varg' : False },
135 : { 'mnem' : 'FnInStrB',              'args' : [],                       'varg' : False },
136 : { 'mnem' : 'FnInStrB3',             'args' : [],                       'varg' : False },
137 : { 'mnem' : 'FnInStrB4',             'args' : [],                       'varg' : False },
138 : { 'mnem' : 'FnLBound',              'args' : ['0x'],                   'varg' : False },
139 : { 'mnem' : 'FnMid',                 'args' : [],                       'varg' : False },
140 : { 'mnem' : 'FnMidB',                'args' : [],                       'varg' : False },
141 : { 'mnem' : 'FnStrComp',             'args' : [],                       'varg' : False },
142 : { 'mnem' : 'FnStrComp3',            'args' : [],                       'varg' : False },
143 : { 'mnem' : 'FnStringVar',           'args' : [],                       'varg' : False },
144 : { 'mnem' : 'FnStringStr',           'args' : [],                       'varg' : False },
145 : { 'mnem' : 'FnUBound',              'args' : ['0x'],                   'varg' : False },
146 : { 'mnem' : 'For',                   'args' : [],                       'varg' : False },
147 : { 'mnem' : 'ForEach',               'args' : [],                       'varg' : False },
148 : { 'mnem' : 'ForEachAs',             'args' : [],                       'varg' : False },
149 : { 'mnem' : 'ForStep',               'args' : [],                       'varg' : False },
150 : { 'mnem' : 'FuncDefn',              'args' : ['func_'],                'varg' : False },
151 : { 'mnem' : 'FuncDefnSave',          'args' : ['func_'],                'varg' : False },
152 : { 'mnem' : 'GetRec',                'args' : [],                       'varg' : False },
153 : { 'mnem' : 'GoSub',                 'args' : ['name'],                 'varg' : False },
154 : { 'mnem' : 'GoTo',                  'args' : ['name'],                 'varg' : False },
155 : { 'mnem' : 'If',                    'args' : [],                       'varg' : False },
156 : { 'mnem' : 'IfBlock',               'args' : [],                       'varg' : False },
157 : { 'mnem' : 'TypeOf',                'args' : ['imp_'],                 'varg' : False },
158 : { 'mnem' : 'IfTypeBlock',           'args' : [],                       'varg' : False },
159 : { 'mnem' : 'Implements',            'args' : ['0x', '0x', '0x', '0x'], 'varg' : False },
160 : { 'mnem' : 'Input',                 'args' : [],                       'varg' : False },
161 : { 'mnem' : 'InputDone',             'args' : [],                       'varg' : False },
162 : { 'mnem' : 'InputItem',             'args' : [],                       'varg' : False },
163 : { 'mnem' : 'Label',                 'args' : ['name'],                 'varg' : False },
164 : { 'mnem' : 'Let',                   'args' : [],                       'varg' : False },
165 : { 'mnem' : 'Line',                  'args' : ['0x'],                   'varg' : False },
166 : { 'mnem' : 'LineCont',              'args' : [],                       'varg' :  True },
167 : { 'mnem' : 'LineInput',             'args' : [],                       'varg' : False },
168 : { 'mnem' : 'LineNum',               'args' : ['name'],                 'varg' : False },
169 : { 'mnem' : 'LitCy',                 'args' : ['0x', '0x', '0x', '0x'], 'varg' : False },
170 : { 'mnem' : 'LitDate',               'args' : ['0x', '0x', '0x', '0x'], 'varg' : False },
171 : { 'mnem' : 'LitDefault',            'args' : [],                       'varg' : False },
172 : { 'mnem' : 'LitDI2',                'args' : ['0x'],                   'varg' : False },
173 : { 'mnem' : 'LitDI4',                'args' : ['0x', '0x'],             'varg' : False },
174 : { 'mnem' : 'LitDI8',                'args' : ['0x', '0x', '0x', '0x'], 'varg' : False },
175 : { 'mnem' : 'LitHI2',                'args' : ['0x'],                   'varg' : False },
176 : { 'mnem' : 'LitHI4',                'args' : ['0x', '0x'],             'varg' : False },
177 : { 'mnem' : 'LitHI8',                'args' : ['0x', '0x', '0x', '0x'], 'varg' : False },
178 : { 'mnem' : 'LitNothing',            'args' : [],                       'varg' : False },
179 : { 'mnem' : 'LitOI2',                'args' : ['0x'],                   'varg' : False },
180 : { 'mnem' : 'LitOI4',                'args' : ['0x', '0x'],             'varg' : False },
181 : { 'mnem' : 'LitOI8',                'args' : ['0x', '0x', '0x', '0x'], 'varg' : False },
182 : { 'mnem' : 'LitR4',                 'args' : ['0x', '0x'],             'varg' : False },
183 : { 'mnem' : 'LitR8',                 'args' : ['0x', '0x', '0x', '0x'], 'varg' : False },
184 : { 'mnem' : 'LitSmallI2',            'args' : [],                       'varg' : False },
185 : { 'mnem' : 'LitStr',                'args' : [],                       'varg' :  True },
186 : { 'mnem' : 'LitVarSpecial',         'args' : [],                       'varg' : False },
187 : { 'mnem' : 'Lock',                  'args' : [],                       'varg' : False },
188 : { 'mnem' : 'Loop',                  'args' : [],                       'varg' : False },
189 : { 'mnem' : 'LoopUntil',             'args' : [],                       'varg' : False },
190 : { 'mnem' : 'LoopWhile',             'args' : [],                       'varg' : False },
191 : { 'mnem' : 'LSet',                  'args' : [],                       'varg' : False },
192 : { 'mnem' : 'Me',                    'args' : [],                       'varg' : False },
193 : { 'mnem' : 'MeImplicit',            'args' : [],                       'varg' : False },
194 : { 'mnem' : 'MemRedim',              'args' : ['name', '0x', 'type_'],  'varg' : False },
195 : { 'mnem' : 'MemRedimWith',          'args' : ['name', '0x', 'type_'],  'varg' : False },
196 : { 'mnem' : 'MemRedimAs',            'args' : ['name', '0x', 'type_'],  'varg' : False },
197 : { 'mnem' : 'MemRedimAsWith',        'args' : ['name', '0x', 'type_'],  'varg' : False },
198 : { 'mnem' : 'Mid',                   'args' : [],                       'varg' : False },
199 : { 'mnem' : 'MidB',                  'args' : [],                       'varg' : False },
200 : { 'mnem' : 'Name',                  'args' : [],                       'varg' : False },
201 : { 'mnem' : 'New',                   'args' : ['imp_'],                 'varg' : False },
202 : { 'mnem' : 'Next',                  'args' : [],                       'varg' : False },
203 : { 'mnem' : 'NextVar',               'args' : [],                       'varg' : False },
204 : { 'mnem' : 'OnError',               'args' : ['name'],                 'varg' : False },
205 : { 'mnem' : 'OnGosub',               'args' : [],                       'varg' :  True },
206 : { 'mnem' : 'OnGoto',                'args' : [],                       'varg' :  True },
207 : { 'mnem' : 'Open',                  'args' : ['0x'],                   'varg' : False },
208 : { 'mnem' : 'Option',                'args' : [],                       'varg' : False },
209 : { 'mnem' : 'OptionBase',            'args' : [],                       'varg' : False },
210 : { 'mnem' : 'ParamByVal',            'args' : [],                       'varg' : False },
211 : { 'mnem' : 'ParamOmitted',          'args' : [],                       'varg' : False },
212 : { 'mnem' : 'ParamNamed',            'args' : ['name'],                 'varg' : False },
213 : { 'mnem' : 'PrintChan',             'args' : [],                       'varg' : False },
214 : { 'mnem' : 'PrintComma',            'args' : [],                       'varg' : False },
215 : { 'mnem' : 'PrintEos',              'args' : [],                       'varg' : False },
216 : { 'mnem' : 'PrintItemComma',        'args' : [],                       'varg' : False },
217 : { 'mnem' : 'PrintItemNL',           'args' : [],                       'varg' : False },
218 : { 'mnem' : 'PrintItemSemi',         'args' : [],                       'varg' : False },
219 : { 'mnem' : 'PrintNL',               'args' : [],                       'varg' : False },
220 : { 'mnem' : 'PrintObj',              'args' : [],                       'varg' : False },
221 : { 'mnem' : 'PrintSemi',             'args' : [],                       'varg' : False },
222 : { 'mnem' : 'PrintSpc',              'args' : [],                       'varg' : False },
223 : { 'mnem' : 'PrintTab',              'args' : [],                       'varg' : False },
224 : { 'mnem' : 'PrintTabComma',         'args' : [],                       'varg' : False },
225 : { 'mnem' : 'PSet',                  'args' : ['0x'],                   'varg' : False },
226 : { 'mnem' : 'PutRec',                'args' : [],                       'varg' : False },
227 : { 'mnem' : 'QuoteRem',              'args' : ['0x'],                   'varg' :  True },
228 : { 'mnem' : 'Redim',                 'args' : ['name', '0x', 'type_'],  'varg' : False },
229 : { 'mnem' : 'RedimAs',               'args' : ['name', '0x', 'type_'],  'varg' : False },
230 : { 'mnem' : 'Reparse',               'args' : [],                       'varg' :  True },
231 : { 'mnem' : 'Rem',                   'args' : [],                       'varg' :  True },
232 : { 'mnem' : 'Resume',                'args' : ['name'],                 'varg' : False },
233 : { 'mnem' : 'Return',                'args' : [],                       'varg' : False },
234 : { 'mnem' : 'RSet',                  'args' : [],                       'varg' : False },
235 : { 'mnem' : 'Scale',                 'args' : ['0x'],                   'varg' : False },
236 : { 'mnem' : 'Seek',                  'args' : [],                       'varg' : False },
237 : { 'mnem' : 'SelectCase',            'args' : [],                       'varg' : False },
238 : { 'mnem' : 'SelectIs',              'args' : [],                       'varg' : False },
239 : { 'mnem' : 'SelectType',            'args' : [],                       'varg' : False },
240 : { 'mnem' : 'SetStmt',               'args' : [],                       'varg' : False },
241 : { 'mnem' : 'Stack',                 'args' : [],                       'varg' : False },
242 : { 'mnem' : 'Stop',                  'args' : [],                       'varg' : False },
243 : { 'mnem' : 'Type',                  'args' : ['rec_'],                 'varg' : False },
244 : { 'mnem' : 'Unlock',                'args' : [],                       'varg' : False },
245 : { 'mnem' : 'VarDefn',               'args' : ['var_'],                 'varg' : False },
246 : { 'mnem' : 'Wend',                  'args' : [],                       'varg' : False },
247 : { 'mnem' : 'While',                 'args' : [],                       'varg' : False },
248 : { 'mnem' : 'With',                  'args' : [],                       'varg' : False },
249 : { 'mnem' : 'WriteChan',             'args' : [],                       'varg' : False },
250 : { 'mnem' : 'ConstFuncExpr',         'args' : [],                       'varg' : False },
251 : { 'mnem' : 'LbConst',               'args' : ['name'],                 'varg' : False },
252 : { 'mnem' : 'LbIf',                  'args' : [],                       'varg' : False },
253 : { 'mnem' : 'LbElse',                'args' : [],                       'varg' : False },
254 : { 'mnem' : 'LbElseIf',              'args' : [],                       'varg' : False },
255 : { 'mnem' : 'LbEndIf',               'args' : [],                       'varg' : False },
256 : { 'mnem' : 'LbMark',                'args' : [],                       'varg' : False },
257 : { 'mnem' : 'EndForVariable',        'args' : [],                       'varg' : False },
258 : { 'mnem' : 'StartForVariable',      'args' : [],                       'varg' : False },
259 : { 'mnem' : 'NewRedim',              'args' : [],                       'varg' : False },
260 : { 'mnem' : 'StartWithExpr',         'args' : [],                       'varg' : False },
261 : { 'mnem' : 'SetOrSt',               'args' : ['name'],                 'varg' : False },
262 : { 'mnem' : 'EndEnum',               'args' : [],                       'varg' : False },
263 : { 'mnem' : 'Illegal',               'args' : [],                       'varg' : False }
}

def translateOpcode(opcode, vbaVer):
    if   (vbaVer == 3):
        if   (  0 <= opcode <=  67):
            return opcode
        elif ( 68 <= opcode <=  70):
            return opcode +  2
        elif ( 71 <= opcode <= 111):
            return opcode +  4
        elif (112 <= opcode <= 150):
            return opcode +  8
        elif (151 <= opcode <= 164):
            return opcode +  9
        elif (165 <= opcode <= 166):
            return opcode + 10
        elif (167 <= opcode <= 169):
            return opcode + 11
        elif (170 <= opcode <= 238):
            return opcode + 12
        else:	# opcode == 239
            return opcode + 24
    elif (vbaVer == 5):
        if   (  0 <= opcode <=  68):
            return opcode
        elif ( 69 <= opcode <=  71):
            return opcode +  1
        elif ( 72 <= opcode <= 112):
            return opcode +  3
        elif (113 <= opcode <= 151):
            return opcode +  7
        elif (152 <= opcode <= 165):
            return opcode +  8
        elif (166 <= opcode <= 167):
            return opcode +  9
        elif (168 <= opcode <= 170):
            return opcode + 10
        else:	# 171 <= opcode <= 252
            return opcode + 11
    elif (vbaVer == 6):
        if   (  0 <= opcode <= 173):
            return opcode
        elif (174 <= opcode <= 175):
            return opcode +  1
        elif (176 <= opcode <= 178):
            return opcode +  2
        else:	# 179 <= opcode <= 260
            return opcode +  3
    else:
        return opcode

def getID(idCode, identifiers):
    internalNames = [
	'<crash>', '0', 'Abs', 'Access', 'AddressOf', 'Alias', 'And', 'Any',
	'Append', 'Array', 'As', 'Assert', 'B', 'Base', 'BF', 'Binary',
	'Boolean', 'ByRef', 'Byte', 'ByVal', 'Call', 'Case', 'CBool', 'CByte',
	'CCur', 'CDate', 'CDec', 'CDbl', 'CDecl', 'ChDir', 'CInt', 'Circle',
	'CLng', 'Close', 'Compare', 'Const', 'CSng', 'CStr', 'CurDir', 'CurDir$',
	'CVar', 'CVDate', 'CVErr', 'Currency', 'Database', 'Date', 'Date$', 'Debug',
	'Decimal', 'Declare', 'DefBool', 'DefByte', 'DefCur', 'DefDate', 'DefDec', 'DefDbl',
	'DefInt', 'DefLng', 'DefObj', 'DefSng', 'DefStr', 'DefVar', 'Dim', 'Dir',
	'Dir$', 'Do', 'DoEvents', 'Double', 'Each', 'Else', 'ElseIf', 'Empty',
	'End', 'EndIf', 'Enum', 'Eqv', 'Erase', 'Error', 'Error$', 'Event',
	'WithEvents', 'Exit', 'Explicit', 'F', 'False', 'Fix', 'For', 'Format',
	'Format$', 'FreeFile', 'Friend', 'Function', 'Get', 'Global', 'Go', 'GoSub',
	'Goto', 'If', 'Imp', 'Implements', 'In', 'Input', 'Input$', 'InputB',
	'InputB', 'InStr', 'InputB$', 'Int', 'InStrB', 'Is', 'Integer', 'Left',
	'LBound', 'LenB', 'Len', 'Lib', 'Let', 'Line', 'Like', 'Load',
	'Local', 'Lock', 'Long', 'Loop', 'LSet', 'Me', 'Mid', 'Mid$',
	'MidB', 'MidB$', 'Mod', 'Module', 'Name', 'New', 'Next', 'Not',
	'Nothing', 'Null', 'Object', 'On', 'Open', 'Option', 'Optional', 'Or',
	'Output', 'ParamArray', 'Preserve', 'Print', 'Private', 'Property', 'PSet', 'Public',
	'Put', 'RaiseEvent', 'Random', 'Randomize', 'Read', 'ReDim', 'Rem', 'Resume',
	'Return', 'RGB', 'RSet', 'Scale', 'Seek', 'Select', 'Set', 'Sgn',
	'Shared', 'Single', 'Spc', 'Static', 'Step', 'Stop', 'StrComp', 'String',
	'String$', 'Sub', 'Tab', 'Text', 'Then', 'To', 'True', 'Type',
	'TypeOf', 'UBound', 'Unload', 'Unlock', 'Unknown', 'Until', 'Variant', 'WEnd',
	'While', 'Width', 'With', 'Write', 'Xor', '#Const', '#Else', '#ElseIf',
	'#End', '#If', 'Attribute', 'VB_Base', 'VB_Control', 'VB_Creatable', 'VB_Customizable', 'VB_Description',
	'VB_Exposed', 'VB_Ext_Key', 'VB_HelpID', 'VB_Invoke_Func', 'VB_Invoke_Property', 'VB_Invoke_PropertyPut', 'VB_Invoke_PropertyPutRef', 'VB_MemberFlags',
	'VB_Name', 'VB_PredecraredID', 'VB_ProcData', 'VB_TemplateDerived', 'VB_VarDescription', 'VB_VarHelpID', 'VB_VarMemberFlags', 'VB_VarProcData',
	'VB_UserMemID', 'VB_VarUserMemID', 'VB_GlobalNameSpace', ',', '.', '"', '_', '!',
	'#', '&', "'", '(', ')', '*', '+', '-',
	' /', ':', ';', '<', '<=', '<>', '=', '=<',
	'=>', '>', '><', '>=', '?', '\\', '^', ':='
    ]

    idCode >>= 1
    if (idCode >= 0x100):
        return identifiers[idCode - 0x100]
    else:
        return internalNames[idCode]

def dumpLine(moduleData, lineStart, lineLength, endian, vbaVer, identifiers, verbose, line):
    varTypes = ['', '?', '%', '&', '!', '#', '@', '?', '$', '?', '?', '?', '?', '?']
    varTypesLong = ['Var', '?', 'Int', 'Lng', 'Sng', 'Dbl', 'Cur', 'Date', 'Str', 'Obj', 'Err', 'Bool', 'Var']
    specials = ['False', 'True', 'Null', 'Empty']
    options = ['Base 0', 'Base 1', 'Compare Text', 'Compare Binary', 'Explicit', 'Private Module']

    print('Line #%d:' % line)
    if (verbose):
        print(hexdump3(moduleData[lineStart:lineStart + lineLength], length=16))
    offset = lineStart
    endOfLine = lineStart + lineLength
    while (offset < endOfLine):
        offset, opcode = getVar(moduleData, offset, endian, False)
        opType = (opcode & ~0x03FF) >> 10
        opcode &= 0x03FF
        translatedOpcode = translateOpcode(opcode, vbaVer)
        if (not translatedOpcode in opcodes):
            print('Unrecognized opcode 0x%04X at offset 0x%08X.' % (opcode, offset))
            return
        instruction = opcodes[translatedOpcode]
        mnemonic = instruction['mnem']
        print('\t', end='')
        if (verbose):
            print('%04X ' % opcode, end='')
        print('%s ' % mnemonic, end='')
        if (mnemonic in ['Coerce', 'CoerceVar', 'DefType']):
            if (opType < len(varTypesLong)):
                print('(%s) ' % varTypesLong[opType], end='')
            else:
                print('(%d)' % opType, end='')
        elif (mnemonic in ['Dim', 'DimImplicit', 'Type']):
            if   (opType ==  8):
                print('(Public) ', end='')
            elif (opType == 16):
                print('(Private) ', end='')
            elif (opType == 32):
                print('(Static) ', end='')
        elif (mnemonic == 'LitVarSpecial'):
            print('(%s)' % specials[opType], end='')
        elif (mnemonic == 'FuncDefn'):
            if   (opType == 1):
                print('(Sub / Property Set) ', end='')
            elif (opType == 2):
                print('(Function / Property Get) ', end='')
            elif (opType == 5):
                print('(Public Sub / Property Set) ', end='')
            elif (opType == 6):
                print('(Public Function / Property Get) ', end='')
        elif (mnemonic in ['ArgsCall', 'ArgsMemCall', 'ArgsMemCallWith']):
            if (opType < 16):
                print('(Call) ', end='')
            else:
                opType -= 16
        elif (mnemonic == 'Option'):
            print(' (%s)' % options[opType], end='')
        elif (mnemonic in ['ReDim', 'RedimAs']):
            if (opType):
                print('(Preserve) ', end='')
        for arg in instruction['args']:
            if (arg == 'name'):
                offset, word = getVar(moduleData, offset, endian, False)
                varName = getID(word, identifiers)
                if (opType < len(varTypes)):
                    strType = varTypes[opType]
                else:
                    strType = ''
                    if (opType == 32):
                        varName = '[' + varName + ']'
                if   (mnemonic == 'OnError'):
                    strType = ''
                    if   (opType == 1):
                        varName = '(Resume Next)'
                    elif (opType == 2):
                        varName = '(GoTo 0)'
                elif (mnemonic == 'Resume'):
                    strType = ''
                    if   (opType == 1):
                        varName = '(Next)'
                    elif (opType != 0):
                        varName = ''
                print('%s%s ' % (varName, strType), end='')
            elif (arg in ['0x', 'imp_']):
                offset, word = getVar(moduleData, offset, endian, False)
                if (mnemonic != 'Open'):
                    print('%s%04X ' % (arg, word), end='')
                else:
                    # This is a rather messy way of processing what is probably
                    # just a bit field but I couldn't figure out a smarter way
                    mode = word & 0x00FF
                    access = (word & 0xFF00) >> 8
                    print('(For ', end='')
                    if   (mode == 0x01):
                        print('Input', end='')
                    elif (mode == 0x02):
                        print('Output', end='')
                    elif (mode == 0x04):
                        print('Random', end='')
                    elif (mode == 0x08):
                        print('Append', end='')
                    elif (mode == 0x20):
                        print('Binary', end='')
                    if   (access == 0x01):
                        print(' Access Read', end='')
                    elif (access == 0x02):
                        print(' Access Write', end='')
                    elif (access == 0x03):
                        print(' Access Read Write', end='')
                    elif (access == 0x10):
                        print(' Lock Read Write', end='')
                    elif (access == 0x20):
                        print(' Lock Write', end='')
                    elif (access == 0x30):
                        print(' Lock Read', end='')
                    elif (access == 0x40):
                        print(' Shared', end='')
                    print(')', end='')
            elif (arg in ['func_', 'var_', 'rec_', 'type_', 'context_']):
                offset, dword = getVar(moduleData, offset, endian, True)
                print('%s%08X ' % (arg, dword), end='')
        if (instruction['varg']):
            offset, wLength = getVar(moduleData, offset, endian, False)
            substring = moduleData[offset:offset + wLength]
            print('0x%04X ' % wLength, end='')
            if (mnemonic in ['LitStr', 'QuoteRem', 'Rem']):
                print('"%s"' % substring, end='')
            elif (mnemonic in ['OnGosub', 'OnGoto']):
                offset1 = offset
                vars = []
                for i in range(wLength / 2):
                    offset1, word = getVar(moduleData, offset1, endian, False)
                    vars.append(getID(word, identifiers))
                print('%s ' % (', '.join(v for v in vars)), end='')
            else:
                hexdump = ' '.join('{:02X}'.format(ord(c)) for c in substring)
                print('%s' % hexdump, end='')
            offset += wLength
            if (wLength & 1):
                offset += 1
        print('')

def pcodeDump(moduleData, vbaProjectData, dirData, identifiers, verbose, disasmonly):
    if (verbose and not disasmonly):
        print(hexdump3(moduleData, length=16))
    # Determine endinanness: PC (little-endian) or Mac (big-endian)
    if (getWord(moduleData, 2, '<') > 0xFF):
        endian = '>'
    else:
        endian = '<'
    # TODO:
    #	- Handle VBA3 modules
    vbaVer = 3
    try:
        version = getWord(vbaProjectData, 2, endian)
        # TODO:
        #	- Handle VBA7
        if (version >= 0x6B):
            # VBA6
            vbaVer = 6
            offset = 0x0019
        else:
            # VBA5
            vbaVer = 5
            offset = 11
            offset = skipStructure(moduleData, offset, endian,  True,  1, False)
            offset += 64
            offset = skipStructure(moduleData, offset, endian, False, 16, False)
            offset = skipStructure(moduleData, offset, endian,  True,  1, False)
            offset += 6
            offset = skipStructure(moduleData, offset, endian,  True,  1, False)
            offset += 77
        dwLength = getDWord(moduleData, offset, endian)
        offset = dwLength + 0x003C
        offset, magic = getVar(moduleData, offset, endian, False)
        if (magic != 0xCAFE):
            return
        offset += 2
        offset, numLines = getVar(moduleData, offset, endian, False)
        pcodeStart = offset + numLines * 12 + 10
        for line in range(numLines):
            offset += 4
            offset, lineLength = getVar(moduleData, offset, endian, False)
            offset += 2
            offset, lineOffset = getVar(moduleData, offset, endian, True)
            dumpLine(moduleData, pcodeStart + lineOffset, lineLength, endian, vbaVer, identifiers, verbose, line)
    except Exception as e:
        print('Error: %s.' % e, file=sys.stderr)
    return

def processFile(fileName, verbose, disasmonly):
    # TODO:
    #	- Handle VBA3 documents
    print('Processing file: %s' % fileName)
    try:
        vbaParser = VBA_Parser(fileName)
        vbaProjects = vbaParser.find_vba_projects()
        if (vbaProjects is None):
            vbaParser.close()
            return
        for vbaRoot, projectPath, dirPath in vbaProjects:
            print('=' * 79)
            if (not disasmonly):
                print('dir stream: %s' % dirPath)
            dirData, codeModules = processDir(vbaParser, dirPath, verbose, disasmonly)
            vbaProjectPath = vbaRoot + 'VBA/_VBA_PROJECT'
            vbaProjectData = process_VBA_PROJECT(vbaParser, vbaProjectPath, verbose, disasmonly)
            identifiers = getTheIdentifiers(vbaProjectData)
            if (not disasmonly):
                print('Identifiers:')
                print('')
                for identifier in identifiers:
                    print('%s' % identifier)
                print('')
                print('_VBA_PROJECT parsing done.')
            if (not disasmonly):
                print('-' * 79)
            print('Module streams:')
            for module in codeModules:
                modulePath = vbaRoot + 'VBA/' + module
                moduleData = vbaParser.ole_file.openstream(modulePath).read()
                print ('%s - %d bytes' % (modulePath, len(moduleData)))
                pcodeDump(moduleData, vbaProjectData, dirData, identifiers, verbose, disasmonly)
    except Exception as e:
        print('Error: %s.' % e, file=sys.stderr)
    vbaParser.close()

if __name__ == '__main__':
    parser = argparse.ArgumentParser(version='%(prog)s version ' + __VERSION__,
	description='Dumps the p-code of VBA-containing documents.')
    parser.add_argument('-n', '--norecurse', action='store_true',
	help="don't recurse into directories")
    parser.add_argument('-d', '--disasmonly', action='store_true',
	help='only disassemble, no stream dumps')
    parser.add_argument('--verbose', action='store_true',
	help='dump the stream contents')
    parser.add_argument('fileOrDir', nargs='+', help='file or dir')
    args = parser.parse_args()
    try:
        for name in args.fileOrDir:
            if os.path.isdir(name):
                for name, subdirList, fileList in os.walk(name):
                    for fname in fileList:
                        fullName = os.path.join(name, fname)
                        processFile(fullName, args.verbose, args.disasmonly)
                    if args.norecurse:
                        while len(subdirList) > 0:
                            del(subdirList[0])
            elif os.path.isfile(name):
                processFile(name, args.verbose, args.disasmonly)
            else:
                print(name + ' does not exist.', file=sys.stderr)
    except Exception as e:
        print('Error: %s.' % e, file=sys.stderr)
        sys.exit(-1)
    sys.exit(0)
