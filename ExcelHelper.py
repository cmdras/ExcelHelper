#!/usr/bin/env python2
# -*- coding: utf-8 -*-
# =============================================================================
# Created By  : Christopher Ras
# Created Date: Sun January 27 2019
# Version 0.0.1
# =============================================================================
# License : MIT
# =============================================================================

import traceback, subprocess, os, shutil
import xml.etree.ElementTree as ET
import codecs
import sys
import zipfile

def GetDataFromXlsx(xlsxFile, tmpDataDir=None, headers=False):
    ''' Main method for extracting data from a Microsoft Excel file (.xlsx)
        Arguments:
        - xlsxFile: string containing the absolute path to the Excel file in
            to be read. For example: '/Users/user/foo.xlsx' or 'C:\user\foo.xlsx'

        - tmpDataDir (Optional): string containing the absolute path to the
            directory in which intermediate files may be stored.
            Default is directory of this file
            For example: '/Users/user/bar' or 'C:\user\bar'

        - headers (Optional): Boolean which specifies whether or not the excel
            file contains headers as its first row. It will parse the output 
            accordingly

        Returns:
        Dictionary of the data stored in the Excel file
        Each sheet will be stored as its own key in the dictionary
        The value of the sheet key is a dictionary of the data in that specific sheet
            The keys of the inner dictionary are the headers if headers==True, or Column letter.
            The value of these keys are arrays representing the column underneath the header

        For example:
        {
            'sheet1' : {
                'header1' : ['foo', 'bar', 'cake'],
                'header2' : ['ice', 'ice', 'baby']
            },

            'sheet2' : {
                'A' : ['Hello', 'World!'],
            },
        }
    '''

    convertSuccess, tmpDir = __ConvertXlsxToXml(xlsxFile, tmpDataDir)
    if not convertSuccess:
        print 'Conversion of {} unsuccessfull'.format(os.path.basename(xlsxFile))
        return

    xlDir = os.path.join(tmpDir, 'xl', 'worksheets')
    sheetFiles = __GetSheetFiles(xlDir)
    xlSharedStringsFile = os.path.join(tmpDir, 'xl', 'sharedStrings.xml')
    parsedResult = __ParseXmlsOfXlsxFile(xlSharedStringsFile, sheetFiles, headers)
    try:
        ArchiveContentsOfTmpDirectory(tmpDir)
    except Exception as e:
        print 'an error occured while removing the tmp directory: {}'.format(tmpDir)
        print traceback.format_exc()

    return parsedResult

def __ConvertXlsxToXml(xlsxFile, tmpDataDir):
    success = False

    curDir = os.path.dirname(os.path.realpath(__file__))
    tmpDir = __ValidateDirectories(curDir, tmpDataDir)

    print 'Trying to unzip the file: {}'.format(xlsxFile)
    try:
        zipRef = zipfile.ZipFile(xlsxFile, 'r')
        zipRef.extractall(tmpDir)
        success = True
        print 'Successfully unzipped {} !'.format(xlsxFile)
    except Exception as e:
        print 'An error occured while converting {} to a xml file: {}'.format(xlsxFile, e)
        print traceback.format_exc()
    finally:
        zipRef.close()

    return (success, tmpDir)

def __ValidateDirectories(curDir, tmpDataDir):
    tmpDir = os.path.join(curDir, 'tmp') if not tmpDataDir else os.path.join(tmpDataDir, 'tmp')
    if not os.path.exists(tmpDir):
        os.mkdir(tmpDir)
    return tmpDir

def __GetSheetFiles(sheetFileDir):
    sheetFiles = []
    for subdir, dirs, files in os.walk(sheetFileDir):
        for file in files:
            ext = os.path.splitext(file)[-1].lower()
            if ext and ext == '.xml':
                sheetFiles.append(os.path.join(sheetFileDir, file))
    return sheetFiles

def __ParseXmlsOfXlsxFile(sharedStringsFile, sheets, headers):
    ''' Adapted from code presented by Simon Duff
        https://simonduff.net/processing_excel_xlsx_files_with_python/
    '''

    UTF8Writer = codecs.getwriter('utf8')
    sys.stdout = UTF8Writer(sys.stdout)

    x = ET.parse(open(sharedStringsFile))
    sharedStrings = x.getroot()
    last_v = None
    last_type = None

    # We have a value if event == 'end' and tag == 'c'
    # The value is in sharedStrings, which is set when the 
    # event == 'start', tag == 'v' and last_type == 's'
    # Last_type is set when event == 'start', tag == 'c' and there is a 't'
    # in the elem.attrib

    sheetNames = __GetSheetNames(sheets)
    sheetDict = {}
    for i in xrange(len(sheetNames)):
        sheetDict[sheetNames[i]] = sheets[i]
    
    result = { sheetName : {} for sheetName in sheetNames }
    for sheetName in sheetNames:
        for event, elem in ET.iterparse(sheetDict[sheetName], events=('start', 'end')):
            uri, tag = elem.tag.split("}")
            if event == "start" and tag == "c":
                    last_v = None
                    if "t" in elem.attrib:
                            last_type = elem.attrib["t"]
                    else:
                            last_type = None
            elif event == "end" and tag == "c":
                    if "r" in elem.attrib:
                            rc = elem.attrib["r"]
                            if last_v != None:
                                    col = __GetColumnOfValue(rc)
                                    # print "RC is ", rc, " = ", last_v
                                    if col not in result[sheetName]:
                                        result[sheetName][col] = []
                                    result[sheetName][col].append(last_v)
            elif event == "start" and tag == "v": # start v tag
                    value = "".join(elem.itertext())
                    if not value:
                        continue
                    if last_type == "s":
                            last_v = "".join(sharedStrings[int(value)].itertext())
                    else:
                            last_v = "{}".format(value)
    
    return __ParseResultWithHeaders(result) if headers else result

def __GetColumnOfValue(rc):
    return rc[0] if rc else 'unkown'

def __ParseResultWithHeaders(result):
    parsed = {}
    for sheet in result.keys():
        parsed[sheet] = {}
        for col in result[sheet].keys():
            header = result[sheet][col].pop(0)
            parsed[sheet][header] = result[sheet][col]

    return parsed

def __GetSheetNames(sheets):
    return [os.path.basename(sheet).split('.')[0] for sheet in sheets]

def ArchiveContentsOfTmpDirectory(tmpDir):
    if os.path.isdir(tmpDir):
        archiveDir = os.path.join(tmpDir, 'Archived')
        if os.path.isdir(archiveDir):
            shutil.rmtree(archiveDir)
        os.mkdir(archiveDir)
        files = os.listdir(tmpDir)
        for file in files:
            if file != 'Archived': # Dont move myself to myself
                shutil.move(os.path.join(tmpDir, file), archiveDir)
