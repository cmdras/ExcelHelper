from External import xlsx2csv
import traceback, datetime, csv, os
import time

def GetDataFromXlsx(xlsxFile, complete=True, headers=False):
    start = time.time()
    excelFile = xlsx2csv.Xlsx2csv(xlsxFile)
    sheetIds = [1] if not complete else [x['id'] for x in excelFile.workbook.sheets]
    csvs = __ConvertToCsv(xlsxFile, excelFile, sheetIds)
    data = __CombineDataFromCsvs(csvs)
    parsedData = __ParseSheetData(data, headers)
    end = time.time()
    print end-start
    return parsedData


def __ConvertToCsv(xlsxFileName, excelFile, sheets):
    csvs = []
    date = datetime.datetime.now()
    datestr = date.strftime("%Y%m%d%H%M")
    if not os.path.isdir("/tmp"):
        os.mkdir("tmp")
    for sheetId in sheets:
        csvName = os.path.join('tmp', r'{}{}{}{}{}'.format(xlsxFileName, sheetId, '-', datestr, '.csv'))
        try:
            excelFile.convert(csvName, sheetId)
            print 'Successfully converted {}, sheet {} to csv!'.format(xlsxFileName, sheetId)
            csvs.append(csvName)
        except Exception as e:
            print 'An error occured while converting {}, sheet {} to csv: {}'.format(xlsxFileName, sheetId, e)
            print traceback.format_exc()
    return csvs

def __CombineDataFromCsvs(csvs):
    data = []
    try:
        for csvFile in csvs:
            with open(csvFile, 'r') as csvreader:
                reader = csv.reader(csvreader)
                data.append([row for row in reader])
    except Exception as e:
        print 'An error occured while combining data from the sheets: '
        print traceback.format_exc()
    return data

def __ParseSheetData(data, headers=False):
    sheetCounter = 1
    result = {}
    try:
        for sheet in data:
            if not sheet:
                continue
            sheetStr = "Sheet{}".format(sheetCounter)
            headerData = sheet.pop(0) if headers else ["col{}".format(i) for i in xrange(len(sheet[0]))]
            result[sheetStr] = __GetColumnData(headerData, sheet)
    except Exception as e:
        print 'An error occured while parsing the sheet data: '
        print traceback.format_exc()
    
    return result
              
def __GetColumnData(headerData, sheet):
    data = {}
    for i in xrange(len(headerData)):
        data[headerData[i]] = [row[i] for row in sheet]
    return data