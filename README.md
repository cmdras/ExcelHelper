# ExcelHelper


A Python method that can extract data from a Excel spreadsheet file (.xlsx) .

Usage:

GetDataFromXlsx('path/to/xlsx/file.xlsx')

->  Returns a dictionary containing data from each sheet. The values of each column is stored in an array.
    
Example:
```
{
    'sheet1' : {
        'header1' : ['foo', 'bar', 'cake'],
        'header2' : ['ice', 'ice', 'baby']
    },

    'sheet2' : {
        'A' : ['Hello', 'World!'],
    },
}
```

# Optional parameters

* tmpDataDir: string containing the absolute path to the
            directory in which intermediate files may be stored.
            Default is directory of this file
            For example: '/Users/user/bar' or 'C:\user\bar'

* headers: Boolean which specifies whether or not the excel
            file contains headers as its first row. It will parse the output 
            accordingly
            
# Some other usage examples:
```
GetDataFromXlsx('path/to/xlsx/file.xlsx', tmpDatadir='path/to/xlsx/tmp/dir')


GetDataFromXlsx('path/to/xlsx/file.xlsx', tmpDatadir='path/to/xlsx/tmp/dir', headers=True)


GetDataFromXlsx('path/to/xlsx/file.xlsx', headers=False')
```

Feedback and tips are very welcome!
