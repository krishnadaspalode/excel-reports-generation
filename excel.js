function exportReport() {
    var workbook = {};

    workbook = {
        SheetNames: ['sheet 1', 'sheet 2'],
        Sheets: {
            'sheet 1': createWorkbook(columnStructure, reportData),
            'sheet 2': createWorkbook(columnStructure, reportData)
        }
    };
    createExcelSheet(workbook, 'ReportName.xlsx');
}
var reportData = {
    "userDetails": [{
        "column1": "Test User 1",
        "column2": "9/17/2016, 1:21:34 PM",
        "column3": "9/18/2016, 1:21:34 PM",
        "column4": "0",
        "column5": "000",
        "column6": "000",
        "column7": "0",
        "column8": "0",
        "column9": "0",
        "column10": "0",
        "column11": "0",
        "column12": "0"
    }, 
    {
        "column1": "Test User 2",
        "column2": "9/19/2016, 1:21:34 PM",
        "column3": "9/21/2016, 1:21:34 PM",
        "column4": "1",
        "column5": "001",
        "column6": "100",
        "column7": "10",
        "column8": "10",
        "column9": "10",
        "column10": "10",
        "column11": "10",
        "column12": "10"
    },
    {
        "column1": "Test User 3",
        "column2": "9/12/2016, 1:21:34 PM",
        "column3": "9/13/2016, 1:21:34 PM",
        "column4": "2",
        "column5": "002",
        "column6": "200",
        "column7": "20",
        "column8": "20",
        "column9": "20",
        "column10": "20",
        "column11": "20",
        "column12": "20"
    }]
};
var columnStructure =
    [
        {
            displayName: 'Column 1',
            exportDisplayName: 'Column 1',
            name: 'column1',
            filter: 'column1',
            hidden: false
        }, {
            displayName: 'Column 2',
            exportDisplayName: 'Column 2',
            name: 'column2',
            hidden: false
        }, {
            displayName: 'Column 3',
            exportDisplayName: 'Column 3',
            name: 'column3',
            hidden: false
        }, {
            displayName: 'Column 4',
            exportDisplayName: 'Column 5',
            name: 'column4',
            hidden: false
        }, {
            displayName: 'Column 5',
            exportDisplayName: 'Partner Id',
            name: 'column5',
            hidden: false
        }, {
            displayName: 'Column 6',
            exportDisplayName: 'Column 6',
            name: 'column6',
            hidden: false
        }, {
            displayName: 'Column 7',
            exportDisplayName: 'Column 7',
            name: 'column7',
            hidden: false
        }, {
            displayName: 'Column 8',
            exportDisplayName: 'Column 8',
            name: 'column8',
            hidden: false
        }, {
            displayName: 'Column 9',
            exportDisplayName: 'Column 9',
            name: 'column9',
            hidden: false
        }, {
            displayName: 'Column 10',
            exportDisplayName: 'Column 10',
            name: 'column10',
            hidden: false
        }, {
            displayName: 'Column 11',
            exportDisplayName: 'Column 11',
            name: 'column11',
            hidden: false
        }, {
            displayName: 'Column 12',
            exportDisplayName: 'Column 12',
            name: 'column12',
            hidden: false
        }
    ];
//column width
var wsColumns =
    [
        {
            wpx: 200
        }, {
            wpx: 200
        }, {
            wpx: 200
        }, {
            wpx: 200
        }, {
            wpx: 200
        }, {
            wpx: 200
        }, {
            wpx: 200
        }, {
            wpx: 200
        }, {
            wpx: 200
        }, {
            wpx: 200
        }, {
            wpx: 200
        }, {
            wpx: 200
        }
    ];

function createWorkbook(columnNames, data) {
    var numberOfColumns = 0,
        numberofRows = 0,
        ref = {},
        i = 0,
        m = 0,
        j, k, l, cellNumber,
        lastCell = '',
        inputValue = 0,
        formatString = "",
        numConvertNotRequired = ['column2', 'column3'],
        sheetData = {};
    if (columnNames !== undefined) {
        numberOfColumns = columnNames.length;
    }
    if (data !== undefined && data.userDetails.length) {
        numberofRows = data.userDetails.length;
    }

    lastCell = getLastCell(numberOfColumns, numberofRows);
    ref = "A1:" + lastCell;

    //Create excel sheet format for the export
    sheetData = {
        "!ref": ref,
        "!cols": wsColumns
    };
    //Go through column length and prepare excel sheet for exporting.
    for (i = 0, m = 0; i < numberOfColumns; i++) {
        if (i >= 26) {
            sheetData['A' + String.fromCharCode(65 + m) + 1] = {
                "h": columnNames[i].exportDisplayName,
                "r": columnNames[i].exportDisplayName,
                "t": "s",
                "v": columnNames[i].exportDisplayName,
                "w": columnNames[i].exportDisplayName
            };
            m++;
        } else {
            sheetData[String.fromCharCode(65 + i) + 1] = {
                "h": columnNames[i].exportDisplayName,
                "r": columnNames[i].exportDisplayName,
                "t": "s",
                "v": columnNames[i].exportDisplayName,
                "w": columnNames[i].exportDisplayName
            };
        }
    }
    //inserting excel data contents
    for (j = 0, cellNumber = 2; j < data.userDetails.length; j++, cellNumber++) {
        var gridRow = data.userDetails[j];
        var isAlphabetCompleted = false;
        for (k = 0, l = 0; k < columnNames.length; k++, l++) {
            //Check if l has reached 26, if yes then reset so that again it will start from A.
            if (l === 26) {
                l = 0;
                isAlphabetCompleted = true;
            }
            var columnValueField = columnNames[k].name;
            if (columnValueField !== undefined) {
                if (gridRow[columnValueField] === null || gridRow[columnValueField] === undefined) {
                    gridRow[columnValueField] = "";
                }
                if (gridRow[columnValueField] !== "") {
                    if ((numConvertNotRequired.indexOf(columnValueField) === -1) && !isNaN(parseInt(gridRow[columnValueField]))) {
                        inputValue = parseInt(gridRow[columnValueField]);
                        formatString = '#,##0';
                        if (isAlphabetCompleted) {
                            sheetData["A" + String.fromCharCode(65 + l) + cellNumber] = {
                                "t": 'n',
                                "v": inputValue,
                                "z": formatString
                            };
                        } else {
                            sheetData[String.fromCharCode(65 + l) + cellNumber] = {
                                "t": 'n',
                                "v": inputValue,
                                "z": formatString
                            };
                        }
                    } else {
                        if (isAlphabetCompleted) {
                            sheetData["A" + String.fromCharCode(65 + l) + cellNumber] = {
                                "h": gridRow[columnValueField],
                                "r": gridRow[columnValueField],
                                "t": "s",
                                "v": gridRow[columnValueField],
                                "w": gridRow[columnValueField]
                            };
                        } else {
                            sheetData[String.fromCharCode(65 + l) + cellNumber] = {
                                "h": gridRow[columnValueField],
                                "r": gridRow[columnValueField],
                                "t": "s",
                                "v": gridRow[columnValueField],
                                "w": gridRow[columnValueField]
                            };
                        }
                    }
                }
            }
        }
    }
    return sheetData;
}

function createExcelSheet(workbook, fileName) {
    /* bookType can be 'xlsx' or 'xlsm' or 'xlsb' */
    // creating the workbook options to be used for excel sheet.
    var wopts = {
        bookType: 'xlsx',
        bookSST: false,
        type: 'binary'
    };
    var wbout = XLSX.write(workbook, wopts);
    /* the saveAs call downloads a file on the local machine */
    saveAs(new Blob([s2ab(wbout)], {
        type: ""
    }), fileName);
}

function s2ab(s) {
    var buf = new ArrayBuffer(s.length);
    var view = new Uint8Array(buf);
    for (var i = 0; i !== s.length; ++i) {
        view[i] = s.charCodeAt(i) & 0xFF;
    }
    return buf;
};

function getLastCell(columnSize, rowSize) {
    var extraxColumnSize = -1;
    var firstCellAscii = "A".charCodeAt(0);
    var lastCellAscii = -1;
    var lastCellAlphabet = "";
    extraxColumnSize = columnSize - 26;
    // need to create a more generic logic. supported number of columns are from
    // A to AZ
    if (extraxColumnSize > 0) {
        lastCellAscii = firstCellAscii + extraxColumnSize - 1;
        lastCellAlphabet = "A" + String.fromCharCode(lastCellAscii);
    } else {
        lastCellAscii = firstCellAscii + columnSize - 1;
        lastCellAlphabet = String.fromCharCode(lastCellAscii);
    }
    return lastCellAlphabet + (rowSize + 1);
};