var fs = require('fs');
var XLSX = require('xlsx');

function xlsx2json(config, cb) {
    var jsonArr = [];
    var writeSuccessArr = [];
    var writeFailedArr = [];
    var workbooks = {};
    for (var i = 0, length = config.length; i < length; i += 1) {
        var job = config[i];
        if ( !(job.input && job.sheet) ) {
            var errMsg = 'A task must contain "input" and "sheet". In\n';
            errMsg += JSON.stringify(job, null, '  ');
            cb(new Error(errMsg));
            return;
        }
        if (!workbooks[job.input]) {
            try {
                workbooks[job.input] = XLSX.readFile(job.input);
            } catch(e) {
                cb(new Error('XLSX parse error: ' + job.input));
                return;
            }
        }
    }
    for (var i = 0, length = config.length; i < length; i += 1) {
        var job = config[i];
        var workbook = workbooks[job.input];    // var sheetNames = workbook.SheetNames;
        var workSheet = workbook.Sheets[job.sheet];

        var arrOfArr = getArrayOfArrayByRange(workSheet, job.range);
        var json = arrOfArr;
        if (!job.raw) {
            json = arrayOfArrayToJSON(arrOfArr);
        }
        if (job.output) {
            writeJSON(json, job.output, length, writeSuccessArr, writeFailedArr);
        }
        jsonArr.push(json);
    }
    cb(null, jsonArr);
}
module.exports = xlsx2json;

/**
 * 获取 sheet 中指定 range 的 行列
 */
function getArrayOfArrayByRange(sheet, rawRange) {
    rawRange = rawRange || sheet['!ref'];
    var arrOfArr = [];
    var range = XLSX.utils.decode_range(rawRange);
    for (var row = range.s.r; row <= range.e.r; row += 1) {
        var rowArr = []
        for (var col = range.s.c; col <= range.e.c; col += 1) {
            var addr = XLSX.utils.encode_cell({ c: col, r: row});
            var value = sheet[addr] && sheet[addr].v;    // sheet[addr] may be undefined
            rowArr.push(value);
        }
        arrOfArr.push(rowArr);
    }
    return arrOfArr;
}

/**
 * 核心函数, 功能如其名
 */
function arrayOfArrayToJSON(arrOfArr) {
    var json = {};
    var tableHead = [];
    for (var rowI = 0, rowL = arrOfArr.length; rowI < rowL; rowI += 1) {
        var rowArr = arrOfArr[rowI];
        for (var colI = 0, colL = rowArr.length; colI < colL; colI += 1) {
            var value = rowArr[colI];
            if (value === undefined) {
                // 行首为空白
                if (colI === 0) {
                    if (rowArr[colI + 1] === undefined) {
                        // 行首连续两个空白, 清空 tableHead
                        tableHead = [];
                    } else {
                        // 行首只有一个空白
                        for (colI = colI + 1; colI < colL; colI += 1) {
                            if (rowArr[colI] === undefined) {
                                break;
                            }
                            tableHead.push(rowArr[colI]);
                        }
                    }
                } else {
                    continue;
                }
            } else {
                // 行首非空白
                if (tableHead.length === 0) {
                    var typeAndValueArr = splitKey(value);
                    var finalObj = generateKey(json, typeAndValueArr);
                    var finalKey = typeAndValueArr[typeAndValueArr.length - 1].value;
                    if (rowArr[colI + 1] !== '[]') {
                        finalObj[finalKey] = rowArr[colI + 1];
                        colI += 1;
                    } else {
                        finalObj[finalKey] = [];
                        colI += 1;
                        while (rowArr[colI + 1] !== undefined) {
                            colI += 1;
                            finalObj[finalKey].push(rowArr[colI]);
                        }
                    }
                } else {
                    for (var tableI = 0, tableL = tableHead.length; tableI < tableL; tableI += 1) {
                        var concatKey = value;
                        if (tableHead[0] !== '[') {
                            concatKey += '.';
                        }
                        concatKey += tableHead[tableI];
                        var typeAndValueArr = splitKey(concatKey);
                        var finalObj = generateKey(json, typeAndValueArr);
                        var finalKey = typeAndValueArr[typeAndValueArr.length - 1].value;
                        finalObj[finalKey] = rowArr[colI + 1 + tableI];
                    }
                    colI += tableHead.length;
                }
            }
        }
    }
    return json;
}

/**
 * 将 a[1].b[3][4].c 转化为 'a',1,'b',3,4,'c', 并标明是 objKey 还是 arrIndex
 */
function splitKey(complexKey) {
    complexKey = complexKey.toString();
    var typeAndValueArr = [];
    var propertyArr = complexKey.split('.');
    for (var i = 0, length = propertyArr.length; i < length; i += 1) {
        var singleKey = propertyArr[i];
        var leftBracketIndex = singleKey.indexOf('[');
        if (leftBracketIndex > 0) {
            var keyBeforeBracket = singleKey.slice(0, leftBracketIndex);
            propertyArr[i] = keyBeforeBracket;
            typeAndValueArr.push({ type: 'objKey', value: propertyArr[i] })
            propertyArr.splice(i + 1, 0, singleKey.slice(leftBracketIndex));
            length += 1;
        } else if (leftBracketIndex === 0) {
            var rightBrachetIndex = singleKey.indexOf(']');
            if (rightBrachetIndex < 0) {
                throw new Error('Split Key error: ' + key);
            }
            propertyArr[i] = singleKey.slice(1, rightBrachetIndex);
            propertyArr[i] = parseInt(propertyArr[i]);
            typeAndValueArr.push({ type: 'arrIndex', value: propertyArr[i] });
            var keyAfterBracket = singleKey.slice(rightBrachetIndex + 1);
            if (keyAfterBracket !== '') {
                propertyArr.splice(i + 1, 0, keyAfterBracket);
                length += 1;
            }
        } else {
            typeAndValueArr.push({ type: 'objKey', value: propertyArr[i] });
        }
    }
    return typeAndValueArr;
}

/**
 * 根据 typeAndValueArr 为 obj 生成没有的 key, 并返回最低级别的 obj 属性
 */
function generateKey(obj, typeAndValueArr) {
    for (var i = 0, length = typeAndValueArr.length; i < length - 1; i += 1) {
        var type = typeAndValueArr[i].type;
        var value = typeAndValueArr[i].value;
        var nextType = typeAndValueArr[i + 1].type;
        var nextValue = typeAndValueArr[i + 1].value;
        if (nextType === 'arrIndex') {
            if (obj[value] === undefined) {
                obj[value] = [];
            }
            if (!Array.isArray(obj[value])) {
                // TODO: 强化这个 error message
                throw new Error(value + ' is not an array')
            }
        } else if (nextType === 'objKey') {
            if (obj[value] === undefined) {
                obj[value] = {};
            }
            if (typeof obj[value] !== 'object' || Array.isArray(obj[value])) {
                // TODO: 强化这个 error message
                throw new Error(value + ' is not an object, maybe an array');
            }
        }
        obj = obj[value];
    }
    return obj;
}

function writeJSON(json, filename, length, successArr, failedArr) {
    var jsonStr = JSON.stringify(json, null, '    ');
    fs.writeFile(filename, jsonStr, 'utf8', function(err) {
        if (err) {
            console.log(err);
            failedArr.push(filename);
            return;
        }
        successArr.push(filename);
        if (successArr.length + failedArr.length === length) {
            console.log('Successful writing:');
            console.log(JSON.stringify(successArr, null, '    '));
            console.log('Failing writing:');
            console.log(JSON.stringify(failedArr, null, '    '));
        }
    });
}
