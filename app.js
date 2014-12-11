var XLSX = require('xlsx');

var url = '/Users/xuanzhi.zhang/Documents/forJSON.xlsx';

var workbook = XLSX.readFile(url);
var sheetNames = workbook.SheetNames;

var sheet1 = workbook.Sheets[sheetNames[0]];
var range = 'A33:A33';

var arrOfArr = getArrayOfArrayByRange(sheet1, range);
 var JSONdata = arrayOfArrayToJSON(arrOfArr);


 function getArrayOfArrayByRange(sheet, rawRange) {
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

 function arrayOfArrayToJSON(arrOfArr) {
     var json = {};
     for (var rowI = 0, rowL = arrOfArr.length; rowI < rowL; rowI += 1) {
         var rowArr = arrOfArr[rowI];
         for (var colI = 0, colL = rowArr.length; colI < colL; colI += 1) {
             var value = rowArr[colI];
             if (colI === 0 && value === undefined) {
                 // TODO: 行首为空
             } else {
                 var typeAndValueArr = splitKey(value);
                 var finalObj = generateKey(json, typeAndValueArr);
                 finalObj[typeAndValueArr[typeAndValueArr.length - 1].value] = rowArr[colI + 1];
                 colI += 1;
                 // TODO
             }
         }
     }
     console.log(JSON.stringify(json, null, '  '));
 }

 /**
  * 将 a[1].b[3][4].c 转化为 'a',1,'b',3,4,'c', 并标明是 objKey 还是 arrIndex
  */
 function splitKey(complexKey) {
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
             typeAndValueArr.push({ type: 'objKey', value: propertyArr[i] })            
         }
     }
     return typeAndValueArr;
 }

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