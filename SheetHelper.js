'use strict';
const Util2 = require('./Util2');

const SheetHelper = {};

/**
 * Converts letters to a number (Excel spreadsheet style):
 * @example convertLetterToNumber('a') // '1'
 * @example convertLetterToNumber('aa') // '27'
 * @returns {Integer} Number code
 */
const convertLetterToNumber = (letterStr, startTotal) => {
    startTotal = startTotal || 0;
    let value = letterStr.toLowerCase().charCodeAt(0) - 96;
    if (letterStr.length === 1) {
        return startTotal + value;
    }
    value = Math.pow(26, letterStr.length - 1) * value;
    return startTotal + convertLetterToNumber(letterStr.slice(1), value);
};
/**
 * Converts numbers to a letter (Excel spreadsheet style):
 * @example convertNumberToLetter(1) // 'a'
 * @example convertNumberToLetter(27) // 'aa'
 * @example convertNumberToLetter(28) // 'ab'
 * @returns {String} Letter representation
 */
const convertNumberToLetter = number => {
    const base26Array = number.toString(26).split('').map(i => parseInt(i, 26)); // [1, 1, 0]
    //convert from [1,1,0] to [26,26]
    for (let i = base26Array.length + 1; i >= 0; i--) {
        if (base26Array[i] === 0 && base26Array[i-1] === 1) {
            base26Array[i] = '26';
            base26Array[i-1] = 0;
        }
    }
    if (base26Array[0] === 0) {
        base26Array.shift();
    }
    return base26Array.map(i => String.fromCharCode(96 + i)).join('');
};
/**
 * Converts R1C1 to A1:
 * @example convertR1C1('R2C[-1]', 'D30') // 'C$2'
 * @returns {String} A1 cell reference
 */
const convertR1C1 = (r1c1Str, currentCell) => {
    r1c1Str = r1c1Str.toLowerCase();
    currentCell = currentCell.toLowerCase();

    const rValue = r1c1Str.split('c')[0].slice(1);
    const cValue = r1c1Str.split('c')[1];
    const currentRow = currentCell.match(/\d+/)[0];
    const currentColumn = convertLetterToNumber(currentCell.match(/[a-z]+/)[0]);
    let convertedRow;
    let convertedColumn;
    let absoluteColumn = false;

    if (parseInt(rValue)) {
        convertedRow = '$'+parseInt(rValue); //absolute ref
    } else if (parseInt(rValue.substring(1, rValue.length - 1))) {
        convertedRow = parseInt(rValue.substring(1, rValue.length - 1)) + currentRow; //relative ref
    } else {
        convertedRow = currentRow; //relative 0 ref
    }

    if (parseInt(cValue)) {
        absoluteColumn = true;
        convertedColumn = parseInt(cValue); //absolute ref
    } else if (parseInt(cValue.substring(1, cValue.length - 1))) {
        convertedColumn = parseInt(cValue.substring(1, cValue.length - 1)) + currentColumn; //relative ref
    } else {
        convertedColumn = currentColumn; //relative 0 ref
    }
    convertedColumn = absoluteColumn ? '$' + convertNumberToLetter(convertedColumn) : convertNumberToLetter(convertedColumn);

    return (convertedColumn + convertedRow).toUpperCase();
};
/**
 * Converts all R1C1 references in a row to A1:
 * @example convertR1C1Row({'1': R2C[-1]'}, 'D30'}) // {'1': 'C$2'}
 * @returns {Object} Converted row
 */
const convertR1C1Row = (row, rowNumber) => {
    const newRow = {};
    for (let key in row) {
        if (row.hasOwnProperty(key)) {
            const rowValue = row[key] + '';
            const absoluteRegExp = /R(\[-?\d+]|\d*)C(\[-?\d+]|\d*)/ig; //tests for absolute cell references

            newRow[key] = rowValue.replace(absoluteRegExp, function(match) {
                return convertR1C1(match, convertNumberToLetter(parseInt(key))+rowNumber);
            });
        }
    }
    return newRow;
};
/**
 * Given 2 google spreadsheet rows, extrapolate the next row
 *
 * @example extrapolateRow({'1': '24', '2': '=B30'}, {'1': '30', '2': '=B31'}) // {'1': '36', '2': '=B32'}
 * @param {Object} row1
 * @param {Object} row2
 * @param {Integer} [row1Number] Used to convert from r1c1 to a1 cells for relative rows
 * @param {Integer} [row2Number] Used to convert from r1c1 to a1 cells for relative rows
 * @returns {Object} Extrapolated row
 */
SheetHelper.extrapolateRow = (row1, row2, row1Number, row2Number) => {
    const newRow = {};
    const cellRegExp = /\$?[a-z]+\$?\d+(?![a-z])/ig; //test for a1 cell references
    //converts all r1c1 references to a1
    if (row1Number && row2Number) {
        row1 = convertR1C1Row(row1, row1Number);
        row2 = convertR1C1Row(row2, row2Number);
    }
    for (let key in row1) {
        if (row1.hasOwnProperty(key) && row2.hasOwnProperty(key) ) {
            let row1Value = row1[key];
            let row2Value = row2[key];
            if (Util2.isString(row1Value) && parseFloat(row1Value[0])) {
                row1Value = row1Value.replace(',', '');
            }
            if (Util2.isString(row2Value) && parseFloat(row2Value[0])) {
                row2Value = row2Value.replace(',', '');
            }
            if (row1Value === row2Value) {
                newRow[key] = row2Value;
            } else if (Date.parse(row1Value) && Date.parse(row2Value)) {
                //extrapolate date
                const date1 = new Date(Date.parse(row1Value));
                const date2 = new Date(Date.parse(row2Value));
                const yearDiff = date2.getFullYear() - date1.getFullYear();
                const monthDiff = date2.getMonth() - date1.getMonth();
                const dayDiff = date2.getDate() - date1.getDate();
                const totalMonthDiff = yearDiff * 12 + monthDiff;

                if (totalMonthDiff !== 0) {
                    date2.setMonth(date2.getMonth() + totalMonthDiff);
                } else {
                    date2.setDate(date2.getDate() + dayDiff);
                }
                newRow[key] = Util2.sprintf('%d/%d/%d', date2.getMonth() + 1, date2.getDate(), date2.getFullYear());
            } else if (parseFloat(row1Value) && parseFloat(row2Value)) {
                //extrapolate number
                const num1 = parseFloat(row1Value);
                const num2 = parseFloat(row2Value);
                const numDiff = num2 - num1;
                newRow[key] = num2 + numDiff + '';
            } else if (Util2.isString(row1Value) && Util2.isString(row2Value) && row1Value.match(cellRegExp) && row2Value.match(cellRegExp)) {
                //extrapolate relative a1 cell
                const matches = [];
                let newRow2Value = row2Value;
                let match;
                while (match = cellRegExp.exec(row2Value)) {
                    matches.push(match);
                }
                cellRegExp.lastIndex = 0;

                const replacePool = {};
                for (let i = 0; i < matches.length; i++) {
                    const a1Cell = matches[i][0]; //i.e. DA$30
                    let rowStr = a1Cell.replace(/[a-z]*/ig, ''); //i.e. $30
                    if (rowStr[0] !== '$') {
                        rowStr = parseInt(rowStr) + 1 + '';
                    }
                    const newCell = a1Cell.replace(/\$?\d+/, rowStr);

                    replacePool[a1Cell] = newCell;
                }
                newRow[key] = newRow2Value.replace(new RegExp(Object.keys(replacePool).join("|"),"gi"), match => {
                    return replacePool[match];
                });
            } else {
                newRow[key] = row2Value;
            }
        }
    }

    return newRow;
};

module.exports = SheetHelper;