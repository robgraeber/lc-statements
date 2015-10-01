'use strict';
const sprintf = require('sprintf-js').sprintf;
const pdf2text = require('pdf2text');

const Util2 = {};
/**
 * Create a formatted string with placeholders:
 * % — yields a literal % character
 * b — yields an integer as a binary number
 * c — yields an integer as the character with that ASCII value
 * d or i — yields an integer as a signed decimal number
 * e — yields a float using scientific notation
 * u — yields an integer as an unsigned decimal number
 * f — yields a float as is; see notes on precision above
 * g — yields a float as is; see notes on precision above
 * o — yields an integer as an octal number
 * s — yields a string as is
 * x — yields an integer as a hexadecimal number (lower-case)
 * X — yields an integer as a hexadecimal number (upper-case)
 * j — yields a JavaScript object or array as a JSON encoded string
 * @example sprintf('%d/%d/%d', date.getMonth() + 1, date.getDate(), date.getFullYear()) // 7/13/2015
 * @example sprintf('Hello %(name)s', user) // Hello Dolly
 * @returns {string} the formatted string
 */
Util2.sprintf = sprintf;

/**
 * Create an array of pages/text from a pdf:
 * @example pdf2text(pdfPath) // [['str1', 'str2', ...][...]]
 * @example pdf2text(pdfBuffer) // [['str1', 'str2', ...][...]]
 * @returns {Array} Array of Pages / string array
 */
Util2.pdf2text = pdf2text;

/**
 * Checks if object is a string:
 * @example isString('str') // true
 * @returns {Boolean}
 */
Util2.isString = obj => typeof obj === 'string' || obj instanceof String;

module.exports = Util2;