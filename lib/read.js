'use strict';
const ExcelJS = require('exceljs');
const _ = require('lodash');
const PS = require('parameter-type');

class ReadExcel {
    constructor() {
        this.WORKBOOK = new ExcelJS.Workbook();
    }

    async fromStream(stream) {
        await this.WORKBOOK.xlsx.read(stream);
    }

    async fromFile(filepath) {
        await this.WORKBOOK.xlsx.readFile(filepath);
        return this;
    }

    async fromBuffer(buffer) {
        await this.WORKBOOK.xlsx.load(buffer);
    }

    toObject(obj, skipFirstRow = true) {
        const sheets = Object.keys(obj);

        const sheetArray = [];
        for (let i = 0; i < sheets.length; i++) {
            const workSheet = this.WORKBOOK.worksheets[i];
            if(!workSheet) continue;

            const sheet = [];
            workSheet.eachRow((row, rowIndex) => {

                if (skipFirstRow && rowIndex == 1) return;

                const cells = row.values;
                cells.shift(0);

                const target = obj[sheets[i]];
                const targetKeys = Object.keys(target);

                const targetObj = {};
                for (let j = 0; j < targetKeys.length; j++) {
                    const condition = target[targetKeys[j]];
                    const v = this.converter(targetKeys[j], rowIndex, cells[j], condition?.optional, condition?.validator, condition?.converter);
                    _.set(targetObj, targetKeys[j], v);
                }

                sheet.push(targetObj);
            });

            sheetArray.push(sheet)
        }

        return sheetArray;
    }

    converter(key, rowIndex, value, isRequired, validator, converter) {
        if (!value && !isRequired) throw new ReferenceError(`${key} can not be null,Row: ${rowIndex}.`);

        if (PS.isObject(value) && value.hyperlink) {
            value = value.text;
        }

        if (PS.isFunction(validator) && !isRequired) {
            if (!validator.call(null, value)) throw new TypeError(`${key} format Error.Row: ${rowIndex}`);
        }

        if (PS.isFunction(converter) && !isRequired) {
            return converter.call(null, value);
        }

        return value;
    }
}

module.exports = ReadExcel;