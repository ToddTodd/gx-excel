'use strict';

const assert = require('assert');

const _ = require('lodash');
const ExcelJS = require('exceljs');
const numeral = require('numeral');

class WriteExcel {
    constructor(props) {
        this.WORKBOOK = new ExcelJS.Workbook();
        if (props && Object.keys(props).length) {
            this.WORKBOOK.model = props;
        }
    }

    addSheet(sheetName, headers, values, props) {
        assert(sheetName, 'Sheet name is required.');
        assert(Array.isArray(headers), 'Sheet header is required.');
        assert(Array.isArray(values), 'Fill value is required.');

        assert(headers.length, 'Sheet header can not be empty array.');
        assert(_.every(headers, x => x.header && x.key), 'Key or header is required.')

        const sheet = this.WORKBOOK.addWorksheet(sheetName, props);
        sheet.columns = headers;

        for (const v of values) {
            const row = [];
            for (const h of headers) {
                let value = _.get(v, h.key);
                if ((void 0) === value) throw new ReferenceError(`key ${h.key} not found`);

                value = this.format(value, h?.type, h?.defaultValue, h?.converter);
                row.push(value);
            }
            sheet.addRow(row);
        }

        return this;
    }

    addSheets(obj) {
        for (const sheet of Object.keys(obj)) {
            const { headers, values, props } = obj[sheet];
            this.addSheet(sheet, headers, values, props);
        }

        return this;
    }

    async toBuffer() {
        const buffer = await this.WORKBOOK.xlsx.writeBuffer();
        return buffer;
    }

    async toFile(filename) {
        await this.WORKBOOK.xlsx.writeFile(filename);
        return filename;
    }

    format(value, type, defaultValue, converter) {
        if ('function' === typeof converter) {
            return converter.call(null, value);
        }

        if (!value && defaultValue) return defaultValue;

        if (type === 'status') {
            if (!value || !value.length) throw new TypeError('Status text can not be null or empty string');

            return value.trim().replace(value[0], value[0].toUpperCase());
        }

        if (type === 'amount') {
            if (('number' !== typeof value) || Number.isNaN(value)) throw new TypeError('Amount must be number');
            return numeral(value).format("0,0.00");
        }

        return value;
    }
}

module.exports = WriteExcel;