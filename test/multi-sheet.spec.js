const assert = require('assert');
const path = require('path');
const fs = require('fs');

const numeral = require('numeral');
const _ = require('lodash');

const { WriteExcel, ReadExcel } = require('../index');

const filepath = path.join(__dirname, 'test.xlsx');

const sheets = {
    sheet1: {
        headers: [
            { header: 'Status', key: 'status', type: 'status' },
            { header: 'Amount', key: 'amount', type: 'amount' },
            { header: 'Sales', key: ':sales.name' },
            { header: 'Username', key: 'name', converter: (v) => v.toUpperCase() }
        ],
        values: [
            { status: 'activated', amount: 10000, ":sales": { name: 'sales01' }, name: 'angle' },
            { status: 'activated', amount: 20000, ":sales": { name: 'sales02' }, name: 'lili' }
        ]
    },
    sheet2: {
        headers: [
            { header: 'Status', key: 'status', type: 'status' },
            { header: 'Amount', key: 'amount', type: 'amount' },
            { header: 'Sales', key: ':sales.name' },
            { header: 'Username', key: 'name', converter: (v) => v.toUpperCase() }
        ],
        values: [
            { status: 'activated', amount: 10000, ":sales": { name: 'sales01' }, name: 'angle' },
            { status: 'activated', amount: 20000, ":sales": { name: 'sales02' }, name: 'lili' }
        ]
    }
};

const reader = {
    sheet1: {
        status: null,
        amount: { converter: (x) => numeral(x).value() },
        ":sales.name": {},
        name: {}
    },
    sheet2: {
        status: null,
        amount: { converter: (x) => numeral(x).value() },
        ":sales.name": {},
        name: {}
    }
};

describe('Write and read multi sheet', async () => {
    after(async () => {
        await fs.unlinkSync(filepath);
    });

    it('Wite excel', async () => {
        await new WriteExcel()
            .addSheets(sheets)
            .toFile(filepath);
    });

    it('Read excel', async () => {
        const Reader = await new ReadExcel()
            .fromFile(filepath);

        const array = Reader.toObject(reader);
        assert.ok(Object.keys(sheets).length == array.length)
    });
});