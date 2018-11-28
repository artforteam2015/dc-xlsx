'use strict';

var XLSX = require('xlsx');
var eachAsync = require('each-async');
var _ = {
    defaults: require('lodash.defaults'),
    map: require('lodash.map')
};
var table_fmt = {
    0:  'General',
    1:  '0',
    2:  '0.00',
    3:  '#,##0',
    4:  '#,##0.00',
    9:  '0%',
    10: '0.00%',
    11: '0.00E+00',
    12: '# ?/?',
    13: '# ??/??',
    14: 'm/d/yy',
    15: 'd-mmm-yy',
    16: 'd-mmm',
    17: 'mmm-yy',
    18: 'h:mm AM/PM',
    19: 'h:mm:ss AM/PM',
    20: 'h:mm',
    21: 'h:mm:ss',
    22: 'm/d/yy h:mm',
    37: '#,##0 ;(#,##0)',
    38: '#,##0 ;[Red](#,##0)',
    39: '#,##0.00;(#,##0.00)',
    40: '#,##0.00;[Red](#,##0.00)',
    45: 'mm:ss',
    46: '[h]:mm:ss',
    47: 'mmss.0',
    48: '##0.0E+0',
    49: '@',
    56: '"上午/下午 "hh"時"mm"分"ss"秒 "',
    65535: 'General'
};

var wsCount = []
var wscols = [];


function getCustomSSF() { return table_fmt; };
function datenum(v, date1904) {
    if (date1904) v += 1462;
    var epoch = Date.parse(v);
    return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
}


function sheet_from_array_of_arrays(data) {
    let counter = 165;
    var ws = {};
    var range = {
        s: {
            c: 10000000,
            r: 10000000
        },
        e: {
            c: 0,
            r: 0
        }
    };
    for (var R = 0; R !== data.length; ++R) {

        let z = ''

        for (var C = 0; C !== data[R].length; ++C) {
            if (typeof(data[R][C]) === "object") {
                if (Object.keys(data[R][C]).length > 1) {
                    z = data[R][C].z
                    data[R][C] = data[R][C].v;
                }
            }
            if (wsCount[C] === undefined) wsCount[C] = {len: 0, val: 0}

            if (data[R][C] !== undefined && data[R][C] !== ' ') {
                wsCount[C].len++
                wsCount[C].val += ('' + data[R][C]).length
            }

            if (range.s.r > R) range.s.r = R;
            if (range.s.c > C) range.s.c = C;
            if (range.e.r < R) range.e.r = R;
            if (range.e.c < C) range.e.c = C;

            var cell = {
                v: data[R][C]
            };
            if (cell.v === null) continue;
            var cell_ref = XLSX.utils.encode_cell({
                c: C,
                r: R
            });

            if (typeof cell.v === 'number') cell.t = 'n';
            else if (typeof cell.v === 'boolean') cell.t = 'b';
            else if (cell.v instanceof Date) {
                cell.t = 'n';
                cell.z = XLSX.SSF._table[14];
                cell.v = datenum(cell.v);
            } else cell.t = 's';

            if (z !== '') {
                cell.z = z;
                let found = false
                let keys = Object.keys(table_fmt)
                keys.forEach((key)=>{
                    if (table_fmt[key] === z) found = true
                })
                if (!found) {
                    table_fmt[counter] = z;
                }
                z = ''
                counter++;
            }
            ws[cell_ref] = cell;
        }
    }
    if (range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);

    for (let i in wsCount) {

        let srLen = wsCount[i].len
        let srVal = wsCount[i].val
        wscols[i] = {wch: srVal / srLen + 2}
    }
    ws['!cols'] = wscols;
    return ws;
}

function Workbook() {
    if (!(this instanceof Workbook)) return new Workbook();
    this.SheetNames = [];
    this.Sheets = {};
}

module.exports = {
    parse: function(mixed, options) {
        var ws;
        if (typeof mixed === 'string') ws = XLSX.readFile(mixed, options);
        else ws = XLSX.read(mixed, options);
        return _.map(ws.Sheets, function(sheet, name) {
            return {
                name: name,
                data: XLSX.utils.sheet_to_json(sheet, {
                    header: 1,
                    raw: true
                })
            };
        });
    },
    parseFileAsync: function(mixed, options, callback) {
        var ws;
        if (typeof mixed === 'string') ws = XLSX.readFile(mixed, options);
        else ws = XLSX.read(mixed, options);
        callback(_.map(ws.Sheets, function(sheet, name) {
            return {
                name: name,
                data: XLSX.utils.sheet_to_json(sheet, {
                    header: 1,
                    raw: true
                })
            };
        }));
    },
    build: function(array, options) {
        var defaults = {
            bookType: 'xlsx',
            bookSST: false,
            type: 'binary'
        };
        var wb = new Workbook();
        array.forEach(function(worksheet) {
            var name = worksheet.name || 'Sheet';
            var data = sheet_from_array_of_arrays(worksheet.data || []);
            wb.SheetNames.push(name);
            wb.Sheets[name] = data;
        });
        var data = XLSX.write(wb, _.defaults(options || {}, defaults));
        if (!data) return false;
        var buffer = new Buffer(data, 'binary');
        return buffer;
    },
    buildAsync: function(array, options, callback) {
        var defaults = {
            bookType: 'xlsx',
            bookSST: false,
            type: 'binary'
        };
        var wb = new Workbook();
        eachAsync(array, function(worksheet, index, done) {
            var name = worksheet.name || 'Sheet';
            var data = sheet_from_array_of_arrays(worksheet.data || []);
            wb.SheetNames.push(name);
            wb.Sheets[name] = data;
            wb.SSF = getCustomSSF();
            done();
        }, function(error){
            var data = XLSX.write(wb, _.defaults(options || {}, defaults));
            if (!data){
                doneBuilding("no data", null);
            }
            else{
                var buffer = new Buffer(data, 'binary');
                doneBuilding(null, buffer);
            }
        });

        function doneBuilding(err, buffer){
            callback(err, buffer);
        }
    }
};
