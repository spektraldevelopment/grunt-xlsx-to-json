'use strict';

var _ = require('lodash');
var xlsx = require('node-xlsx');

var prepareString = function(str) {
    var _str = String(str).trim();
    return (str == null || !_str.length) ? null : _str;
};

var parse = function(filepath, options) {
    var json = {};

    options = _.extend({
        set: function(resultJsonObject, columnName, sheetName, key, value) {
            if (!resultJsonObject.hasOwnProperty(columnName)) {
                resultJsonObject[columnName] = {};
            }

            resultJsonObject[columnName][key] = value;
        }
    }, options);

    xlsx.parse(filepath).forEach(function(sheet) {

        var
            sheetData = (sheet.data || []) || [],
            columnArray = [],
            i;

        //Skip the first empty column and the key column
        for (i = 2; i < sheetData[0].length; i += 1) {
            columnArray.push(sheetData[0][i]);
        }

        json[sheet.name] = [];

        _.each(columnArray, function(col) {
            var obj = {};
            json[sheet.name].push(obj);
        });

        _.each(sheet.data, function(row) {
            var
                slicedRow = row.slice(1, row.length),
                key = prepareString(slicedRow.shift());

            if (key != null) {
                _.each(slicedRow, function(value, index) {
                    value = prepareString(value);
                    json[sheet.name][index][key] = value;
                });
            }
        });
    });
    return json;
};

module.exports = parse;