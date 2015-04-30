'use strict';

var _ = require('lodash');
var xlsx = require('node-xlsx');

var prepareString = function (str) {
	var _str = String(str).trim();

	return (str == null || !_str.length) ? null : _str;
};

var parse = function (filepath, options) {
	var json = {};

	options = _.extend({
		set: function (resultJsonObject, columnName, sheetName, key, value) {
			if (!resultJsonObject.hasOwnProperty(columnName)) {
				resultJsonObject[columnName] = {};
			}

			resultJsonObject[columnName][key] = value;
		}
	}, options);

	xlsx.parse(filepath).forEach(function (sheet) {
		var columnNames = (sheet.data || []).shift() || [];

		columnNames.shift();
		json[sheet.name] = [];

		_.each(columnNames, function(col) {
			var obj = {};
			json[sheet.name].push(obj);
		});

		_.each(sheet.data, function (row) {
			var key = prepareString(row.shift()), i;
			if (key != null) {
				_.each(row, function (value, index) {
					value = prepareString(value);
					json[sheet.name][index][key] = value;
				});	
			} 
		});
	});
	return json;
};

module.exports = parse;
