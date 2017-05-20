'use strict';
var gutil = require('gulp-util');
var through = require('through2');
var XLSX = require('xlsx');
var File = require('vinyl');
var winston = require('winston');


module.exports = function (options) {
    options = options || {};
    var savePath = 'locales/__lng__/__ns__.json';
    // give default path if resPath not provided
    if (options.destFile) {
        savePath = options.destFile;
    }
    var withNameSpaces = (savePath.indexOf('__ns__') !== -1);

    if(options.levelDebug)
    {
        winston.level = options.levelDebug;
    }

// stringifies JSON and makes it human-readable if asked
function stringify(jsonObj) {
    if (true || options.readable) {
        return JSON.stringify(jsonObj, null, 4);
    } else {
        return JSON.stringify(jsonObj);
    }
}
/**
 * excel filename or workbook to json
 * @param fileName
 * @param headRow
 * @param valueRow
 * @returns {{}} json
 */
var toJson = function (fileName, colKey, colValArray, rowStart, rowHeader) {
    winston.info("Convert toJson");
    var workbook;
    if (typeof fileName === 'string') {
        workbook = XLSX.readFile(fileName);
    } else {
        workbook = fileName;
    }
    var worksheet = workbook.Sheets[workbook.SheetNames[0]];

    // json to return
    var json = {};
    var langMapByCol = {};
    var langMapByLang = {};
    var refToJsonNestedObj = {};
    var refToJsonNestedKey = {};
    var lastConcatNestedKey = '';
    for (var key in worksheet) {
        if (worksheet.hasOwnProperty(key)) {
            var cell = worksheet[key];
            var match = /([A-Z]+)(\d+)/.exec(key);
            if (!match) {
                continue;
            }
            var col = match[1]; // ABCD
            var row = match[2]; // 1234
            var value = cell.v;

            if (row == rowHeader) {
                if (col !== colKey) {
                    winston.log('debug',key+'='+cell.v);
                    json[value] = {};
                    langMapByCol[col] = value;
                    langMapByLang[value] = true;
                }
            } else if (row >= rowStart) {
                winston.log('debug',langMapByCol);
                winston.log('debug',langMapByLang);
                winston.log('debug',json);
                if (col == colKey) {
                    lastConcatNestedKey = value;
                    var i18nKeyArray = value.split('.');
                    for (var oneLang in langMapByLang) {
                        if(langMapByLang.hasOwnProperty(oneLang)) {
                            winston.log('debug','oneLang=', oneLang);
                            var jsonTmp = json[oneLang];
                            for (var ind in i18nKeyArray) {
                                if (i18nKeyArray.hasOwnProperty(ind)) {
                                    var indexName = i18nKeyArray[ind];
                                    if (!jsonTmp.hasOwnProperty(indexName)) {
                                        winston.log('indexName=', indexName, ' jsonTmp=', jsonTmp);
                                        jsonTmp[indexName] = (ind == i18nKeyArray.length - 1 ? undefined : {});
                                    }
                                    refToJsonNestedObj[oneLang] = jsonTmp;
                                    refToJsonNestedKey[oneLang] = indexName;
                                    jsonTmp = jsonTmp[indexName];
                                }
                            }
                        }
                    }

                } else {
                    for (var oneColVal in colValArray) {
                        if (colValArray.hasOwnProperty(oneColVal)) {
                            if (col == colValArray[oneColVal]) {
                                var currentLang = langMapByCol[col];
                                winston.log('debug','currentLang='+currentLang, 'refToJsonNestedObj=', refToJsonNestedObj);
                                if (typeof refToJsonNestedObj[currentLang][refToJsonNestedKey[currentLang]] === 'object') {
                                    winston.warn('ERROR', col + row + '=' + value, 'cannot be set into', '"' + lastConcatNestedKey + '"', 'ALREADY EXISTS AS OBJECT: ' + lastConcatNestedKey + '=', refToJsonNestedObj[currentLang][refToJsonNestedKey[currentLang]]);
                                } else if (refToJsonNestedObj[currentLang][refToJsonNestedKey[currentLang]] !== undefined) {
                                    winston.warn('ERROR', col + row + '=' + value, 'cannot be set into', '"' + lastConcatNestedKey + '"', 'ALREADY DEFINED : ' + lastConcatNestedKey + '=', refToJsonNestedObj[currentLang][refToJsonNestedKey[currentLang]]);
                                } else {
                                    winston.log('set value in ' + refToJsonNestedKey[currentLang] + ' of', refToJsonNestedObj[currentLang], 'with value', value);
                                    refToJsonNestedObj[currentLang][refToJsonNestedKey[currentLang]] = value;
                                }
                            }
                        }
                    }

                }
            }
        }
    }
    return json;
};

function filePath(savePath, jsonObj, lang, key) {
    var writeObj = {};



    if (key) {
        savePath = savePath.replace(new RegExp('__ns__', 'g'), key);
        writeObj[key] = jsonObj;
    } else {
        savePath = savePath.replace(new RegExp('__ns__', 'g'), 'translation');
        writeObj = jsonObj;
    }

    savePath = savePath.replace(new RegExp('__lng__', 'g'), lang);
    winston.log('debug','savePath='+savePath, writeObj);

    return new File({
        cwd: '.',
        path: savePath, // put each translation file in a folder
        contents: new Buffer(stringify(writeObj)),
    });
};

    return through.obj(function (file, enc, cb) {
        var task = this;
        if (file.isNull()) {
            this.push(file);
            return cb();
        }

        if (file.isStream()) {
            this.emit('error', new gutil.PluginError(PLUGIN_NAME, 'Streaming not supported'));
            return cb();
        }

        var arr = [];
        for (var i = 0; i < file.contents.length; ++i) arr[i] = String.fromCharCode(file.contents[i]);
        var bString = arr.join("");

        /* Call XLSX */
        var workbook = XLSX.read(bString, {type: "binary"});

        var json = toJson(workbook, options.colKey || 'A', options.colValArray || ['B'], options.rowStart || 2, options.rowHeader || 1);
        for (var lang in json) {
            if (json.hasOwnProperty(lang)) {
                if (withNameSpaces) {
                    Object.keys(json[lang]).forEach(function (ns) {
                        task.push(filePath(savePath, json[lang][ns], lang, ns));
                    });
                } else {
                    task.push(filePath(savePath, json[lang], lang, ''));
                }
            }
        };

        if (options.trace) {
            winston.log('debug',"convert file :" + file.path);
        }

        cb();
    });
};
