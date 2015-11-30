'use strict';
var gutil = require('gulp-util');
var through = require('through2');
var XLSX = require('xlsx');
var File = require('vinyl');


module.exports = function (options) {
    var task = this; // task is a reference to the through stream
    options = options || {};
    var savePath = 'locales/__lng__/__ns__.json';
    // give default path if resPath not provided
    if (options.destFile) {
        savePath = options.destFile;
    }
    var withNameSpaces = (savePath.indexOf('__ns__') !== -1);

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
    console.log("toJson");
    var workbook;
    if (typeof fileName === 'string') {
        workbook = XLSX.readFile(fileName);
    } else {
        workbook = fileName;
    }
    var worksheet = workbook.Sheets[workbook.SheetNames[0]];
    var namemap = [];

    // json to return
    var json = {};
    var langMapByCol = {};
    var langMapByLang = {};
    var curRow = 0;
    //console.log(worksheet);
    //console.log(worksheet['!merges']);
    //console.log(worksheet.A3.XF);
    var refToJsonNestedObj = {};
    var refToJsonNestedKey = {};
    var lastConcatNestedKey = '';
    for (var key in worksheet) {
        if (worksheet.hasOwnProperty(key)) {
            var cell = worksheet[key];
            //console.log(key+'='+cell.v);
            var match = /([A-Z]+)(\d+)/.exec(key);
            if (!match) {
                continue;
            }
            var col = match[1]; // ABCD
            var row = match[2]; // 1234
            var value = cell.v;

            if (row == rowHeader) {
                if (col !== colKey) {
                    console.log(key+'='+cell.v);
                    json[value] = {};
                    langMapByCol[col] = value;
                    langMapByLang[value] = true;
                }
            } else if (row >= rowStart) {
                console.log(langMapByCol);
                console.log(langMapByLang);
                console.log(json);
                if (col == colKey) {
                    lastConcatNestedKey = value;
                    var i18nKeyArray = value.split('.');
                    for (var oneLang in langMapByLang) {
                        if(langMapByLang.hasOwnProperty(oneLang)) {
                            console.log('oneLang=', oneLang);
                            var jsonTmp = json[oneLang];
                            //console.log('jsonTmp=', jsonTmp);
                            for (var ind in i18nKeyArray) {
                                if (i18nKeyArray.hasOwnProperty(ind)) {
                                    var indexName = i18nKeyArray[ind];
                                    if (!jsonTmp.hasOwnProperty(indexName)) {
                                        console.log('indexName=', indexName, ' jsonTmp=', jsonTmp);
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
                                console.log('currentLang='+currentLang, 'refToJsonNestedObj=', refToJsonNestedObj);
                                if (typeof refToJsonNestedObj[currentLang][refToJsonNestedKey[currentLang]] === 'object') {
                                    console.warn('ERROR', col + row + '=' + value, 'cannot be set into', '"' + lastConcatNestedKey + '"', 'ALREADY EXISTS AS OBJECT: ' + lastConcatNestedKey + '=', refToJsonNestedObj[currentLang][refToJsonNestedKey[currentLang]]);
                                } else if (refToJsonNestedObj[currentLang][refToJsonNestedKey[currentLang]] !== undefined) {
                                    console.warn('ERROR', col + row + '=' + value, 'cannot be set into', '"' + lastConcatNestedKey + '"', 'ALREADY DEFINED : ' + lastConcatNestedKey + '=', refToJsonNestedObj[currentLang][refToJsonNestedKey[currentLang]]);
                                } else {
                                    console.log('set value in ' + refToJsonNestedKey[currentLang] + ' of', refToJsonNestedObj[currentLang], 'with value', value);
                                    refToJsonNestedObj[currentLang][refToJsonNestedKey[currentLang]] = value;
                                }
                            }
                        }
                    }

                }

                //console.log(json);
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
    console.log('savePath='+savePath, writeObj);

    return new File({
        cwd: '.',
        path: savePath, // put each translation file in a folder
        contents: new Buffer(stringify(writeObj)),
    });
};
/*
var fileName = 'i18n.xls';
var options = {
    destFile: './__lng__/operateur.__ns__-__lng__.json',
    namespaces: true
}
var json = toJson(fileName, 'A', ['B', 'C'], 2, 1);
var files = [];
for (var lang in json) {
    if (json.hasOwnProperty(lang)) {
        if (options.namespaces) {
            Object.keys(json[lang]).forEach(function (ns) {
                files.push(filePath(json[lang][ns], lang, ns));
            });
        } else {
            files.push(filePath(json[lang], lang, ''));
        }
    }
}
console.log(json);
console.log(files);*/



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
        var files = [];
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
        }

        //console.log(json);

        if (options.trace) {
            console.log("convert file :" + file.path);
        }

        cb();
    });
};
