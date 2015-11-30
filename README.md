# gulp-i18n-excel2json
> Export Excel files (XLSX/XLS) to json files.

* Manage multiple or single json output file.
* Manage dynamic files path.
* Manage nested i18n keys.

## examples
Format of excel file :

| Key     | fr    | de |
| --------|---------|-------|
|a.b1	|value-fr-a.b1|	value-de-a.b1|
|a.b2	|value-fr-a.b2|	value-de-a.b2|
|b	|value-fr-b|	value-de-b|
|c	|value-fr-c|	value-de-c|
|d.z.y.a1|	value-fr-d.z.y.a1|	value-de-d.z.y.a1|
|d.z.y.a2|	value-fr-d.z.y.a2|	value-de-d.z.y.a2|

### output in single json by lang
fr.json :
```javascript
{
    "a": {
        "b1": "value-fr-a.b1",
        "b2": "value-fr-a.b2"
    },
    "b": "value-fr-b",
    "c": "value-fr-c",
    "d": {
        "z": {
            "y": {
                "a1": "value-fr-d.z.y.a1",
                "a2": "value-fr-d.z.y.a2"
            }
        }
    }
}
```
de.json :
```javascript
{
    "a": {
        "b1": "value-de-a.b1",
        "b2": "value-de-a.b2"
    },
    "b": "value-de-b",
    "c": "value-de-c",
    "d": {
        "z": {
            "y": {
                "a1": "value-de-d.z.y.a1",
                "a2": "value-de-d.z.y.a2"
            }
        }
    }
}
```
### output in multiple json by lang and namespaces : 
fr-a.json :
```javascript
{
    "a": {
        "b1": "value-fr-a.b1",
        "b2": "value-fr-a.b2"
    }
}
```
...

## Usage
First, install `gulp-i18n-excel2json` as a development dependency:

```shell
> npm install --save-dev gulp-i18n-excel2json
```

Then, add it to your `gulpfile.js`:

```javascript
var i18nExcel2json = require('gulp-i18n-excel2json');

gulp.task('i18n', function() {
    gulp.src('config/**.xlsx')
        .pipe(i18nExcel2json({
            destFile : '__lng__/translation.__ns__.json',
            readable: true,
            colKey: 'A',
            colValArray: ['B', 'C'],
            rowStart: 2,
            rowHeader: 1
        }))
        .pipe(gulp.dest('build'))
});
```


## API

### i18n-excel2json([options])

#### options.destFile
Type: `string`

Default: `locales/__lng__/__ns__.json`

The filenames path of output.

`__lng__` : replaced by current lang

`__ns__`: replace by current namespace (each top level of i18n keys)

#### options.readable
Type: `boolean`

Default: `true`

Output human-readable json files (multiple lines).

#### options.colKey
Type: `string`

Default: `A`

The column name from excel file representing i18n keys.

#### options.colValArray
Type: `array[string]`

Default: `['B']`

List of excel columns to output each as a language.

#### options.rowStart
Type: `number`

Default: `2`

Start to output json after the specified excel line.

#### options.rowHeader
Type: `number`
Default: `1`

Excel line representing the header with the lang key for each translation.

## License
MIT &copy; Kirakishin (fork from Chris)
