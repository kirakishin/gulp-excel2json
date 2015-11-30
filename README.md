# gulp-i18n-excel2json
> Excel (XLSX/XLS) to json.


## Usage
First, install `gulp-i18n-excel2json` as a development dependency:

```shell
> npm install --save-dev gulp-i18n-excel2json
```

Then, add it to your `gulpfile.js`:

```javascript
var i18n-excel2json = require('gulp-i18n-excel2json');

gulp.task('i18n', function() {
    gulp.src('config/**.xlsx')
        .pipe(excel2json({
            destFile : '__lng__/operateurTest.__ns__-__lng__.json',
            pretty: true,
            colKey: 'A',
            colValArray: ['B', 'C'],
        }))
        .pipe(gulp.dest('build'))
});
```


## API

### i18n-excel2json([options])

#### options.headRow
Type: `number`

Default: `1`

The row number of head. (Start from 1).

#### options.valueRowStart
Type: `number`

Default: `3`

The start row number of values. (Start from 1)

#### options.trace
Type: `Boolean`

Default: `false`

Whether to log each file path while convert success.

## License
MIT &copy; Chris
