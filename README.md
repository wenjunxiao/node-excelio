# excelio

 读写excel工具


## Usage

```js
const ExcelIO = require('excelio');
```

### `ExcelWriter`

  Writer excel class.
```js
const writer = new ExcelWriter();
```

### `createWriter`

  Create writer instance.

### `ExcelReader`

  Read excel class.
```js
const reader = new ExcelReader();
```

### `createReader`

  Create reader instance.

## API

### `ExcelWriter`

```js
let writer = new ExcelWrite({})
```
  支持一下选项
* `alignment` 默认对齐方式,`{horizontal: 'left',vertical: 'left'}`
* `px` 宽度是否使用像素值，如果`true`则传入的宽度都必须是像素，否则传入的宽度的是指字符数(注意中文)
* `fontSize` 默认字体大小
* `width` 默认宽度（像素还是字符数，取决于`px`）
* `minWidth` 最小宽度（像素还是字符数，取决于`px`）
* `NaN` 数字列为空时的表示符号，比如`-`
* `showGridLines` 是否显示表格线
* `titleLine` 标题行行号，比如`0`
* `titleOpts` 标题行的默认选项，`title()`时使用或者指定`titleLine`了之后匹配的行时使用
* `cellOpts` 单元格默认选项，填充单元格时自动使用
* `border2end` 是否增加边框

#### `withoutGridLines()`

  生成的excel没有网格,需要使用者自定义`border`。

#### `sheet(name)`

  创建一个`sheet`页，并将新的`sheet`页作为当前活动页，后续所有的操作都是基于该`sheet`页的操作

#### `newSheet(name)`

  创建并返回新的`sheet`页,返回的`sheet`对象支持所有的单元格的操作，参考例子[examples/multiple.js](examples/multiple.js)

#### `go(row[, cell = 0])`

  跳转到某一行到的指定列（默认未行首）

#### `skipRow(rows = 1[, cells = 0])`

  跳过指定行（默认当前行）并指定列（默认行首，如果cells小于0表示从当前列往回的列数，特别的cells=-1时，写入的内容将在同一列）,参考例子[examples/skip.js](examples/skip.js)

#### `skipCell(cells = 1)`

  跳过当前单元格

#### `rowIndex()`

  当前行位置

#### `colIndex()`

  当前列位置

#### `width(width, colIndex=-1)`

  设置指定列的宽度,默认当前列（像素还是字符数，取决于`px`）

#### `chWidth(width, colIndex=-1)`
  设置指定列的中文宽度,默认当前列，如果是像素等同于`width()`，如果是字符数，则会乘以对应的系数（像素还是字符数，取决于`px`）

#### `row(cells=0)`

  移动到下一行的某个位置,默认是行到起始位置。

#### `title(value, options={})`

  填充标题单元格，会自动应用默认选项`titleOpts`，`options`参考`cell`

#### `titles(values, options={}|[])`

  填充标题行，会自动应用默认选项`titleOpts`，`options`可以是数组分别指定每一列的选项，可以是统一的选项对象，
  通用选项参考`cell()`，特殊选项如下:

* `newLine` 是否在新的行填充单元格，默认`true`，如果为`false`表示在当前行后续单元格填充

#### `fillRow(values, options={}|[])`

  填充数据行，`options`可以是数组分别指定每一列的选项，可以是统一的选项对象，
  通用选项参考`cell()`，特殊选项如下:

* `newLine` 是否在新的行填充单元格，默认`true`，如果为`false`表示在当前行后续单元格填充

#### `fill(data, options={}|[])`

  填充数据表，`options`可以是数组分别指定每一列的选项，可以是统一的选项对象

#### `cell(value, options={})`

  填充单元个的值,`options`可以设置宽度和字体,`options`可选值:

  * `font`: 设置字体,`{bold: true, name: '微软雅黑', sz: 10}`,更多选项参考[Cell Styles](#cell-styles)
  * `width`: 设置单元格宽度
  * `bgColor`: 设置单元格背景
  * `fgColor`: 设置单元格字体颜色

#### `percent(value, precision, options)`

  填充百分比单元格，`value`为实际值
  * `value` 实际小数值，不需要乘以100，比如实际值是68%，传入0.68
  * `precision` 小数的精度，如果想百分化之后保留两位，则小数的精度需要4
  * `options` 同`cell`

```js
writer.percent(0.068, 4) // => 6.80%
writer.percent(0.068, 2) // => 7%
```

#### `currency(value, currency, precision, options)`

  填充货币单元格
  * `value` 实际小数值，不需要乘以100，比如实际值是68%，传入0.68
  * `precision` 小数的精度，如果想百分化之后保留两位，则小数的精度需要4
  * `options` 同`cell`

```js
writer.currency(1.234, '$', 2) // => $1.23
writer.currency(1.234, '¥', 3) // => ¥1.234
```

#### `utc(value, format, options)`

  填充日期单元格，参考例子[examples/date.js](examples/date.js)
  > 注意日期在Excel中显示的时候会转换成UTC的时间进行格式化展示
  * `value` 日期`Date`/`String`
  * `format` 日期格式
  * `options` 同`cell`

```js
const date = new Date('2018-01-01T15:16:00.000Z') // => 2018-01-01T15:16:00.000Z = 2018-01-01T23:16:00+08:00
writer.utc(date, 'YYYY-MM-DD HH:mm:ss') // => 无论运行在哪个时区都显示：2018-01-01 15:16:00
writer.utc('2018-01-01 23:16:00') // => 运行在UTC时区：2018-01-01 23:16:00，运行在北京时间：2018-01-01 15:16:00
```

#### `date(value, format, options)`

  填充日期单元格，展示当前时间在当前时区的格式，参考例子[examples/date.js](examples/date.js)
  * `value` 日期`Date`/`String`
  * `format` 日期格式
  * `options` 同`cell`

```js
const date = new Date('2018-01-01T15:16:00.000Z') // => 2018-01-01T15:16:00.000Z = 2018-01-01T23:16:00+08:00
writer.date(date, 'YYYY-MM-DD HH:mm:ss') // => 运行在UTC：2018-01-01 15:16:00，运行在北京时间：2018-01-01 23:16:00
writer.date('2018-01-01 23:16:00') // => 无论哪个时区都会显示：2018-01-01 23:16:00
```

#### `string(value, options)`

  填充文本单元格,可以用于编号是数字的，避免被格式化成科学计数
  * `value` 内容，可以是字符串、数字
  * `options` 同`cell`

#### `boolean(value, options)`

  填充布尔单元格
  * `value` 值
  * `options` 同`cell`

```js
writer.boolean(1) // => TRUE
writer.boolean(0) // => FALSE
writer.boolean('1') // => TRUE
writer.boolean('0') // => TRUE
writer.boolean(null) // => FALSE
```

#### `number(value, options)`

  填充数字单元格
  * `value` 内容，可以是字符串、数字
  * `options` 同`cell`

#### `color(bgColor, fgColor)`

  设置当前单元格颜色
  * `bgColor`为背景颜色
  * `fgColor`为字体颜色(RGB)

#### `fgColor(color)`

  设置当前单元格字体颜色(RGB)

#### `bgColor(color)`

  设置当前单元格背景颜色(RGB)

#### `mergeCell(cells=1)`

  合并当前单元和下N个单元格,默认下一个单元格。

#### `merge(rs, cs, re, ce)`

  合并指定单元格
  * `rs`:行起始位置
  * `cs`:列起始位置
  * `re`:行结束位置
  * `ce`:列结束位置

#### `border(rs, cs, re, ce, color, style = 'thin', options = {})`

  设置当前格的边框
  * `rs`:行起始位置
  * `cs`:列起始位置
  * `re`:行结束位置
  * `ce`:列结束位置
  * `color`:边框颜色
  * `style`:边框样式,更多样式参考[Cell Styles](#cell-styles)
  * `options`:选项(`outer`和`inner`可同时设置，只设置一个表示只设置内部或外部边框)
      - `outer`:是否设置外部边框，`true`或者`{color:'',style:''}`，`color`或`style`可选
      - `inner`:是否设置内部边框，`true`或者`{color:'',style:''}`，`color`或`style`可选

```js
writer.border(0, 0, 10, 10, '000000', 'thin', {
  // outer: true,
  outer: {
    style: 'thick' // 外部边框设置粗
  },
  inner: true
});
```

#### `border2end(r, c, color, style = 'thin', options = {})`

  给指定位置到表格结束设置边框
  * `r`:行起始位置
  * `c`:列起始位置
  * `color`:边框颜色
  * `style`:边框样式,更多样式参考[Cell Styles](#cell-styles)
  * `options`:同`border()`

#### `watermark(image)`

  给当前Excel所有sheet页添加水印，如果是`newSheet`返回的对象则只给对应sheet页添加水印
* `image` PNG水印图片文件或Buffer

#### `withoutWatermark()`
  
  取消之前设置的水印，build之前有效，如果是`newSheet`返回的对象只取消对应sheet的水印

#### `build(options = {})`

  构建并返回Excel内容,默认为`binary`,`options`选项

* `type` 内容类型:binary,base64,buffer,默认binary

#### `save(filename, options = {})`

  保存到指定文件

### `ExcelWriter.setDefaultWatermark(image)`

  设置默认水印，只对设置之后创建的Excel对象有效

### `ExcelReader`

#### `sheetNames()`

  返回所有Sheet页的数组

#### `sheet(name|index)`

  指定当前操作的sheet页，可以是sheetName，也可以是指定的序号

#### `header(titles, opts)`

  跳转到指定标题
* `titles` 必须的标题转换规则
* `opts` 可选的标题转换规则

  转换规则
```js
{
  "字段": "表格中的标题", // 不指定数据类型
  "时间": ["表格中的标题", "date"] // 指定数据类型
}
```
  支持指定的类型如下:
* `date` 当地时间，运行在每个时区得到的当地时间都相同
* `utc` UTC时间，运行在每个时区得到的实际时间都相同

#### `map(fn)`

  每一行数据的处理，默认返回当前行数据

### `watermark(buffer, image[, ignoreNonExcel])`

  给Excel文件增加水印，只支持Excel2017(`.xlsx`)之后的版本
* `buffer` Excel文件buffer
* `image` PNG水印图片文件或Buffer
* `ignoreNonExcel` 是否忽略非Excel文件，直接返回原文件内容

```js
const fs = require('fs')
const Excel = require('excelio');
Excel.watermark(fs.readFileSync('test.xlsx'), fs.readFileSync('test.png'));
```

#### `watermark.remove(buffer)`

  移除Excel文件中的水印，只支持Excel2017(`.xlsx`)之后的版本
* `buffer` Excel文件buffer
```js
const fs = require('fs')
Excel.watermark.remove(fs.readFileSync('test.xlsx'));
```

## Cell Styles

Cell styles are specified by a style object that roughly parallels the OpenXML structure.  The style object has five
top-level attributes: `fill`, `font`, `numFmt`, `alignment`, and `border`. 更多选项参考[xlsx-cell-styles][]


| Style Attribute | Sub Attributes | Values |
| :-------------- | :------------- | :------------- |
| fill            | patternType    |  `"solid"` or `"none"`  |
|                 | fgColor        |  `COLOR_SPEC`           |        
|                 | bgColor        |  `COLOR_SPEC`           |
| font            | name           |  `"Calibri"` // default |
|                 | sz             |  `"11"` // font size in points |
|                 | color          |  `COLOR_SPEC`        |
|                 | bold           |  `true` or `false` |
|                 | underline      |  `true` or `false` |
|                 | italic         |  `true` or `false` |
|                 | strike         |  `true` or `false` |
|                 | outline        |  `true` or `false` |
|                 | shadow         |  `true` or `false` |
|                 | vertAlign      |  `true` or `false` |
| numFmt          |                |  `"0"`  // integer index to built in formats, see StyleBuilder.SSF property |
|                 |                |  `"0.00%"` // string matching a built-in format, see StyleBuilder.SSF |
|                 |                |  `"0.0%"`  // string specifying a custom format |
|                 |                |  `"0.00%;\\(0.00%\\);\\-;@"` // string specifying a custom format, escaping special characters |
|                 |                |  `"m/dd/yy"` // string a date format using Excel's format notation |
| alignment       | vertical       | `"bottom"` or `"center"` or `"top"` |
|                 | horizontal     | `"bottom"` or `"center"` or `"top"` |
|                 | wrapText       |  `true \ false` |
|                 | readingOrder   |  `2` // for right-to-left |
|                 | textRotation   | Number from `0` to `180` or `255` (default is `0`) |
|                 |                |  `90` is rotated up 90 degrees |
|                 |                |  `45` is rotated up 45 degrees |
|                 |                | `135` is rotated down 45 degrees | 
|                 |                | `180` is rotated down 180 degrees |
|                 |                | `255` is special,  aligned vertically |
| border          | top            | `{ style: BORDER_STYLE, color: COLOR_SPEC }` |
|                 | bottom         | `{ style: BORDER_STYLE, color: COLOR_SPEC }` |
|                 | left           | `{ style: BORDER_STYLE, color: COLOR_SPEC }` |
|                 | right          | `{ style: BORDER_STYLE, color: COLOR_SPEC }` |
|                 | diagonal       | `{ style: BORDER_STYLE, color: COLOR_SPEC }` |
|                 | diagonalUp     | `true` or `false` |
|                 | diagonalDown   | `true` or `false` |


**COLOR_SPEC**: Colors for `fill`, `font`, and `border` are specified as objects, either:
* `{ auto: 1}` specifying automatic values
* `{ rgb: "FFFFAA00" }` specifying a hex ARGB value
* `{ theme: "1", tint: "-0.25"}` specifying an integer index to a theme color and a tint value (default 0)
* `{ indexed: 64}` default value for `fill.bgColor`

**BORDER_STYLE**: Border style is a string value which may take on one of the following values:
 * `thin`
 * `medium`
 * `thick`
 * `dotted`
 * `hair`
 * `dashed`
 * `mediumDashed`
 * `dashDot`
 * `mediumDashDot`
 * `dashDotDot`
 * `mediumDashDotDot`
 * `slantDashDot`


Borders for merged areas are specified for each cell within the merged area.  So to apply a box border to a merged area of 3x3 cells, border styles would need to be specified for eight different cells:
* left borders for the three cells on the left,
* right borders for the cells on the right
* top borders for the cells on the top
* bottom borders for the cells on the left

## Example

```js
const ExcelIO = require('excelio');
const writer = ExcelIO.createWriter();
writer.withoutGridLines().sheet('Test Sheet');
const tables = {
  'table 1': [['1-01', '1-02', '1-03', '1-04'], ['1-11', '1-12', '1-13', '1-14']],
  'table 2': [['2-01', '2-02', '2-03', '2-04'], ['2-11', '2-12', '2-13', '2-14']],
};
Object.keys(tables).forEach((table, i)=>{
  let cell = i * 3 + 1;
  writer.go(0); // 回到首行
  tables[table].forEach(data=>{
    let row = writer.rowIndex() + 2;
    writer.skipRow()
      .width(5, cell).width(5, cell + 1)
      .row(cell).cell(data[0]).cell(data[1])
      .row(cell).cell(data[2]).cell(data[3])
      .border(row, cell, writer.rowIndex(), cell + 1, '#000000')
  })
});

const binary = writer.build();
// write to file or response to http res.write(buffe, 'binary')
```

  the result as follows

    +----+----+     +----+----+
    |1-01|1-02|     |2-01|2-02|
    +----+----+     +----+----+
    |1-03|1-04|     |2-03|2-04|
    +----+----+     +----+----+

    +----+----+     +----+----+
    |1-11|1-12|     |2-11|2-12|
    +----+----+     +----+----+
    |1-13|1-14|     |2-13|2-14|
    +----+----+     +----+----+


### 读取Excel

```js
const reader = new ExcelReader()

reader.readFile('file.xlsx');
reader.sheet(0).header({
  '日期': ['date', 'date'],
  '描述': 'description',
}, {
  '非必须字段': 'optional',
});
const data = reader.map(v => v && v.date ? v : null);
console.log(data);
```

[xlsx-cell-styles]: https://www.npmjs.com/package/xlsx-style#cell-styles "xlsx-style"
