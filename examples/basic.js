const path = require('path');
const Excel = require('../index');
let writer = Excel.createWriter({
  NaN: '-'
});
const titleOpts = {
  font: {
    bold: true,
    name: '微软雅黑',
    sz: 10
  }
};
writer.withoutGridLines().sheet('Basic')
  .row()
  .cell('A', titleOpts).width(1)
  .cell('AB', titleOpts).width(2)
  .cell('ABC', titleOpts).width(3)
  .cell('ABCD', titleOpts).width(4)
  .cell('ABCDE', titleOpts).width(5)
  .cell('ABCDEF', titleOpts).width(6)
  .cell('ABCDEFG', titleOpts).width(7)
  .cell('ABCDEFGH', titleOpts).width(8)
  .cell('ABCDEFGHI', titleOpts).width(9)
  .cell('ABCDEFGHIJ', titleOpts).width(10)
  .cell('ABCDEFGHIJKLMNO', titleOpts).width(15)
  .cell('ABCDEFGHIJKLMNOPQRST', titleOpts).width(20)
  .cell('ABCDEFGHIJKLMNOPQRSTUVWXYZ', titleOpts).width(26)
  .cell('中文', titleOpts).chWidth(2)
  .cell('中文字', titleOpts).chWidth(3)
;

writer.border2end(0, 0);
writer.save(path.resolve(__dirname, 'basic.xlsx'));