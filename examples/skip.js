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
writer.withoutGridLines().sheet('Single')
  .row().cell('Title1', titleOpts).width(10)
  .cell('Title2', titleOpts).width(12)
  .cell('Title3', titleOpts).width(12);


writer.row().cell('Cell11').cell('Cell12').cell('Cell13')
.skipRow(2, -1).cell('Cell33')
.skipRow(1, -1).cell('Cell43')
.skipRow(1, -2).cell('Cell42')
writer.border2end(0, 0, 'red');
const filename = path.resolve(__dirname, 'skip.xlsx');
writer.save(filename);
console.log('saved =>', filename);
