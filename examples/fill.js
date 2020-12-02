const path = require('path');
const Excel = require('../index');
let writer = Excel.createWriter({
  NaN: '-',
  titleOpts: {
    font: {
      bold: true,
      name: '微软雅黑',
      sz: 10
    }
  },
  showGridLines: false,
  border2end: true
});

writer.sheet('Sheet1')
  .row().title('Title1', 8).title('Title2', 10).title('Title3', 12)
  .fillRow([11, 12, 13])
  .fill([[21, 22, 23], [31, 32, 33]])
  .newSheet('Sheet2')
  .titles(['Title1', 'Title2', 'Title3'], [8, 10, 12])
  .fill([[11, 12, 13], [21, 22, 23], [31, 32, 33]])

const filename = path.resolve(__dirname, 'fill.xlsx');
writer.save(filename);
console.log('saved =>', filename);