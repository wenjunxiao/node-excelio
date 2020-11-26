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
  .row().title('Title1', 80).title('Title2', 100).title('Title3', 120)
  .fillRow([11, 12, 13])
  .fill([[21, 22, 23], [31, 32, 33]])
  .newSheet('Sheet2')
  .titles(['Title1', 'Title2', 'Title3'], [80, 100, 120])
  .fill([[11, 12, 13], [21, 22, 23], [31, 32, 33]])

writer.save(path.resolve(__dirname, 'fill.xlsx'));