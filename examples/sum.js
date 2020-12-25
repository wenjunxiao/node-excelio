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
  .row().title('订单ID', 8).title('数量', 10).title('金额', 12)
  .row().cell(1).number(11).currency(12.31, '$')
  .row().cell(2).number(12).currency(12.32, '$')
  .row().cell(3).number(13).currency(12.33, '$')
  .row().cell('汇总').sum().sum()
  .newSheet('Sheet2')
  .titles(['Title1', 'Title2', 'Title3'], [8, 10, 12])
  .fill([[11, 12, 13], [21, 22, 23], [31, 32, 33]])
  .row().sum().sum().sum()

const filename = path.resolve(__dirname, 'sum.xlsx');
writer.save(filename);
console.log('saved =>', filename);