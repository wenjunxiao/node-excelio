const path = require('path');
const { read } = require('xlsx-style');
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

const reader = Excel.createReader({});
reader.read(writer.build({ type: 'buffer' }))
reader.active(0).header({
  '订单ID': 'id',
  '数量': 'quality',
  '金额': 'amount'
}).tailor(['汇总']);

console.log('map =>', reader.map(v => v));

reader.reset();
while(reader.hasNext()) {
  console.log('next =>', reader.next());
}
