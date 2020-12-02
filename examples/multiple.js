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
['Test1', 'Test2', 'Test3'].map((name, i) => {
  const sheet = writer.newSheet(name);
  sheet.row().cell(name + 'Title1', titleOpts).width(10)
    .cell(name + 'Title2', titleOpts).width(12)
    .cell(name + 'Title3', titleOpts).width(12);
  const data = [{
      v1: 11 + i * 1000,
      v2: 12 + i * 1000,
      v3: 13 + i * 1000
    },
    {
      v1: 21 + i * 1000,
      v2: 22 + i * 1000,
      v3: 23 + i * 1000
    },
    {
      v1: 31 + i * 1000,
      v2: 32 + i * 1000,
      v3: 33 + i * 1000
    }
  ]
  for (let d of data) {
    sheet.row().cell(d.v1).currency(d.v2, '$').number(d.v3)
  }
  sheet.border2end(0, 0, '#000000');
})
writer.withoutGridLines()
const filename = path.resolve(__dirname, 'multiple.xlsx');
writer.save(filename);
console.log('saved =>', filename);