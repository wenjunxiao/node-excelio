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
writer.withoutGridLines().sheet('Single').row()
  .row(1).cell('Title1', titleOpts).width(10)
  .cell('Title2', titleOpts).width(12)
  .cell('Title3', titleOpts).width(12);

const data = [{
    v1: 11,
    v2: 12,
    v3: 13
  },
  {
    v1: 21,
    v2: 22,
    v3: 23
  },
  {
    v1: 31,
    v2: 32,
    v3: 33
  }
]

for (let d of data) {
  writer.row(1).cell(d.v1).currency(d.v2, '$').number(d.v3)
}
writer.border2end(1, 1, '#9bd6c4', 'thick', {
  // outer: true,
  // outer: {
  //   style: 'thick',
  //   color: '#9bd6c4'
  // },
  inner: false
});
const filename = path.resolve(__dirname, 'border.xlsx');
writer.save(filename);
console.log('saved =>', filename);