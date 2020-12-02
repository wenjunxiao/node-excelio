const path = require('path');
const ExcelIO = require('../index');
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
const filename = path.resolve(__dirname, 'example.xlsx');
writer.save(filename);
console.log('saved =>', filename);