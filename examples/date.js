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
const moment = require('moment')
writer.withoutGridLines().sheet('Date')
  .row().cell('utc', titleOpts).width(160)
  .cell('expect', titleOpts).width(160)
  .cell('date', titleOpts).width(160)
  .cell('expect', titleOpts).width(160);

// 数据库存储值：2019-07-31 07:18:00
// 北京日期对象：2019-07-30T23:18:00.000Z
// UTC日期对象：2019-07-31T07:18:00.000Z
// 与数据库相同：2019-07-31 07:18:00
// 北京时间预期：2019-07-31 15:18:00
// UTC时间预期：2019-07-31 07:18:00
const date = new Date('2019-07-31 07:18:00')
const utc = moment.utc(moment(date).format('YYYY-MM-DD HH:mm:ss')).utcOffset(8).toDate()
const bj = moment.utc(moment(date).format('YYYY-MM-DD HH:mm:ss')).subtract(8, 'hours').toDate()
const data = [
  // 不论时区与数据库保持一致
  {
    v1: date,
    v2: process.env.TZ === 'UTC' ? '2019-07-31 07:18:00' : '2019-07-30 23:18:00',
    v3: date,
    v4: '2019-07-31 07:18:00'
  }, { // 数据库存储的是UTC时间，显示成北京时间
    v1: utc,
    v2: '2019-07-31 07:18:00',
    v3: moment.utc(moment(date).format('YYYY-MM-DD HH:mm:ss')).utcOffset(8).format('YYYY-MM-DD HH:mm:ss'),
    v4: '2019-07-31 15:18:00'
  }, { // 数据库存储的是北京时间，显示成北京时间
    v1: bj,
    v2: '2019-07-30 23:18:00',
    v3: date,
    v4: '2019-07-31 07:18:00'
  }
]

for (let d of data) {
  writer.row().utc(d.v1, 'YYYY-MM-DD HH:mm:ss').cell(d.v2).date(d.v3, 'YYYY-MM-DD HH:mm:ss').cell(d.v4);
}
writer.row().row()
const d = new Date('2018-01-01T15:16:00.000Z')
writer.utc(d, 'YYYY-MM-DD HH:mm:ss').cell('2018-01-01 15:16:00')
writer.utc('2018-01-01 23:16:00').cell(process.env.TZ === 'UTC' ? '2018-01-01 23:16:00' : '2018-01-01 15:16:00')
writer.row()
writer.date(d, 'YYYY-MM-DD HH:mm:ss').cell(process.env.TZ === 'UTC' ? '2018-01-01 15:16:00' : '2018-01-01 23:16:00')
writer.date('2018-01-01 23:16:00').cell('2018-01-01 23:16:00')
writer.border2end(0, 0, '000000');
writer.save(path.resolve(__dirname, 'date.xlsx'));