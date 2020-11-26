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
  .row().cell('UTC时间', titleOpts).width(160)
  .cell('当地时间', titleOpts).width(160);
const date = new Date()
writer
  .row().utc(date).date(date)
  .row().cell('2019-11-26 07:10:41').cell('2019-11-26 15:10:41')


const reader = Excel.createReader({});
reader.read(writer.build({ type: 'buffer' }))
reader.sheet(0).header({
  'UTC时间': ['utc', 'utc'],
  '当地时间': ['date', 'date']
});
const result = reader.map(v => {
  if (v && v.date) {
    console.log(v)
    v.utc_ = moment.utc(v.utc).format('YYYY-MM-DD HH:mm:ss');
    v.utc_local = moment(v.utc).format('YYYY-MM-DD HH:mm:ss');
    v.date_ = moment(v.date).format('YYYY-MM-DD HH:mm:ss');
    v.utc_hours = v.utc.getHours();
    v.date_hours = v.date.getHours();
    return v;
  }
  return null;
});

console.log(result)
