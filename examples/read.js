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
  .row().cell('UTC时间', titleOpts).width(16)
  .cell('当地时间', titleOpts).width(16);
const date = new Date()
writer
  .row().utc(date).date(date)
  .row().cell('2019-11-26 07:10:41').cell('2019-11-26 15:10:41')

console.log('date =>', date);
const reader = Excel.createReader({});
reader.read(writer.build({ type: 'buffer' }))
reader.active(0).header({
  'UTC时间': ['utc', 'utc'],
  '当地时间': ['date', 'date']
});
const result = reader.map(v => {
  if (v && v.date) {
    v.utc_time = moment.utc(v.utc).format('YYYY-MM-DD HH:mm:ss');
    v.utc2local = moment(v.utc).format('YYYY-MM-DD HH:mm:ss');
    v.date_time = moment(v.date).format('YYYY-MM-DD HH:mm:ss');
    v.utc_hours = v.utc.getUTCHours();
    v.date_hours = v.date.getHours();
    return v;
  }
  return null;
});

console.log(JSON.stringify(result, null, 2));
