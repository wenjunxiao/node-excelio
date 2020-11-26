'use strict';

const ExcelWriter = require('./writer');
const ExcelReader = require('./reader');
const watermark = require('./watermark');

module.exports = {
  watermark,
  ExcelWriter: ExcelWriter,
  ExcelReader: ExcelReader,
  createWriter: (options = {}) => {
    return new ExcelWriter(options);
  },
  createReader: (options = {}) => {
    return new ExcelReader(options);
  }
};
