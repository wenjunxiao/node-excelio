'use strict';

const ExcelWriter = require('./writer');
const ExcelReader = require('./reader');
const watermark = require('./watermark');

module.exports = {
  watermark,
  ExcelWriter: ExcelWriter,
  ExcelReader: ExcelReader,
  /**
   * @returns {ExcelWriter}
   */
  createWriter: (options = {}) => {
    return new ExcelWriter(options);
  },
  /**
   * @returns {ExcelReader}
   */
  createReader: (options = {}) => {
    return new ExcelReader(options);
  }
};
