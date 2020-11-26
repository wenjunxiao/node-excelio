const XLSX = require('xlsx-style');
const _ = require('lodash');

const fmtValue = cell => cell && (cell.v == cell.w ? cell.v : (cell.w || cell.v));

const excel2utc = v => 86400 * v * 1000 - 2209161600000;

class ExcelReader {

  constructor(options) {
    this._opts = _.assign({}, options || {});
    this._formatted = this._opts.formatted === true;
    this._types = this._opts.types || {};
  }

  readFile (filename) {
    this.wb = XLSX.readFile(filename);
    return this;
  }

  read (data) {
    this.wb = XLSX.read(data, {
      type: Buffer.isBuffer(data) ? 'buffer' : 'binary'
    });
    return this;
  }

  sheetNames () {
    return this.wb.SheetNames;
  }

  sheet (name) {
    if (isNaN(name)) {
      this.curSheet = this.wb.Sheets[name];
    } else {
      this.curSheet = this.wb.Sheets[this.wb.SheetNames[name]];
    }
    const range = XLSX.utils.decode_range(this.curSheet['!ref']);
    this.rowIdx = range.s.r;
    this.colIdx = range.s.c;
    this.rows = range.e.r;
    this.cols = range.e.c;
    return this;
  }

  _cellValue (cell, type) {
    let v = this._formatted ? fmtValue(cell) : cell && cell.v;
    if (type) {
      if (type === 'utc') { // UTC 时间
        if (typeof v === 'number') {
          v = new Date(excel2utc(v));
        } else {
          v = new Date(v);
          v = new Date(Date.UTC(v.getFullYear(), v.getMonth(), v.getDate(), v.getHours(),
            v.getMinutes(), v.getSeconds(), v.getMilliseconds()));
        }
      } else if (type === 'date') {
        if (typeof v === 'number') {
          v = new Date(excel2utc(v));
          v = new Date(v.getUTCFullYear(), v.getUTCMonth(), v.getUTCDate(),
            v.getUTCHours(), v.getUTCMinutes(), v.getUTCSeconds(), v.getUTCMilliseconds());
        } else {
          v = new Date(v);
        }
      }
    }
    return v;
  }

  row () {
    if (this.rowIdx < this.rows) {
      const data = [];
      for (let i = 0; i < this.cols; i++) {
        const ref = XLSX.utils.encode_cell({
          c: i,
          r: this.rowIdx
        });
        const cell = this.curSheet[ref];
        data.push(this._cellValue(cell, this._types[i]));
      }
      this.rowIdx++;
      return data
    }
  }

  header (titles, opts, mapper) {
    let ts = Object.assign({}, titles);
    let os = Object.assign({}, opts);
    let types = Object.assign({}, this._opts.types);
    for (; this.rowIdx <= this.rows; this.rowIdx++) {
      const fs = {};
      const headers = this.headers = {};
      for (let i = 0; i <= this.cols; i++) {
        const ref = XLSX.utils.encode_cell({
          c: i,
          r: this.rowIdx
        });
        const cell = this.curSheet[ref];
        if (cell && cell.v) {
          let repeat = 1; // repeats[cell.v] = (repeats[cell.v] || 0) + 1;
          let n = mapper ? mapper(cell.v, repeat) || cell.v : cell.v;
          if (n !== cell.v) {
            while (headers[n]) {
              const nn = mapper(cell.v, ++repeat);
              if (nn === n) {
                n = null;
                break;
              } else {
                n = nn;
              }
            }
            if (!n) continue;
          }
          if (ts[n] !== undefined) {
            const tn = ts[n];
            if (Array.isArray(tn)) {
              headers[n] = fs[i] = tn[0];
              types[tn[0]] = tn[1];
            } else {
              headers[n] = fs[i] = tn;
            }
            delete ts[n];
          } else if (os[n] !== undefined) {
            headers[n] = fs[i] = os[n];
            delete os[n];
          }
        }
      }
      if (Object.keys(ts).length === 0) {
        this.fs = fs;
        this._types = types;
        this.rowIdx++;
        return this;
      }
      ts = Object.assign({}, titles);
      os = Object.assign({}, opts);
    }
    throw new Error('没有找到指定的头:' + Object.keys(ts).join(','));
  }

  has (title) {
    if (this.headers) {
      return !!this.headers[title];
    }
    return false;
  }

  map (fn) {
    const data = [];
    if (this.fs) {
      const ks = Object.keys(this.fs);
      for (; this.rowIdx <= this.rows; this.rowIdx++) {
        const row = {};
        for (let i of ks) {
          const ref = XLSX.utils.encode_cell({
            c: i,
            r: this.rowIdx
          });
          const cell = this.curSheet[ref];
          const name = this.fs[i];
          const type = this._types[name];
          row[name] = this._cellValue(cell, type);
        }
        let v = fn ? fn(row) : row;
        if (v) {
          data.push(v);
        }
      }
    } else {
      let row;
      while ((row = this.row())) {
        let v = fn ? fn(row) : row;
        if (v) {
          data.push(v);
        }
      }
    }
    return data;
  }
}

module.exports = ExcelReader;
