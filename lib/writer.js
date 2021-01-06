'use strict';

const fs = require('fs');
const _ = require('lodash');
const XLSX = require('xlsx-style');
const JSZip = require('xlsx-style/jszip');
const iconv = require('iconv-lite');
const crypto = require('crypto');

const char2px = (chr, sz) => {
  chr = chr || 0;
  return (chr * 8 + Math.ceil(chr / 10) * 5) * Math.ceil(sz / 10);
}
const px = px => px || 0;

function md5 (s) {
  return crypto
    .createHash('md5')
    .update(s, 'utf8')
    .digest('hex');
}

function buildBorder (color, style) {
  color = color.replace(/^#+/, '')
  return {
    top: {
      style: style,
      color: {
        rgb: color
      }
    },
    bottom: {
      style: style,
      color: {
        rgb: color
      }
    },
    left: {
      style: style,
      color: {
        rgb: color
      }
    },
    right: {
      style: style,
      color: {
        rgb: color
      }
    }
  };
}

function checkPrecision (precision, options) {
  if (typeof precision === 'object') {
    return [undefined, precision];
  }
  return [precision, options];
}

const defaults = {
  watermark: null
}

class Sheet {
  constructor(options, watermark) {
    this.rowIdx = -1;
    this.colIdx = -1;
    this.maxCol = 0;
    this._opts = _.assign({}, options || {});
    this.sheetName = this._opts.sheetName || 'Sheet1';
    this.owner = this._opts.owner || {};
    this.NaN = this._opts.NaN || '';
    this._undefined = this._opts.undefined || '';
    this._px = this._opts.px === true;
    this._width = this._px ? px : char2px;
    this.s = {
      "!row": [{
        wpx: 67
      }]
    };
    this['!cols'] = [];
    this['!merges'] = [];
    this._watermark = watermark;
  }

  loadFromSheet(sheet) {
    const range = XLSX.utils.decode_range(sheet['!ref']);
    Object.assign(this, sheet);
    this.rowIdx = range.e.r;
    this.colIdx = this.maxCol = range.e.c;
    return this;
  }

  clear () {
    this.rowIdx = -1;
    this.colIdx = -1;
    this.maxCol = 0;
    this['!cols'] = [];
    this['!merges'] = [];
    return this;
  }

  rename (name) {
    if (name !== this.sheetName) {
      let cur = this.sheetName;
      this.sheetName = name;
      this.owner.rename(name, cur);
    }
    return this;
  }

  rowIndex () {
    return this.rowIdx;
  }

  colIndex () {
    return this.colIdx;
  }

  skipRow (rows = 1, cells = 0) {
    if (this.colIdx > this.maxCol) {
      this.maxCol = this.colIdx;
    }
    this.rowIdx += rows;
    if (cells < 0) {
      this.colIdx += cells;
    } else {
      this.colIdx = cells - 1;
    }
    return this;
  }

  skipCell (cells = 1) {
    this.colIdx += cells;
    return this;
  }

  go (row, cell = 0) {
    if (this.colIdx > this.maxCol) {
      this.maxCol = this.colIdx;
    }
    this.rowIdx = row - 1;
    this.colIdx = cell - 1;
    return this;
  }

  row (cells = 0) {
    if (this.colIdx > this.maxCol) {
      this.maxCol = this.colIdx;
    }
    ++this.rowIdx;
    this.colIdx = cells - 1;
    return this;
  }

  number (v, options) {
    v = this.formatNumber(v);
    if (isNaN(v)) return this.cell(v, options);
    return this.cell(v, options, 'n');
  }

  boolean (v, options) {
    return this.cell(v, options, 'b');
  }

  string (v, options) {
    return this.cell(v, options, 's');
  }

  utc (v, format, options) {
    if (typeof format === 'object') {
      options = format
      format = options.format
    }
    if (!format) {
      format = 'YYYY-MM-DD HH:mm:ss'
    }
    if (v instanceof Date) {
      v = v.toISOString();
    }
    return this.cell(v, options, 'd', format);
  }

  date (v, format, options) {
    if (typeof format === 'object') {
      options = format
      format = options.format
    }
    if (!(v instanceof Date)) {
      v = new Date(v)
    }
    if (!format) {
      format = 'YYYY-MM-DD HH:mm:ss'
    }
    v = new Date(Date.UTC(v.getFullYear(), v.getMonth(), v.getDate(), v.getHours(),
      v.getMinutes(), v.getSeconds(), v.getMilliseconds()));
    return this.cell(v.toISOString(), options, 'd', format);
  }

  formatNumber (v, precision) {
    if (v === null || v === undefined) return this.NaN;
    if (isNaN(precision) || !v.toFixed) return v.toString().replace(/,/g, '');
    return v.toFixed(precision).replace(/,/g, '');
  }

  currency (v, currency, precision, options) {
    [precision, options] = checkPrecision(precision, options);
    v = this.formatNumber(v, precision);
    if (currency) {
      return this.cell(v, options, 'n', currency + '#,##0.00');
    } else if (/^([^\d\-\.])+(.*)$/.test(v)) {
      let prefix = RegExp.$1;
      v = RegExp.$2;
      return this.cell(v, options, 'n', prefix + '#,##0.00');
    } else if (isNaN(v)) {
      return this.cell(v, options);
    }
    return this.cell(v, options, 'n', '4');
  }

  percent (v, precision, options) {
    [precision, options] = checkPrecision(precision, options);
    v = this.formatNumber(v, precision);
    if (isNaN(v)) {
      return this.cell(v, options);
    }
    return this.cell(v, options, 'n', '0.00%');
  }

  sum (options, start = 1) {
    const colIndex = this.colIdx + 1;
    let v = 0;
    let format;
    for (let i = start; i < this.rowIdx; i++) {
      let cellRef = XLSX.utils.encode_cell({
        c: colIndex,
        r: i
      });
      let cell = this[cellRef];
      if (cell) {
        if (cell.z) {
          format = cell.z;
        }
        v = v + Number(cell.v);
      }
    }
    return this.cell(v, options, 'n', format);
  }

  titles (vs, options) {
    options = typeof options === 'object' ? options : { width: options };
    if (this.rowIdx === -1 || options.newLine !== false) {
      this.row();
    }
    if (Array.isArray(options)) {
      vs.forEach((v, i) => {
        this.title(v, options[i])
      });
    } else {
      if (this._opts.titleOpts) {
        options = _.merge({}, this._opts.titleOpts, options);
      }
      for (let v of vs) {
        this.cell(v, options, 's');
      }
    }
    return this;
  }

  title (v, options) {
    options = typeof options === 'object' ? options : { width: options };
    if (this._opts.titleOpts) {
      options = _.merge({}, this._opts.titleOpts, options);
    }
    return this.cell(v, options, 's');
  }

  fillRow (vs, options) {
    options = typeof options === 'object' ? options : { width: options };
    if (this.rowIdx === -1 || options.newLine !== false) {
      this.row();
    }
    if (Array.isArray(options)) {
      vs.forEach((v, i) => {
        this.cell(v, options[i])
      });
    } else {
      for (let v of vs) {
        this.cell(v, options);
      }
    }
    return this;
  }

  fill (data, options) {
    if (Array.isArray(options)) {
      for (let vs of data) {
        this.row();
        vs.forEach((v, i) => {
          this.cell(v, options[i])
        });
      }
    } else {
      for (let vs of data) {
        this.row();
        for (let v of vs) {
          this.cell(v, options);
        }
      }
    }
    return this;
  }

  cell (v, options, type, format) {
    if (!v && typeof v === 'undefined') {
      v = this._undefined;
    }
    const colIdx = ++this.colIdx;
    if (this._opts.titleOpts && this._opts.titleLine >= 0) { // 标题行
      options = typeof options === 'object' ? options : { width: options };
      if (this.rowIdx === this._opts.titleLine) {
        options = _.merge({}, this._opts.titleOpts, options);
      } else if (this._opts.cellOpts) {
        options = _.merge({}, this._opts.cellOpts, options);
      }
    } else if (this._opts.cellOpts) {
      options = typeof options === 'object' ? options : { width: options };
      options = _.merge({}, this._opts.cellOpts, options);
    }
    let cell = {
      v: v,
      s: {
        alignment: options && options.alignment || this._opts.alignment || {
          horizontal: 'left',
          vertical: 'left'
        }
      }
    };
    if (format) {
      cell.z = format;
    }
    if (type) {
      cell.t = type;
    } else if (typeof v === 'number') {
      cell.t = 'n';
    } else if (typeof v === 'boolean') {
      cell.t = 'b';
    } else if (v instanceof Date) {
      cell.t = 'd';
    } else {
      cell.t = 's';
    }
    if (options || options === 0) {
      if (options.font) {
        cell.s.font = _.assign(cell.s.font || {}, options.font);
      }
      let fsz = cell.s.font && cell.s.font.sz || this._opts.fontSize || 10;
      let width = this._width(typeof options === 'object' ? options.width : options, fsz);
      if (width === 0) {
        width = char2px(iconv.encode(v.toString(), 'gbk').length);
      }
      if (!width) {
        width = this._width(this._opts.width, fsz);
      }
      if (width) {
        if (this._opts.minWidth && width < this._opts.minWidth) {
          width = this._width(this._opts.minWidth, fsz);
        }
        let col = this['!cols'][colIdx];
        if (!col) {
          col = this['!cols'][colIdx] = {};
        }
        if (!(col.wpx && col.wpx > width)) {
          col.wpx = width;
        }
      } else if (v) {
        if (!this['!cols'][colIdx]) {
          this['!cols'][colIdx] = {
            wpx: iconv.encode(v.toString(), 'gbk').length * 8
          };
        }
      }
      if (options.bgColor || options.fgColor) {
        this.color(options.bgColor, options.fgColor);
      }
      if (options.type) {
        cell.t = options.type;
      }
    }
    let cellRef = XLSX.utils.encode_cell({
      c: colIdx,
      r: this.rowIdx
    });
    this.curCell = this[cellRef] = cell;
    return this;
  }

  wrap () {
    let cellRef = XLSX.utils.encode_cell({
      c: this.colIdx,
      r: this.rowIdx
    });
    let cell = this[cellRef];
    if (!cell.s) {
      cell.s = {};
    }
    if (!cell.s.alignment) {
      cell.s.alignment = {};
    }
    cell.s.alignment.wrapText = true;
    return this;
  }

  chWidth (width, colIndex = -1) {
    if (this._px) {
      return this.width(width, colIndex);
    }
    return this.width(width * 1.8, colIndex);
  }

  width (width, colIndex = -1) {
    if (colIndex < 0) {
      colIndex = this.colIdx;
      let col = this['!cols'][colIndex];
      if (!col) {
        col = this['!cols'][colIndex] = {};
      }
      let fsz = this.curCell.s.font && this.curCell.s.font.sz || this._opts.fontSize || 10;
      col.wpx = this._width(width, fsz);
    } else {
      let col = this['!cols'][colIndex];
      if (!col) {
        col = this['!cols'][colIndex] = {};
      }
      let cellRef = XLSX.utils.encode_cell({
        c: colIndex,
        r: this.rowIdx
      });
      let fsz = cellRef && cellRef.s && cellRef.s.font && cellRef.s.font.sz || 10;
      col.wpx = this._width(width, fsz);
    }
    return this;
  }

  color (bgColor, fgColor) {
    if (bgColor) {
      let fill = this.curCell.s.fill;
      if (!fill) {
        fill = this.curCell.s.fill = {};
      }
      fill.fgColor = {
        rgb: bgColor.replace(/^#+/, '')
      };
    }
    if (fgColor) {
      let font = this.curCell.s.font;
      if (!font) {
        font = this.curCell.s.font = {};
      }
      font.color = {
        rgb: fgColor.replace(/^#+/, '')
      };
    }
    return this;
  }

  fgColor (color) {
    return this.color(null, color);
  }

  bgColor (color) {
    return this.color(color, null);
  }

  watermark (image) {
    if (image && !Buffer.isBuffer(image)) {
      throw new Error('水印必须是PNG图片Buffer');
    }
    this._watermark = image;
    return this;
  }

  withoutWatermark () {
    this._watermark = null;
    return this;
  }

  getBorderCell (row, col) {
    let cellRef = XLSX.utils.encode_cell({
      c: col,
      r: row
    });
    let cell = this[cellRef];
    if (!cell) {
      cell = this[cellRef] = {
        v: '',
        s: {}
      };
    }
    if (!cell.s.border) {
      cell.s.border = {};
    }
    return cell;
  }

  border (rs, cs, re, ce, color = '#000000', style = 'thin', options = {}) {
    options = options || {};
    if (options.hasOwnProperty('outer') || options.hasOwnProperty('inner')) {
      if (options.inner !== false) {
        const bd = buildBorder(options.inner.color || color, options.inner.style || style);
        for (let ri = rs; ri <= re; ri++) {
          for (let ci = cs; ci <= ce; ci++) {
            let cell = this.getBorderCell(ri, ci);
            cell.s.border = Object.assign({}, bd);
            if (ri === rs) {
              delete cell.s.border.top;
            }
            if (ri === re) {
              delete cell.s.border.bottom;
            }
            if (ci === cs) {
              delete cell.s.border.left;
            }
            if (ci === ce) {
              delete cell.s.border.right;
            }
          }
        }
      }
      if (options.outer !== false) {
        options.outer = options.outer || { color, style };
      }
      if (options.outer) {
        const bd = buildBorder(options.outer.color || color, options.outer.style || style);
        let cell = this.getBorderCell(rs, cs); // 第一个单元格 左上
        cell.s.border.top = bd.top;
        cell.s.border.left = bd.left;
        cell = this.getBorderCell(rs, ce); // 第一行最后一个单元格 右上
        cell.s.border.top = bd.top;
        cell.s.border.right = bd.right;
        cell = this.getBorderCell(re, cs); // 最后一行第一个单元格 左下
        cell.s.border.bottom = bd.bottom;
        cell.s.border.left = bd.left;
        cell = this.getBorderCell(re, ce); // 最后一行最后一个单元格 右下
        cell.s.border.bottom = bd.bottom;
        cell.s.border.right = bd.right;
        for (let ri = rs + 1; ri < re; ri++) {
          cell = this.getBorderCell(ri, cs);
          cell.s.border.left = bd.left;
          cell = this.getBorderCell(ri, ce);
          cell.s.border.right = bd.right;
        }
        for (let ci = cs + 1; ci < ce; ci++) {
          cell = this.getBorderCell(rs, ci);
          cell.s.border.top = bd.top;
          cell = this.getBorderCell(re, ci);
          cell.s.border.bottom = bd.bottom;
        }
      }
    } else {
      const bd = buildBorder(color, style);
      for (let ri = rs; ri <= re; ri++) {
        for (let ci = cs; ci <= ce; ci++) {
          let cell = this.getBorderCell(ri, ci);
          cell.s.border = bd;
        }
      }
    }
    return this;
  }

  border2end (r, c, color = '#000000', style = 'thin', options = {}) {
    let maxCol = this.colIdx > this.maxCol ? this.colIdx : this.maxCol;
    return this.border(r, c, this.rowIdx, maxCol, color, style, options);
  }

  mergeCell (cells = 1) {
    this.merge(this.rowIdx, this.colIdx, this.rowIdx, this.colIdx + cells);
    this.colIdx += cells;
    return this;
  }

  mergeRow (rows = 1) {
    this.merge(this.rowIdx, this.colIdx, this.rowIdx + rows, this.colIdx);
    return this;
  }

  merge (rs, cs, re, ce) {
    this['!merges'].push({
      s: {
        r: rs,
        c: cs
      },
      e: {
        r: re,
        c: ce
      }
    });
    return this;
  }

  end () {
    if (this.colIdx > this.maxCol) {
      this.maxCol = this.colIdx;
    }
    if (this._opts.border2end) {
      let opts = typeof this._opts.border2end === 'object' ? this._opts.border2end : {
        color: this._opts.border2end
      }
      if (opts.color === true) {
        opts.color = '#000000';
      }
      this.border2end(0, 0, opts.color, opts.style || 'thin');
    }
    this['!ref'] = XLSX.utils.encode_range({
      s: {
        c: 0,
        r: 0
      },
      e: {
        c: this.maxCol,
        r: this.rowIdx < 0 ? 0 : this.rowIdx
      }
    });
    return this;
  }
}

class ExcelWriter {

  constructor(options) {
    this.Sheets = {};
    this.SheetNames = [];
    this.rowIdx = -1;
    this.colIdx = -1;
    this.maxCol = 0;
    /**
     * @type Sheet
     */
    this.curSheet = null;
    this._opts = _.assign({
      bookType: 'xlsx',
      bookSST: false,
      type: 'binary'
    }, options || {});
    this.NaN = this._opts.NaN || '';
    this._watermark = defaults.watermark;
  }

  loadFromReader (reader) {
    return this.loadFromWorkbook(reader.wb);
  }

  loadFromFile (filename, options) {
    return this.loadFromWorkbook(XLSX.readFile(filename, options));
  }

  loadFromWorkbook (wb) {
    this.SheetNames = [].slice.call(wb.SheetNames);
    for (let name of this.SheetNames) {
      this.Sheets[name] = new Sheet(Object.assign({}, this._opts, {
        owner: this,
        sheetName: name
      }), this._watermark).loadFromSheet(wb.Sheets[name]);
    }
    return this;
  }

  withoutGridLines () {
    this._opts.showGridLines = false;
    return this;
  }

  rowIndex () {
    return this.curSheet.rowIndex();
  }

  colIndex () {
    return this.curSheet.colIndex();
  }

  /**
   * 新增Sheet
   * @param {*} name Sheet名称
   * @returns {Sheet}
   */
  newSheet (name) {
    const sheet = this.Sheets[name];
    if (sheet) return sheet;
    this.SheetNames.push(name);
    this.Sheets[name] = new Sheet(Object.assign({}, this._opts, {
      owner: this,
      sheetName: name
    }), this._watermark);
    return this.Sheets[name];
  }

  /**
   * 切换到指定的Sheet，如果不存在则新增
   * @param {*} name Sheet名
   * @returns {this}
   */
  sheet (name) {
    this.endSheet();
    this.curSheet = this.Sheets[name];
    if (!this.curSheet) {
      this.SheetNames.push(name);
      this.curSheet = this.Sheets[name] = new Sheet(Object.assign({}, this._opts, {
        owner: this,
        sheetName: name
      }), this._watermark);
    }
    return this;
  }

  /**
   * @returns {Sheet} 当前操作Sheet
   */
  active () {
    if (!this.curSheet) {
      this.curSheet = this.Sheets['Sheet1'] = new Sheet(Object.assign({}, this._opts, {
        owner: this,
        sheetName: 'Sheet1'
      }), this._watermark);
      this.SheetNames.push('Sheet1');
    }
    return this.curSheet;
  }

  rename (name, from) {
    if (this.SheetNames.indexOf(name) > -1) {
      throw new Error(`Sheet with name [${name}] already exists`);
    }
    if (from) {
      if (from !== name) {
        let pos = this.SheetNames.indexOf(from);
        this.SheetNames.splice(pos, 1, name);
        this.Sheets[name] = this.Sheets[from];
        delete this.Sheets[from];
        this.Sheets[name].rename(name);
      }
    } else if (name !== this.curSheet.sheetName) {
      from = this.curSheet.sheetName;
      let pos = this.SheetNames.indexOf(from);
      this.SheetNames.splice(pos, 1, name);
      this.Sheets[name] = this.Sheets[from];
      delete this.Sheets[from];
      this.curSheet.rename(name);
    }
    return this;
  }

  skipRow (rows = 1, cells = 0) {
    this.curSheet.skipRow(rows, cells);
    return this;
  }

  skipCell (cells = 1) {
    this.curSheet.skipCell(cells);
    return this;
  }

  go (row, cell = 0) {
    this.curSheet.go(row, cell);
    return this;
  }

  row (cells = 0) {
    this.curSheet.row(cells);
    return this;
  }

  number (v, options) {
    this.curSheet.number(v, options);
    return this;
  }

  boolean (v, options) {
    this.curSheet.boolean(v, options);
    return this;
  }

  string (v, options) {
    this.curSheet.string(v, options);
    return this;
  }

  utc (v, format, options) {
    this.curSheet.utc(v, format, options);
    return this;
  }

  date (v, format, options) {
    this.curSheet.date(v, format, options);
    return this;
  }

  formatNumber (v, precision) {
    return this.curSheet.formatNumber(v, precision);
  }

  currency (v, currency, precision, options) {
    this.curSheet.currency(v, currency, precision, options);
    return this;
  }

  percent (v, precision, options) {
    this.curSheet.percent(v, precision, options);
    return this;
  }

  sum (options, start = 1) {
    this.curSheet.sum(options, start);
    return this;
  }

  title (v, options) {
    this.curSheet.title(v, options);
    return this;
  }

  titles (vs, options) {
    this.curSheet.titles(vs, options);
    return this;
  }

  fillRow (vs, options) {
    this.curSheet.fillRow(vs, options);
    return this;
  }

  fill (data, options) {
    this.curSheet.fill(data, options);
    return this;
  }

  cell (v, options, type, format) {
    this.curSheet.cell(v, options, type, format);
    return this;
  }

  wrap () {
    this.curSheet.wrap();
    return this;
  }

  width (width, colIndex = -1) {
    this.curSheet.width(width, colIndex);
    return this;
  }

  chWidth (width, colIndex = -1) {
    this.curSheet.chWidth(width, colIndex);
    return this;
  }

  color (bgColor, fgColor) {
    this.curSheet.color(bgColor, fgColor);
    return this;
  }

  border (rs, cs, re, ce, color = '#000000', style = 'thin', options = {}) {
    this.curSheet.border(rs, cs, re, ce, color, style, options);
    return this;
  }

  border2end (r, c, color = '#000000', style = 'thin', options = {}) {
    this.curSheet.border2end(r, c, color, style, options);
    return this;
  }

  mergeCell (cells = 1) {
    this.curSheet.mergeCell(cells);
    return this;
  }

  mergeRow (rows = 1) {
    this.curSheet.mergeRow(rows);
    return this;
  }

  merge (rs, cs, re, ce) {
    this.curSheet.merge(rs, cs, re, ce);
    return this;
  }

  watermark (image) {
    if (image && !isValidWater(image)) {
      throw new Error('水印必须是PNG图片文件或Buffer');
    }
    this._watermark = image;
    return this;
  }

  withoutWatermark () {
    this._watermark = null;
    for (let name of this.SheetNames) {
      this.Sheets[name].withoutWatermark();
    }
    return this;
  }

  endSheet () {
    for (let name of this.SheetNames) {
      this.Sheets[name].end();
    }
    this.curSheet = null;
    return this;
  }

  build2target (options) {
    this.endSheet();
    const opts = _.assign({}, this._opts, options);
    let zip = null;
    for (let i = 0; i < this.SheetNames.length; i++) {
      const sheet = this.Sheets[this.SheetNames[i]];
      let watermark = sheet._watermark || this._watermark
      if (watermark && typeof watermark === 'string') {
        watermark = fs.readFileSync(watermark);
      }
      if (watermark) {
        if (!zip) {
          zip = JSZip()
          zip.load(XLSX.write(this, Object.assign({}, opts, { type: 'binary' })));
        }
        const name = md5(watermark.toString('binary')) + '.png';
        zip.folder("xl/media").file(name, watermark);
        const filename = `xl/worksheets/sheet${i + 1}.xml`;
        zip.file(filename, Buffer.from(zip.file(filename).asBinary(true)
          .replace(/(<\/worksheet>)/img, '<picture r:id="watermark"/>\n$1'), 'binary'));
        zip.folder("xl/worksheets/_rels").file(`sheet${i + 1}.xml.rels`,
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          `<Relationship Id="watermark" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${name}"/>` +
          '</Relationships>'
        );
      }
    }
    if (zip) {
      zip.file('[Content_Types].xml', Buffer.from(zip.file('[Content_Types].xml').asBinary(true)
        .replace(/(<Types\s+[^>]*>)/img, '$1<Default Extension="png" ContentType="image/png"/>')));
      const type = options && options.type || 'binary';
      switch (type) {
        case "base64": return zip.generate({ type: "base64" });
        case "binary": return zip.generate({ type: "string" });
        case "buffer": return zip.generate({ type: "nodebuffer" });
        default: throw new Error("Unrecognized type " + o.type);
      }
    }
    return XLSX.write(this, opts);
  }

  /**
   * 构建并返回数据
   * @param {{}} options
   * @param {String} [options.type] 数据类型:binary,base64,buffer,默认binary
   * @returns {*}
   */
  build (options = {}) {
    return this.build2target(options);
  }

  save (filename, options = {}) {
    return fs.writeFileSync(filename, this.build2target(Object.assign({}, options, { type: 'buffer' })));
  }
}

function isValidWater (image) {
  if (Buffer.isBuffer(image)) {
    return true;
  }
  return /\.png$/.test(image);
}

function setDefaultWatermark (image) {
  if (image && !isValidWater(image)) {
    throw new Error('水印必须是PNG图片文件或Buffer');
  }
  defaults.watermark = image;
}

function getDefaultWatermark () {
  return defaults.watermark;
}

module.exports = ExcelWriter;
module.exports.setDefaultWatermark = setDefaultWatermark;
module.exports.getDefaultWatermark = getDefaultWatermark;
