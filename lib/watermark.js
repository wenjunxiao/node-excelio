const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const JSZip = require('xlsx-style/jszip');

function md5 (s) {
  return crypto
    .createHash('md5')
    .update(s, 'utf8')
    .digest('hex');
}

function relationship(id, name) {
  return `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${name}"/>`
}

function watermark (buffer, image, ignoreNonExcel) {
  const hexHeader = Buffer.isBuffer(buffer) ? buffer.slice(0, 4).toString('hex') : Buffer.from(buffer.slice(0, 4), 'binary').toString('hex');
  if (ignoreNonExcel && hexHeader !== '504b0304') {
    return buffer;
  }
  let zip = JSZip();
  zip.load(buffer);
  if (ignoreNonExcel && !zip.file('xl/workbook.xml')) {
    return buffer;
  }
  if (typeof image === 'string') {
    image = fs.readFileSync('image');
  }
  const name = md5(image.toString('binary')) + '.png';
  zip.folder("xl/media").file(name, image);
  const rels = zip.folder('xl/worksheets/_rels');
  let rms = [];
  Object.keys(zip.files).forEach(function (filename) {
    if (/^xl\/worksheets\/(sheet.*)/.test(filename)) {
      const sheet = RegExp.$1;
      let id = 'watermark';
      let old = "";
      let content = zip.file(filename).asBinary(true)
        .replace(/((?:<picture\W[^>]+>)?)(\s*<\/worksheet>)/img, ($0, $1, $2) => {
          old = $1 && $1.replace(/^[\s\S]*r:id="(\w+)"[\s\S]*$/, '$1');
          if (old === id) {
            id = 'WM' + Date.now();
          }
          return `<picture r:id="${id}"/>` + $2;
        });
      if (old && content.indexOf(`r:id="${old}"`) > 0) { // 还有引用图片的不删除
        old = "";
      }
      zip.file(filename, Buffer.from(content, 'binary'));
      let rel = rels.file(`${sheet}.rels`);
      rel = (rel && rel.asBinary(true) || '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '</Relationships>').replace(/(<\/Relationships>)/im, relationship(id, name) + '$1');
      if (old) {
        rel = rel.replace(new RegExp(`\\s*<Relationship\\s+Id="${old}"[^>]*>`, 'img'), $0 => {
          if (/Target="([^"]+)"/.test($0)) {
            rms.push(path.relative('/', path.resolve('/' + path.dirname(filename), RegExp.$1)));
          }
          return '';
        });
      }
      rels.file(`${sheet}.rels`, Buffer.from(rel, 'binary'));
    }
  });
  let type = zip.file('[Content_Types].xml').asBinary(true);
  if (!(/<Default\s*Extension="png"\s*ContentType="image\/png"\s*\/>/img.test(type))) {
    zip.file('[Content_Types].xml', Buffer.from(type.replace(/(<Types\s+[^>]*>)/img,
      '$1<Default Extension="png" ContentType="image/png"/>'), 'binary'));
  }
  for (let rm of rms) {
    zip.remove(rm);
  }
  return zip.generate({ type: "nodebuffer" });
}

function remove (buffer) {
  let zip = JSZip();
  zip.load(buffer);
  const rels = zip.folder('xl/worksheets/_rels');
  let rms = [];
  Object.keys(zip.files).forEach(function (filename) {
    if (/^xl\/worksheets\/(sheet.*)/.test(filename)) {
      const sheet = RegExp.$1;
      let old = "";
      let content = zip.file(filename).asBinary(true)
        .replace(/((?:<picture\W[^>]+>)?)(\s*<\/worksheet>)/img, ($0, $1, $2) => {
          old = $1 && $1.replace(/^[\s\S]*r:id="(\w+)"[\s\S]*$/, '$1');
          return $2;
        });
      if (old && content.indexOf(`r:id="${old}"`) > 0) { // 还有引用图片的不删除
        old = "";
      }
      zip.file(filename, Buffer.from(content, 'binary'));
      let rel = rels.file(`${sheet}.rels`);
      if (old && rel) {
        rel = rel.asBinary(true).replace(new RegExp(`\\s*<Relationship\\s+Id="${old}"[^>]*>`, 'img'), $0 => {
          if (/Target="([^"]+)"/.test($0)) {
            rms.push(path.relative('/', path.resolve('/' + path.dirname(filename), RegExp.$1)));
          }
          return '';
        });
        rels.file(`${sheet}.rels`, Buffer.from(rel, 'binary'));
      }
    }
  });
  for (let rm of rms) {
    zip.remove(rm);
  }
  return zip.generate({ type: "nodebuffer" });
}

module.exports = watermark;
watermark.remove = remove;
