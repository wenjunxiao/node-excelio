declare namespace ExcelIO {
  /**
   * 边框风格
   */
  type BorderStyle =
    | 'thin'
    | 'medium'
    | 'thick'
    | 'dotted'
    | 'hair'
    | 'dashed'
    | 'mediumDashed'
    | 'dashDot'
    | 'mediumDashDot'
    | 'dashDotDot'
    | 'mediumDashDotDot'
    | 'slantDashDot';

  interface BorderOptions {
    /**
     * 颜色，比如`#000000`
     */
    color: 'string';
    /**
     * 边框风格
     */
    style: BorderStyle;
  }

  /**
   * 边框选项
   */
  interface BorderOption {
    /**
     * 外边框
     */
    outer?: BorderOptions;
    /**
     * 内边框
     */
    inner?: BorderOptions;
  }
  /**
   * 单元格数据类型
   */
  type CellType =
    | 'n' // number
    | 'b' // boolean
    | 'd' // date
    | 's' // string
    ;
  /**
   * 单元格选项
   */
  interface CellOption {
    /**
     * 对齐方式，比如`{horizontal: 'left',vertical: 'left'}`
     */
    alignment?: any;
    /**
     * 字体，比如`{bold: true, name: '微软雅黑', sz: 10}`
     */
    font?: any;
    /**
     * 单元格宽度
     */
    width?: number;
    /**
     * 背景颜色
     */
    bgColor?: string;
    /**
     * 字体颜色
     */
    fgColor?: string;
    /**
     * 单元格数据类型
     */
    type?: CellType;
  }
  type BuildType =
    | 'binary'
    | 'base64'
    | 'buffer'
    ;
  interface BuildOptions {
    /**
     * 构建类型，默认`binary`
     */
    type: BuildType;
  }
  interface ExcelOption {
    /**
     * 文档类型，默认`xlsx`
     */
    bookType?: string;
    /**
     * 构建类型，默认`binary`
     */
    type: BuildType;
    /**
     * 默认对齐方式：`{horizontal: 'left',vertical: 'left'}`
     */
    alignment?: object;
    /**
     * 宽度是否使用像素值，如果`true`则传入的宽度都必须是像素，否则传入的宽度的是指字符数(注意中文)
     */
    px?: boolean;
    /**
     * 默认字体大小，默认10
     */
    fontSize?: number;
    /**
     * 默认宽度（像素还是字符数，取决于`px`）
     */
    width?: number;
    /**
     * 最小宽度（像素还是字符数，取决于`px`）
     */
    minWidth?: number;
    /**
     * 数字列为空时的表示符号，比如`-`
     */
    NaN?: string;
    /**
     * 是否显示表格线
     */
    showGridLines?: boolean;
    /**
     * 标题行行号，比如`0`
     */
    titleLine?: number;
    /**
     * 标题行的默认选项，`title()`时使用或者指定`titleLine`了之后匹配的行时使用
     */
    titleOpts?: CellOption;
    /**
     * 单元格默认选项，填充单元格时自动使用
     */
    cellOpts?: CellOption;
    /**
     * 是否增加边框
     */
    border2end?: boolean;
  }

  interface SheetOption {
    /**
     * Sheet名称
     */
    sheetName?: string;
    /**
     * 默认对齐方式：`{horizontal: 'left',vertical: 'left'}`
     */
    alignment?: object;
    /**
     * 宽度是否使用像素值，如果`true`则传入的宽度都必须是像素，否则传入的宽度的是指字符数(注意中文)
     */
    px?: boolean;
    /**
     * 默认字体大小，默认10
     */
    fontSize?: number;
    /**
     * 默认宽度（像素还是字符数，取决于`px`）
     */
    width?: number;
    /**
     * 最小宽度（像素还是字符数，取决于`px`）
     */
    minWidth?: number;
    /**
     * 数字列为空时的表示符号，比如`-`
     */
    NaN?: string;
    /**
     * 标题行行号，比如`0`
     */
    titleLine?: number;
    /**
     * 标题行的默认选项，`title()`时使用或者指定`titleLine`了之后匹配的行时使用
     */
    titleOpts?: CellOption;
    /**
     * 单元格默认选项，填充单元格时自动使用
     */
    cellOpts?: CellOption;
    /**
     * 是否增加边框
     */
    border2end?: boolean;
  }

  type CellOptions = number | CellOption;
  interface Sheet {
    /**
     * 清除Sheet页内容
     */
    clear (): this;

    /**
     * 重命名Sheet页名称
     * @param name 名称
     */
    rename (name: string): this;

    /**
     * 当前行号
     */
    rowIndex (): number;

    /**
     * 当前列号
     */
    colIndex (): number;

    /**
     * 跳过指定的行数，并定位到指定的单元格
     * @param {number} [rows = 1] 行数
     * @param {number} [cells = 0] 单元格数
     */
    skipRow (rows?: number, cells?: number): this;
    /**
     * 跳过指定的单元格数
     * @param {number} [cells = 1] 单元格数
     */
    skipCell (cells?: number): this;
    /**
     * 跳转到指定的行的单元格
     * @param {number} row 指定的行
     * @param {number} [cell = 0] 指定的单单元格
     */
    go (row: number, cell?: number): this;
    /**
     * 跳转到下一行指定的单元格（默认行首）
     * @param {number} [cells = 0] 指定的单格
     */
    row (cells?: number): this;
    /**
     * 跳转到下一个单元格，并写入数字
     * @param v 数字
     * @param options 单元格选项
     */
    number (v: number, options?: CellOptions): this;
    /**
     * 跳转到下一个单元格，并写入布尔值
     * @param v 布尔值
     * @param options 
     */
    boolean (v: boolean, options?: CellOptions): this;
    /**
     * 跳转到下一个单元格，并写入字符串
     * @param v 字符串
     * @param options 单元格选项
     */
    string (v: string, options?: CellOptions): this;
    /**
     * 跳转到下一个单元格，并写入UTC时间
     * @param v 时间
     * @param format 时间格式化字符串
     * @param options 单元格选项
     */
    utc (v: string | Date, format?: string, options?: CellOptions): this;
    /**
     * 跳转到下一个单元格，并写入当前时区的时间
     * @param v 时间
     * @param format 时间格式化字符串
     * @param options 单元格选项
     */
    date (v: string | Date, format?: string, options?: CellOptions): this;
    /**
     * 跳转到下一个单元格，并写入金额
     * @param v 金额
     * @param currency 币种
     * @param precision 
     * @param options 
     */
    currency (v: number | string, currency?: string, precision?: number, options?: CellOptions): this;
    /**
     * 跳转到下一个单元格，并写入百分数
     * @param v 百分数
     * @param precision 精度
     * @param options 单元格选项
     */
    percent (v: any, precision?: number, options?: CellOptions): this;
    /**
     * 在当前行设置标题
     * @param vs 标题数组
     * @param options 单元格选项
     */
    titles (vs: string[], options?: CellOptions): this;
    /**
     * 在当前单元格写入标题
     * @param v 标题
     * @param options 单元格选项
     */
    title (v: string, options?: CellOptions): this;
    /**
     * 填充当前数据行
     * @param vs 数据
     * @param options 单元格选项
     */
    fillRow (vs: any[], options?: CellOptions): this;
    /**
     * 填充整个表格
     * @param data 二维数据数组
     * @param options 单元格选项
     */
    fill (data?: any[][], options?: CellOptions | CellOptions[]): this;
    /**
     * 跳转到下一个单元格，并写入值
     * @param v 值
     * @param options 单元格选项
     * @param type 单元格类型
     * @param format 单元格格式
     */
    cell (v: any, options?: CellOptions, type?: string, format?: string): this;
    /**
     * 按照中文方式设置单元格宽度
     * @param width 中文字数或者像素值
     * @param {number} [colIndex = -1] 指定列，默认当前列
     */
    chWidth (width: number, colIndex?: number): this;
    /**
     * 设置单元格宽度
     * @param width 字符数或者像素值
     * @param {number} [colIndex = -1] 指定列，默认当前列
     */
    width (width: number, colIndex?: number): this;
    /**
     * 设置当前单元格颜色
     * @param bgColor 背景色
     * @param fgColor 前景色
     */
    color (bgColor: string, fgColor: string): this;
    /**
     * 设置单元格前景色（字体颜色）
     * @param color 颜色
     */
    fgColor (color: string): this;
    /**
     * 设置单元格背景色
     * @param color 颜色
     */
    bgColor (color: string): this;
    /**
     * 添加水印
     * @param image 水印图片
     */
    watermark (image: any): this;
    /**
     * 不使用水印
     */
    withoutWatermark (): this;
    /**
     * 设置表格指定单元格的边框
     * @param rs 起始行
     * @param cs 起始列
     * @param re 结束行
     * @param ce 结束列
     * @param {string} [color=#000000] 边框颜色
     * @param {BorderStyle} [style = 'thin'] 边框样式
     * @param options 
     */
    border (rs: number, cs: number, re: number, ce: number, color?: string, style?: BorderStyle, options?: BorderOption): this;
    /**
     * 从指定位置开始设置整个表格边框
     * @param r 起始行
     * @param c 起始列
     * @param {string} [color=#000000] 边框颜色
     * @param {BorderStyle} [style = 'thin'] 边框样式
     * @param options 
     */
    border2end (r: number, c: number, color?: string, style?: BorderStyle, options?: {}): this;
    /**
     * 合并单元格
     * @param {number} [cells = 1] 要合并的单元格数
     */
    mergeCell (cells?: number): this;
    /**
     * 合并行
     * @param {number} [rows = 1] 要合并的行数
     */
    mergeRow (rows?: number): this;
    /**
     * 合并指定的单元格
     * @param rs 起始行
     * @param cs 起始列
     * @param re 结束行
     * @param ce 结束列
     */
    merge (rs: number, cs: number, re: number, ce: number): this;
    /**
     * 表格构建完成
     */
    end (): this;
  }
  type SheetConstructor = new (options?: ExcelOption) => Sheet;
  interface ExcelWriter {
    /**
     * 去掉默认表格线
     */
    withoutGridLines (): this;
    /**
     * 当前行号
     */
    rowIndex (): number;
    /**
     * 当前列号
     */
    colIndex (): number;
    /**
     * 新增Sheet
     * @param name 名称
     */
    newSheet (name: string): Sheet;
    /**
     * 切换到指定Sheet，如果不存在则新增
     */
    sheet (name: string): this;
    /**
     * 获取当前Sheet
     */
    active (): Sheet;
    /**
     * 重命名Sheet
     * @param name 新名称
     * @param from 原名称
     */
    rename (name: string, from: string): this;
    /**
     * 跳过指定的行数，并定位到指定的单元格
     * @param {number} [rows = 1] 行数
     * @param {number} [cells = 0] 单元格数
     */
    skipRow (rows?: number, cells?: number): this;
    /**
     * 跳过指定的单元格数
     * @param {number} [cells = 1] 单元格数
     */
    skipCell (cells?: number): this;
    /**
     * 跳转到指定的行的单元格
     * @param {number} row 指定的行
     * @param {number} [cell = 0] 指定的单单元格
     */
    go (row: number, cell?: number): this;
    /**
     * 跳转到下一行指定的单元格（默认行首）
     * @param {number} [cells = 0] 指定的单格
     */
    row (cells?: number): this;
    /**
     * 跳转到下一个单元格，并写入数字
     * @param v 数字
     * @param options 单元格选项
     */
    number (v: number, options?: CellOptions): this;
    /**
     * 跳转到下一个单元格，并写入布尔值
     * @param v 布尔值
     * @param options 
     */
    boolean (v: boolean, options?: CellOptions): this;
    /**
     * 跳转到下一个单元格，并写入字符串
     * @param v 字符串
     * @param options 单元格选项
     */
    string (v: string, options?: CellOptions): this;
    /**
     * 跳转到下一个单元格，并写入UTC时间
     * @param v 时间
     * @param format 时间格式化字符串
     * @param options 单元格选项
     */
    utc (v: string | Date, format?: string, options?: CellOptions): this;
    /**
     * 跳转到下一个单元格，并写入当前时区的时间
     * @param v 时间
     * @param format 时间格式化字符串
     * @param options 单元格选项
     */
    date (v: string | Date, format?: string, options?: CellOptions): this;
    /**
     * 跳转到下一个单元格，并写入金额
     * @param v 金额
     * @param currency 币种
     * @param precision 
     * @param options 
     */
    currency (v: number | string, currency?: string, precision?: number, options?: CellOptions): this;
    /**
     * 跳转到下一个单元格，并写入百分数
     * @param v 百分数
     * @param precision 精度
     * @param options 单元格选项
     */
    percent (v: any, precision?: number, options?: CellOptions): this;
    /**
     * 在当前行设置标题
     * @param vs 标题数组
     * @param options 单元格选项
     */
    titles (vs: string[], options?: CellOptions): this;
    /**
     * 在当前单元格写入标题
     * @param v 标题
     * @param options 单元格选项
     */
    title (v: string, options?: CellOptions): this;
    /**
     * 填充当前数据行
     * @param vs 数据
     * @param options 单元格选项
     */
    fillRow (vs: any[], options?: CellOptions): this;
    /**
     * 填充整个表格
     * @param data 二维数据数组
     * @param options 单元格选项
     */
    fill (data?: any[][], options?: CellOptions | CellOptions[]): this;
    /**
     * 跳转到下一个单元格，并写入值
     * @param v 值
     * @param options 单元格选项
     * @param type 单元格类型
     * @param format 单元格格式
     */
    cell (v: any, options?: CellOptions, type?: string, format?: string): this;
    /**
     * 按照中文方式设置单元格宽度
     * @param width 中文字数或者像素值
     * @param {number} [colIndex = -1] 指定列，默认当前列
     */
    chWidth (width: number, colIndex?: number): this;
    /**
     * 设置单元格宽度
     * @param width 字符数或者像素值
     * @param {number} [colIndex = -1] 指定列，默认当前列
     */
    width (width: number, colIndex?: number): this;
    /**
     * 设置当前单元格颜色
     * @param bgColor 背景色
     * @param fgColor 前景色
     */
    color (bgColor: string, fgColor: string): this;
    /**
     * 设置表格指定单元格的边框
     * @param rs 起始行
     * @param cs 起始列
     * @param re 结束行
     * @param ce 结束列
     * @param {string} [color=#000000] 边框颜色
     * @param {BorderStyle} [style = 'thin'] 边框样式
     * @param options 
     */
    border (rs: number, cs: number, re: number, ce: number, color?: string, style?: BorderStyle, options?: BorderOption): this;
    /**
     * 从指定位置开始设置整个表格边框
     * @param r 起始行
     * @param c 起始列
     * @param {string} [color=#000000] 边框颜色
     * @param {BorderStyle} [style = 'thin'] 边框样式
     * @param options 
     */
    border2end (r: number, c: number, color?: string, style?: BorderStyle, options?: {}): this;
    /**
     * 合并单元格
     * @param {number} [cells = 1] 要合并的单元格数
     */
    mergeCell (cells?: number): this;
    /**
     * 合并行
     * @param {number} [rows = 1] 要合并的行数
     */
    mergeRow (rows?: number): this;
    /**
     * 合并指定的单元格
     * @param rs 起始行
     * @param cs 起始列
     * @param re 结束行
     * @param ce 结束列
     */
    merge (rs: number, cs: number, re: number, ce: number): this;
    /**
     * 添加水印
     * @param image 水印图片
     */
    watermark (image: any): this;
    /**
     * 不使用水印
     */
    withoutWatermark (): this;
    /**
     * 完成当前Sheet
     */
    endSheet (): this;
    /**
     * 构建并返回数据，默认
     * @param {BuildOptions} options 
     */
    build (options?: BuildOptions): this;
    /**
     * 保存到文件
     * @param filename 文件名
     * @param options 
     */
    save (filename: string, options?: {}): this;
  }
  type ExcelWriterConstructor = new (options?: ExcelOption) => ExcelWriter;

  /**
   * 创建一个ExcelWriter
   * @param options 选项
   */
  function createWriter (options?: ExcelOption): ExcelWriter;

  interface ExcelReader {
    /**
     * 从文件读入数据
     * @param filename 文件
     */
    readFile (filename?: string): this;
    /**
     * 从文件内容读取数据
     * @param {Buffer|String} data 数据
     */
    read (data: string | Buffer): this;
    /**
     * 获取所有Sheet的名字
     */
    sheetNames (): string[];
    /**
     * 切换到指定的Sheet
     * @param name 名字
     */
    sheet (name?: string): this;
    /**
     * 获取当前行的数据
     */
    row (): any[];
    /**
     * 映射标题和字段的关联关系
     * @param titles 必须字段标题映射关系
     * @param opts 可选字段标题映射关系
     * @param {function(name, i)} mapper 重复标题时的映射关系
     * @example titles
     * `{
        "字段": "表格中的标题", // 不指定数据类型
        "时间": ["表格中的标题", "date"] // 指定数据类型
      }`
     * @example mapper
     `const mapper = (name, i) => {
        switch (name) {
          case '分类':
            if (i === 1) {
              return '一级分类';
            } else if (i === 2) {
              return '二级分类';
            }
            break;
        }
      };`
     */
    header (titles: object, opts?: object, mapper?: function): this;
    /**
     * 是否存在标题
     * @param title 标题
     */
    has (title: string): boolean;
    /**
     * 每一行数据的处理，默认返回当前行数据
     * @param fn 处理函数
     */
    map (fn: function): any[];
  }
  interface ReaderOption {
    /**
     * 内容是否是格式化显示的
     */
    formatted?: boolean;
    /**
     * 类型映射关系
     */
    types?: object;
  }
  type ExcelReaderConstructor = new (options?: ReaderOption) => ExcelReader;
  /**
   * 创建一个ExcelReader
   * @param options 选项
   */
  function createReader (options?: ReaderOption): ExcelReader;
}

export = ExcelIO;
