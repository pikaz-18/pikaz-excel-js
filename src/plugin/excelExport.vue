<!--
 * @Author: zouzheng
 * @Date: 2020-04-30 11:42:13
 * @LastEditors: zouzheng
 * @LastEditTime: 2020-05-07 14:08:21
 * @Description: 这是excel导出组件（页面）
 -->
<template>
  <div class="excel-export" @click="exportExcel">
    <slot></slot>
  </div>
</template>

<script>
// workbook对象
function Workbook () {
  if (!(this instanceof Workbook)) return new Workbook();
  this.SheetNames = [];
  this.Sheets = {};
}
import { saveAs } from 'file-saver'
import XLSX from 'yxg-xlsx-style'
export default {
  props: {
    // 文件类型
    bookType: {
      type: String,
      default: 'xlsx'
    },
    // 文件名
    filename: {
      type: String,
      default: 'excel'
    },
    // 表格配置
    sheet: {
      type: Array,
      default: () => {
        return []
      }
    },
    // 处理数据前
    beforeStart: {
      type: Function,
      // bookType:文件类型,filename:文件名,sheet:表格数据
      default: (bookType, filename, sheet) => { }
    },
    // 导出前
    beforeExport: {
      type: Function,
      // filename:文件名,sheet:表格数据,blob:文件流
      default: (bookType, filename, blob) => { }
    },
    onError: {
      type: Function,
      // err:错误信息
      default: (err) => { }
    }
  },
  components: {},
  data () {
    return {
      // 默认配置
      default: {
        sheetName: 'sheet' + new Date().getTime(),
        globalStyle: {
          border: {
            top: {
              style: 'thin'
            },
            bottom: {
              style: 'thin'
            },
            left: {
              style: 'thin'
            },
            right: {
              style: 'thin'
            }
          },
          font: {
            name: '宋体',
            sz: 12,
            color: { rgb: "000000" },
            bold: false,
            italic: false,
            underline: false
          },
          alignment: {
            horizontal: "center",
            vertical: "center"
          },
          fill: {
            fgColor: { rgb: "ffffff" },
          }
        },
      },
      // 枚举类
      enum: {
        // 文件类型
        bookType: ['xlsx', 'xls']
      }
    }
  },
  created () {
  },
  mounted () {
  },
  methods: {
    /**
     * @name: 转化时间格式
     * @param {type} 
     * @return: 
     */
    datenum (v, date1904) {
      if (date1904) v += 1462;
      const epoch = Date.parse(v);
      return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
    },
    /**
     * @name: 设置数据类型
     * @param {type} 
     * @return: 
     */
    sheet_from_array_of_arrays (data, opts) {
      let ws = {};
      const range = {
        s: {
          c: 1000000000,
          r: 1000000000
        },
        e: {
          c: 0,
          r: 0
        }
      };
      for (let R = 0; R != data.length; ++R) {
        for (let C = 0; C != data[R].length; ++C) {
          if (range.s.r > R) range.s.r = R;
          if (range.s.c > C) range.s.c = C;
          if (range.e.r < R) range.e.r = R;
          if (range.e.c < C) range.e.c = C;
          let cell = {
            v: data[R][C]
          };
          if (cell.v == null) continue;
          let cell_ref = XLSX.utils.encode_cell({
            c: C,
            r: R
          });

          if (typeof cell.v === 'number') cell.t = 'n';
          else if (typeof cell.v === 'boolean') cell.t = 'b';
          else if (cell.v instanceof Date) {
            cell.t = 'n';
            cell.z = XLSX.SSF._table[14];
            cell.v = this.datenum(cell.v);
          } else cell.t = 's';

          ws[cell_ref] = cell;
        }
      }
      if (range.s.c < 1000000000) ws['!ref'] = XLSX.utils.encode_range(range);
      return ws;
    },

    /**
     * @name: 转换格式
     * @param {type} 
     * @return: 
     */
    s2ab (s) {
      const buf = new ArrayBuffer(s.length);
      const view = new Uint8Array(buf);
      for (let i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
      return buf;
    },

    /**
     * @name:导出excel 
     * @param {type} 
     * @return: 
     */
    exportExcel () {
      const beforeStart = this.beforeStart(this.bookType, this.filename, this.sheet)
      if (beforeStart === false) {
        return
      }
      if (!this.sheet || this.sheet.length <= 0) {
        this.onError('Table data cannot be empty')
        return
      }
      const wb = new Workbook()
      this.sheet.forEach((item, index) => {
        let {
          title,
          tHeader,
          multiHeader,
          table,
          merges,
          keys,
          colWidth,
          sheetName,
          globalStyle,
          cellStyle
        } = item
        sheetName = sheetName || this.default.sheetName
        globalStyle = globalStyle || this.default.globalStyle
        // 处理标题格式
        if (title || title === 0 || title === '') {
          // 取表头、多级表头中的最大值
          const tHeaderLength = tHeader && tHeader.length || 0
          const multiHeaderLength = multiHeader && Math.max(...multiHeader.map(m => m.length)) || 0
          const titleLength = Math.max(tHeaderLength, multiHeaderLength)
          // 第一个元素为title，剩余以空字符串填充
          title = [title].concat(Array(titleLength - 1).fill(''))
          // 若没有合并，则默认处理标题的合并
          if (!merges) {
            const mergeSecond = String.fromCharCode(64 + titleLength) + '1'
            merges = [`A1:${mergeSecond}`]
          }
        }
        //表头对应字段
        let data = table.map(v => keys.map(j => v[j]))
        tHeader && data.unshift(tHeader);
        title && data.unshift(title);
        const ws = this.sheet_from_array_of_arrays(data);
        if (merges && merges.length > 0) {
          if (!ws['!merges']) ws['!merges'] = [];
          merges.forEach(item => {
            ws['!merges'].push(XLSX.utils.decode_range(item))
          })
        }
        // 如果没有列宽则自适应
        if (!colWidth) {
          // 基准比例，以12为标准
          const benchmarkRate = globalStyle.font.sz / 12
          //设置worksheet每列的最大宽度,并+2调整一点列宽
          const sheetColWidth = data.map(row => row.map(val => {
            /*先判断是否为null/undefined*/
            if (val == null) {
              return {
                'wch': 10 * benchmarkRate + 2
              };
            }
            /*再判断是否为中文*/
            else if (val.toString().charCodeAt(0) > 255) {
              return {
                'wch': val.toString().length * 2 * benchmarkRate + 2
              };
            } else {
              return {
                'wch': val.toString().length * benchmarkRate + 2
              };
            }
          }))
          /*以第一行为初始值*/
          let result = sheetColWidth[0];
          for (let i = 1; i < sheetColWidth.length; i++) {
            for (let j = 0; j < sheetColWidth[i].length; j++) {
              if (result[j]['wch'] < sheetColWidth[i][j]['wch']) {
                result[j]['wch'] = sheetColWidth[i][j]['wch'];
              }
            }
          }
          ws['!cols'] = result;
        } else {
          ws['!cols'] = colWidth.map(i => {
            return { wch: i }
          })
        }

        /* add worksheet to workbook */
        wb.SheetNames.push(sheetName);
        wb.Sheets[sheetName] = ws;
        let dataInfo = wb.Sheets[wb.SheetNames[index]];

        //全局样式
        (function () {
          const { border, font, alignment, fill } = globalStyle;
          Object.keys(dataInfo).forEach(i => {
            if (i == '!ref' || i == '!merges' || i == '!cols') {
            } else {
              dataInfo[i.toString()].s = {
                border,
                font,
                alignment,
                fill
              }
            }
          });
        })();

        // 单个样式
        (function () {
          if (!cellStyle || cellStyle.length <= 0) {
            return
          }
          const { border, font, alignment, fill } = cellStyle;
          cellStyle.forEach(s => {
            dataInfo[s.cell].s = {
              border,
              font,
              alignment,
              fill
            }
          });
        })();
      })
      // 类型默认为xlsx
      let bookType = this.enum.bookType.filter(i => i === this.bookType)[0] || this.enum.bookType[0];
      this.writeExcel(wb, bookType, this.filename)
    },
    /**
     * @name: 导出excel文件
     * @param {type} 
     * @return: 
     */
    writeExcel (wb, bookType, filename) {
      const wbout = XLSX.write(wb, {
        bookType: bookType,
        bookSST: false,
        type: 'binary'
      });
      const blob = new Blob([this.s2ab(wbout)], {
        type: "application/octet-stream"
      })
      const beforeExport = this.beforeExport(blob, bookType, filename)
      if (beforeExport === false) {
        return
      }
      saveAs(blob, `${filename}.${bookType}`);
    }
  },
  computed: {},
  watch: {},
}
</script>

<style scoped>
</style>