<!--
 * @Author: zouzheng
 * @Date: 2020-04-30 11:42:13
 * @LastEditors: zouzheng
 * @LastEditTime: 2020-05-06 14:49:24
 * @Description: 这是excel导出组件（页面）
 -->
<template>
  <div class="excel-export-component" @click="exportExcel">
    <button class="btn">导出</button>
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
  props: {},
  components: {},
  data () {
    return {
      // 文件类型
      bookType: 'xlsx',
      // 文件名
      filename: 'excel',
      // 表格配置
      sheet: [
        {
          // 标题
          title: ['测试', '', '', '', '', '', '', '', ''],
          // 表头
          tHeader: ['测试一', '测试一', '测试二322', '测试二', '测试二123', '测试二', '测试二223', '测试二', '测试二'],
          // 表格数据
          table: [{ NAME: '', VESSEL_LENGTH: 'weqeqeeweqewqewqwe', CARGO_NAME: '123', 'DEADWEIGHT_TONNAGE': 1, 'NET_TONNAGE': 2, 'ANCHORAGE_ID': 4, 'EXP_ARCHORAGE_TIME': 5, 'AC_ARCHORAGE_TIME': 6, 'RECOMMEND_BERTH': 3 }],
          // 数据对应的键名
          keys: ['NAME', 'VESSEL_LENGTH', 'CARGO_NAME', 'DEADWEIGHT_TONNAGE', 'NET_TONNAGE', 'ANCHORAGE_ID', 'EXP_ARCHORAGE_TIME', 'AC_ARCHORAGE_TIME', 'RECOMMEND_BERTH'],
          // 合并单元格
          merges: ['A1:I1'],
          // 列宽
          // colWidth: [8, 8, 8, 8, 8, 8, 8, 8, 8],
          // 表名
          sheetName: '表1',
          // 全局样式
          globalStyle: {
            // 边框
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
            // 文字格式
            font: {
              // 字体
              name: '宋体',
              // 字号
              sz: 12,
              // 字体颜色
              color: { rgb: "000000" },
              // 粗体
              bold: false,
              // 斜体
              italic: false,
              // 下划线
              underline: false
            },
            // 对齐方式
            alignment: {
              // 水平方向
              horizontal: "center",
              // 垂直方向
              vertical: "center"
            },
            // 背景色
            fill: {
              fgColor: { rgb: "ffffff" },
            }
          },
          // 单个单元格样式
          cellStyle: []
        }
      ],
      // 默认配置
      default: {
        sheetName: new Date().getTime(),
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
      enum: {
        bookType: ['xlsx', 'xls']
      }
    }
  },
  created () {
  },
  mounted () {
  },
  methods: {
    exportExcel () {
      this.export_json_to_excel()
    },
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
     * @name: 转换类型
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
    export_json_to_excel () {
      if (!this.sheet || this.sheet.length <= 0) {
        return
      }
      const wb = new Workbook()
      this.sheet.forEach((item, index) => {
        let {
          title,
          tHeader,
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
        //表头对应字段
        let data = table.map(v => keys.map(j => v[j]))
        data.unshift(tHeader);
        data.unshift(title);

        const ws = this.sheet_from_array_of_arrays(data);
        if (merges.length > 0) {
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
      saveAs(new Blob([this.s2ab(wbout)], {
        type: "application/octet-stream"
      }), `${filename}.${bookType}`);
    }
  },
  computed: {},
  watch: {},
}
</script>

<style scoped>
</style>