<!--
 * @Author: zouzheng
 * @Date: 2020-04-30 11:42:13
 * @LastEditors: zouzheng
 * @LastEditTime: 2020-04-30 15:11:28
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
      // 标题
      title: ['测试', '', '', '', '', '', '', '', ''],
      // 表头
      tHeader: ['测试一', '', '测试二', '测试二', '测试二', '测试二', '测试二', '测试二', '测试二'],
      // 表格数据
      table: [{ NAME: '123', VESSEL_LENGTH: '123', CARGO_NAME: '123', 'DEADWEIGHT_TONNAGE': 1, 'NET_TONNAGE': 2, 'ANCHORAGE_ID': 4, 'EXP_ARCHORAGE_TIME': 5, 'AC_ARCHORAGE_TIME': 6, 'RECOMMEND_BERTH': 3 }],
      // 数据对应的键名
      keys: ['NAME', 'VESSEL_LENGTH', 'CARGO_NAME', 'DEADWEIGHT_TONNAGE', 'NET_TONNAGE', 'ANCHORAGE_ID', 'EXP_ARCHORAGE_TIME', 'AC_ARCHORAGE_TIME', 'RECOMMEND_BERTH'],
      // 合并单元格
      merges: ['A1:I1', 'A2:B2'],
      // 列宽自适应
      autoWidth: true,
      sheetName: '',
      // 文件类型
      bookType: 'xlsx',
      filename: '测试',
      sheet: [
        {
          // 标题
          title: ['测试', '', '', '', '', '', '', '', ''],
          // 表头
          tHeader: ['测试一', '', '测试二', '测试二', '测试二', '测试二', '测试二', '测试二', '测试二'],
          // 表格数据
          table: [{ NAME: '123', VESSEL_LENGTH: '123', CARGO_NAME: '123', 'DEADWEIGHT_TONNAGE': 1, 'NET_TONNAGE': 2, 'ANCHORAGE_ID': 4, 'EXP_ARCHORAGE_TIME': 5, 'AC_ARCHORAGE_TIME': 6, 'RECOMMEND_BERTH': 3 }],
          // 数据对应的键名
          keys: ['NAME', 'VESSEL_LENGTH', 'CARGO_NAME', 'DEADWEIGHT_TONNAGE', 'NET_TONNAGE', 'ANCHORAGE_ID', 'EXP_ARCHORAGE_TIME', 'AC_ARCHORAGE_TIME', 'RECOMMEND_BERTH'],
          // 合并单元格
          merges: ['A1:I1', 'A2:B2'],
          // 列宽自适应
          autoWidth: true,
          sheetName: ''
        }
      ]
    }
  },
  created () {
  },
  mounted () {
  },
  methods: {
    exportExcel () {
      const tHeader = this.tHeader
      const title = this.title
      //表头对应字段
      const filterVal = this.keys
      const list = this.table
      const data = this.formatJson(filterVal, list)
      data.map(item => {
        item.map((i, index) => {
          if (!i) {
            item[index] = ''
          }
        })
      })
      const merges = this.merges
      this.export_json_to_excel({
        title: title,
        header: tHeader,
        data,
        merges,
        filename: this.filename,
        autoWidth: this.autoWidth,
        bookType: this.bookType
      })
    },
    /**
     * @name: 转化数据格式
     * @param {type} 
     * @return: 
     */
    formatJson (filterVal, jsonData) {
      return jsonData.map(v => filterVal.map(j => v[j]))
    },
    /**
     * @name: 转化时间格式
     * @param {type} 
     * @return: 
     */
    datenum (v, date1904) {
      if (date1904) v += 1462;
      var epoch = Date.parse(v);
      return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
    },

    sheet_from_array_of_arrays (data, opts) {
      var ws = {};
      var range = {
        s: {
          c: 1000000000,
          r: 1000000000
        },
        e: {
          c: 0,
          r: 0
        }
      };
      for (var R = 0; R != data.length; ++R) {
        for (var C = 0; C != data[R].length; ++C) {
          if (range.s.r > R) range.s.r = R;
          if (range.s.c > C) range.s.c = C;
          if (range.e.r < R) range.e.r = R;
          if (range.e.c < C) range.e.c = C;
          var cell = {
            v: data[R][C]
          };
          if (cell.v == null) continue;
          var cell_ref = XLSX.utils.encode_cell({
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


    s2ab (s) {
      var buf = new ArrayBuffer(s.length);
      var view = new Uint8Array(buf);
      for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
      return buf;
    },

    /**
     * @name:导出excel 
     * @param {type} 
     * @return: 
     */
    export_json_to_excel ({
      title,
      header,
      data,
      filename,
      merges = [],
      autoWidth = true,
      bookType = 'xlsx',
      SheetName = 'SheetJS'
    } = {}) {
      /* original data */
      filename = filename || 'excel'
      data = [...data]
      data.unshift(header);
      data.unshift(title);

      var ws_name = SheetName;
      var wb = new Workbook(),
        ws = this.sheet_from_array_of_arrays(data);

      if (merges.length > 0) {
        if (!ws['!merges']) ws['!merges'] = [];
        merges.forEach(item => {
          ws['!merges'].push(XLSX.utils.decode_range(item))
        })
      }

      if (autoWidth) {
        /*设置worksheet每列的最大宽度*/
        const colWidth = data.map(row => row.map(val => {
          /*先判断是否为null/undefined*/
          if (val == null) {
            return {
              'wch': 10
            };
          }
          /*再判断是否为中文*/
          else if (val.toString().charCodeAt(0) > 255) {
            return {
              'wch': val.toString().length * 2
            };
          } else {
            return {
              'wch': val.toString().length
            };
          }
        }))
        /*以第一行为初始值*/
        let result = colWidth[0];
        for (let i = 1; i < colWidth.length; i++) {
          for (let j = 0; j < colWidth[i].length; j++) {
            if (result[j]['wch'] < colWidth[i][j]['wch']) {
              result[j]['wch'] = colWidth[i][j]['wch'];
            }
          }
        }
        ws['!cols'] = result;
      }

      /* add worksheet to workbook */
      wb.SheetNames.push(ws_name);
      wb.Sheets[ws_name] = ws;
      var dataInfo = wb.Sheets[wb.SheetNames[0]];

      const borderAll = {  //单元格外侧框线
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
      };
      //给所以单元格加上边框
      for (var i in dataInfo) {
        if (i == '!ref' || i == '!merges' || i == '!cols' || i == 'A1') {

        } else {
          dataInfo[i + ''].s = {
            border: borderAll
          }
        }
      }

      // 去掉标题边框
      let arr = ["A1", "B1", "C1", "D1", "E1", "F1", "G1", "H1", "I1", "J1", "K1", "L1", "M1", "N1", "O1", "P1", "Q1", "R1", "S1", "T1", "U1", "V1", "W1", "X1", "Y1", "Z1"];
      arr.some(function (v) {
        let a = merges[0].split(':')
        if (v == a[1]) {
          dataInfo[v].s = {}
          return true;
        } else {
          dataInfo[v].s = {}
        }
      })

      //设置主标题样式
      dataInfo["A1"].s = {
        font: {
          name: '宋体',
          sz: 18,
          color: { rgb: "ff0000" },
          bold: true,
          italic: false,
          underline: false
        },
        alignment: {
          horizontal: "center",
          vertical: "center"
        },
        // fill: {
        //   fgColor: { rgb: "008000" },
        // },
      };

      var wbout = XLSX.write(wb, {
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
.excel-export-component {
  width: 100%;
  height: 100%;
}
</style>