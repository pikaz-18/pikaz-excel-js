<!--
 * @Author: zouzheng
 * @Date: 2020-05-07 14:57:19
 * @LastEditors: zouzheng
 * @LastEditTime: 2020-05-08 14:23:19
 * @Description: 这是exportDemo组件（页面）
 -->
<template>
  <div class="export-demo">
    <excel-export :bookType="bookType" :filename="filename" :sheet="sheet" :on-error="onError">
      <div class="export-btn">点击导出</div>
    </excel-export>
  </div>
</template>

<script>
import { ExcelExport } from '../../plugin'
export default {
  props: {},
  components: { ExcelExport },
  data () {
    return {
      bookType: 'xlsx',
      filename: 'export-demo',
      sheet: [
        {
          title: '插件信息',
          tHeader: ['作者', '语言', '插件名称', '插件用途'],
          table: [{ author: 'pikaz', language: 'javascript', name: 'pikaz-excel-js', application: 'excel的导入导出' }],
          keys: ['author', 'language', 'name', 'application'],
          sheetName: '插件信息',
          cellStyle: [
            {
              cell: 'A1',
              font: {
                sz: 14,
                color: { rgb: "ffffff" },
                bold: true,
              },
              fill: {
                fgColor: { rgb: "ff7e00" },
              }
            }
          ]
        },
        {
          title: '插件信息2',
          multiHeader: [['基础信息', '', '', '详细信息'], ['作者', '语言', '插件名称', '插件用途']],
          table: [{ author: 'pikaz', language: 'javascript', name: 'pikaz-excel-js', application: 'excel的导入导出' }],
          keys: ['author', 'language', 'name', 'application'],
          merges: ['A2:C2'],
          colWidth: [16, 16, 16, 20],
          sheetName: '插件信息2',
          globalStyle: {
            font: {
              color: { rgb: "ff7e00" },
            }
          },
          cellStyle: [
            {
              cell: 'A1',
              font: {
                name: '宋体',
                sz: 14,
                color: { rgb: "ffffff" },
                bold: true
              },
              fill: {
                fgColor: { rgb: "ff7e00" },
              }
            }
          ]
        }
      ]
    }
  },
  created () {
  },
  mounted () {
  },
  methods: {
    /**
     * @name: 导出错误
     * @param {String} err/错误信息 
     * @return: 
     */
    onError (err) {
      console.log(err)
    }
  },
  computed: {},
  watch: {},
}
</script>

<style scoped>
.export-demo {
  display: flex;
  justify-content: space-between;
}
.export-btn {
  display: flex;
  justify-content: center;
  align-items: center;
  width: 80px;
  height: 30px;
  background: #ff7e00;
  border-radius: 5px;
  color: #ffffff;
  cursor: pointer;
}
.code {
  margin-left: 100px;
  width: calc(100% - 100px);
  background: #eee;
}
</style>