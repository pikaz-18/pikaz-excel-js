<!--
 * @Author: zouzheng
 * @Date: 2020-04-30 15:05:31
 * @LastEditors: zouzheng
 * @LastEditTime: 2020-04-30 17:39:34
 * @Description: 这是excel导入组件（页面）
 -->
<template>
  <div class="excel-import-component">
    <input type="file" @change="importFile(this)" id="importFile" style="display: none" accept=".xls,.xlsx" />
    <button @click="uploadFile">导入</button>
  </div>
</template>

<script>
import XLSX from 'xlsx'
export default {
  props: {},
  components: {},
  data () {
    return {
      imFile: '',
      beforeUpload: (e) => {
        // return false
      }
    }
  },
  created () {
  },
  mounted () {
    this.imFile = document.getElementById("importFile")
  },
  methods: {
    uploadFile () {
      // 点击导入按钮
      this.imFile.click();
    },
    /**
     * @name: 导入文件
     * @param {type} 
     * @return: 
     */
    importFile () {
      // 导入excel
      let obj = this.imFile;
      // 导入前
      // if (!this.beforeUpload()) {
      //   return
      // }
      const before = this.beforeUpload()
      if (before === false) {
        return
      }
      console.log(obj.files[0])
      // 无导入文件
      if (!obj.files) {
        return;
      }
      var f = obj.files[0];
      var reader = new FileReader();
      let $t = this;
      // 导入时
      reader.onprogress = function (e) {
        console.log(e)
      }
      reader.onload = function (e) {
        var data = e.target.result;
        if ($t.rABS) {
          $t.wb = XLSX.read(btoa(this.fixdata(data)), {
            // 手动转化
            type: "base64"
          });
        } else {
          $t.wb = XLSX.read(data, {
            type: "binary"
          });
        }
        let json = XLSX.utils.sheet_to_json($t.wb.Sheets[$t.wb.SheetNames[0]]);
        console.log(json)
        $t.dealFile(json); // analyzeData: 解析导入数据
      };
      if (this.rABS) {
        reader.readAsArrayBuffer(f);
      } else {
        reader.readAsBinaryString(f);
      }
    },
    dealFile (data) {
      // 处理导入的数据
      this.imFile.value = "";
      if (data.length <= 0) {
        // 导入失败
      } else {
        //导入成功，处理数据
      }
    },
    /**
     * @name: 文件流转BinaryString
     * @param {type} 
     * @return: 
     */
    fixdata (data) {
      var o = "";
      var l = 0;
      var w = 10240;
      for (; l < data.byteLength / w; ++l) {
        o += String.fromCharCode.apply(
          null,
          new Uint8Array(data.slice(l * w, l * w + w))
        );
      }
      o += String.fromCharCode.apply(null, new Uint8Array(data.slice(l * w)));
      return o;
    },
  },
  computed: {},
  watch: {},
}
</script>

<style scoped>
</style>