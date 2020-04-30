<!--
 * @Author: zouzheng
 * @Date: 2020-04-30 15:05:31
 * @LastEditors: zouzheng
 * @LastEditTime: 2020-04-30 15:15:36
 * @Description: 这是excel导入组件（页面）
 -->
<template>
  <div>
    <input type="file" @change="importFile(this)" id="imFile" style="display: none"
      accept="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel" />
    <Button type="primary" @click="uploadFile">导入</Button>
  </div>
</template>

<script>
export default {
  props: {},
  components: {},
  data () {
    return {
      fullscreenLoading: false, // 加载中
      imFile: "", // 导入文件el
      errorMsg: "", // 错误信息内容
    }
  },
  created () {
  },
  mounted () {
    this.imFile = document.getElementById("imFile");
  },
  methods: {
    uploadFile: function () {
      // 点击导入按钮
      this.imFile.click();
    },

    importFile: function () {
      // 导入excel
      this.fullscreenLoading = true;
      let obj = this.imFile;
      if (!obj.files) {
        this.fullscreenLoading = false;
        return;
      }
      var f = obj.files[0];
      var reader = new FileReader();
      let $t = this;
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
    dealFile: function (data) {
      // 处理导入的数据
      this.imFile.value = "";
      this.fullscreenLoading = false;
      if (data.length <= 0) {
        this.errorMsg = "请导入正确信息";
      } else {
        //导入成功，处理数据
      }
    },

    fixdata: function (data) {
      // 文件流转BinaryString
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

<style lang='less' scoped>
</style>