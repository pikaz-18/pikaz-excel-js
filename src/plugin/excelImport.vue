<!--
 * @Author: zouzheng
 * @Date: 2020-04-30 15:05:31
 * @LastEditors: zouzheng
 * @LastEditTime: 2020-05-06 17:25:57
 * @Description: 这是excel导入组件（页面）
 -->
<template>
  <div class="excel-import" @click="importFileClick">
    <input type="file" @change="importFile(this)" id="importFile" style="display: none" accept=".xls,.xlsx" />
    <slot></slot>
    <button>导入</button>
  </div>
</template>

<script>
import XLSX from 'xlsx'
export default {
  props: {
    //  导入前
    beforeImport: {
      type: Function,
      default: () => { }
    },
    // 导入时
    onProgress: {
      type: Function,
      default: () => { }
    },
    // 状态改变
    onChange: {
      type: Function,
      default: () => { }
    },
    onSuccess: {
      type: Function,
      default: () => { }
    },
    onError: {
      type: Function,
      default: () => { }
    }
  },
  components: {},
  data () {
    return {
      imFile: '',
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
    this.imFile = document.getElementById("importFile")
  },
  methods: {
    /**
     * @name: 点击导入按钮
     * @param {type} 
     * @return: 
     */
    importFileClick () {
      this.imFile.click();
    },
    /**
     * @name: 导入文件
     * @param {type} 
     * @return: 
     */
    importFile () {
      // 导入excel
      const obj = this.imFile;
      // 无导入文件
      if (!obj.files) {
        this.onError('No imported file')
        return
      }
      const file = obj.files[0];
      // 导入前
      const beforeImport = this.beforeImport(file)
      this.onChange(file)
      if (beforeImport === false) {
        return
      }
      // 文件类型必须为xlsx或者xls
      const bookType = file.name.substr(file.name.length - 4, file.name.length - 1)
      const type = this.emum.bookType.some(e => {
        if (bookType.indexOf(e)) {
          return true
        }
        return false
      })
      if (!type) {
        this.onError('The file type must be "xlsx" or "xls"', file)
        return
      }
      const reader = new FileReader();
      const $t = this;
      // 导入时
      reader.onprogress = e => {
        this.onProgress(e, file)
      }
      // 导入完成
      reader.onload = e => {
        const data = e.target.result;
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
        $t.dealFile(json, file); // 解析导入数据
      };
      if (this.rABS) {
        reader.readAsArrayBuffer(file);
      } else {
        reader.readAsBinaryString(file);
      }
    },
    /**
     * @name: 处理导入的数据
     * @param {type} 
     * @return: 
     */
    dealFile (data, file) {
      this.imFile.value = "";
      if (data.length <= 0) {
        // 导入失败
        this.onChange(file)
        this.onError('The import failed', file)
        return
      } else {
        //导入成功
        this.onChange(file)
        this.onSuccess(data, file)
        return
      }
    },
    /**
     * @name: 文件流转BinaryString
     * @param {type} 
     * @return: 
     */
    fixdata (data) {
      const o = "";
      const l = 0;
      const w = 10240;
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