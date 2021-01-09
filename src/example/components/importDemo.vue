<!--
 * @Author: zouzheng
 * @Date: 2020-05-08 11:34:18
 * @LastEditors: zouzheng
 * @LastEditTime: 2020-05-08 13:47:55
 * @Description: 这是XXX组件（页面）
 -->
<template>
  <div class="import-demo">
    <excel-import :on-success="onSuccess">
      <div class="import-btn">点击导入</div>
    </excel-import>
    <div class="code" v-if="json">
      {{json}}
    </div>
  </div>
</template>

<script>
import { ExcelImport } from '../../plugin'
export default {
  props: {},
  components: { ExcelImport },
  data () {
    return {
      json: ''
    }
  },
  created () {
  },
  mounted () {
  },
  methods: {
    /**
     * @name: 导入成功
     * @param {Array} response/生导入数据 
     * @param {Object} file/导入文件 
     * @return: 
     */
    onSuccess (response, file) {
      this.json = JSON.stringify(response)
    }
  },
  computed: {},
  watch: {},
}
</script>

<style scoped>
.import-demo {
  display: flex;
  justify-content: space-between;
}
.import-btn {
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
  padding: 20px 0;
  margin-left: 100px;
  width: calc(100% - 100px);
  background: #eee;
}
</style>