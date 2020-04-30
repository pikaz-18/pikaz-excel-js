<!--
 * @Author: zouzheng
 * @Date: 2020-04-30 11:23:07
 * @LastEditors: zouzheng
 * @LastEditTime: 2020-04-30 15:16:19
 * @Description: 这是XXX组件（页面）
 -->
<template>
  <div id="app">
    <div class="ep">
      <excel-export></excel-export>
      <excel-import></excel-import>
    </div>
  </div>
</template>

<script>
import ExcelExport from '../src/plugin/excelExport'
import ExcelImport from '../src/plugin/excelImport'
export default {
  name: 'App',
  components: { ExcelExport, ExcelImport }
}
</script>

<style>
</style>
