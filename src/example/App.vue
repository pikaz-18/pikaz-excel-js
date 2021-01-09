<!--
 * @Author: zouzheng
 * @Date: 2020-04-30 11:23:07
 * @LastEditors: zouzheng
 * @LastEditTime: 2020-05-08 14:23:30
 * @Description: 这是XXX组件（页面）
 -->
<template>
  <div id="app">
    <div class="header">
      <img class="logo" src="./assets/img/logo.png" />
    </div>
    <div class="demo">
      <h3>导出demo</h3>
      <export-demo></export-demo>
    </div>
    <div class="demo">
      <h3>导入demo</h3>
      <import-demo></import-demo>
    </div>
    <div class="address">
      <a href="https://github.com/pikaz-18/pikaz-excel-js/tree/master/src/example/components">demo代码地址</a>
    </div>
  </div>
</template>

<script>
import ExportDemo from './components/exportDemo'
import ImportDemo from './components/importDemo'
export default {
  name: 'App',
  components: { ExportDemo, ImportDemo }
}
</script>

<style>
#app {
  display: flex;
  flex-direction: column;
  align-items: center;
  width: 100vw;
}
.header {
  display: flex;
  justify-content: center;
  align-items: center;
  background: #eee;
  width: 100%;
  height: 230px;
}
.logo {
  width: 600px;
  object-fit: contain;
}
.demo {
  margin: 20px 0;
  width: 1200px;
}
.demo > h3 {
  margin-bottom: 20px;
}
.address {
  display: flex;
  margin: 20px 0;
  width: 1200px;
  font-weight: bold;
}
.address > a {
  color: #ff7e00;
}
</style>

