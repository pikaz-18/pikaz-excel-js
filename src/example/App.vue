<!--
 * @Author: zouzheng
 * @Date: 2020-04-30 11:23:07
 * @LastEditors: zouzheng
 * @LastEditTime: 2020-05-07 14:59:24
 * @Description: 这是XXX组件（页面）
 -->
<template>
  <div id="app">
    <div class="header">
      <img class="logo" src="./assets/img/logo.png" />
    </div>
    <h2>导出demo</h2>
    <div></div>
  </div>
</template>

<script>
import ExportDemo from './components/exportDemo'
export default {
  name: 'App',
  components: { ExportDemo }
}
</script>

<style>
#app {
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
</style>

