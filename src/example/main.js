/*
 * @Author: zouzheng
 * @Date: 2020-04-30 11:23:07
 * @LastEditors: zouzheng
 * @LastEditTime: 2020-05-07 14:43:33
 * @Description: 这是XXX组件（页面）
 */
import Vue from 'vue'
import App from './App.vue'
import './assets/css/base.css'

Vue.config.productionTip = false

new Vue({
  render: h => h(App),
}).$mount('#app')
