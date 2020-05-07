<!--
 * @Author: zouzheng
 * @Date: 2020-04-30 11:23:12
 * @LastEditors: zouzheng
 * @LastEditTime: 2020-05-07 13:46:29
 * @Description: 这是XXX组件（页面）
 -->
## Introduction

这个项目是在工作中并没有找到一个能够快速开箱即用的excel导入导出插件，js里比较知名的xlsx插件没办法修改样式，而xlsx-style插件需要修改源码，都比较麻烦，所以对其做了简单的封装，做到开箱即用。

## Features

* 支持导出excel文件，并可设置列宽，边框，字体，字体颜色，字号，对齐方式，背景色等样式。
* 支持excel文件导入，生成json数据。

## Installation

### With npm or yarn 

```bash
yarn add pikaz-excel-js

npm i -S pikaz-excel-js
```

**请确保vue版本在2.0以上**

## For Vue-cli

### Export:

#### Typical use:
``` html
<excel-export :sheet="sheet">
   <div>导出</div>
</excel-export>
```
.vue file:
``` js
  import {ExcelExport} from 'pikaz-excel-js'
  ...
  export default {
        components: {
            ExcelExport,
        },
        data () {
          return {
            sheet:[
              [
                title:"水果的味道",
                tHeader:["荔枝","柠檬"],
                table:[{litchi:"甜",lemon:"酸"}],
                keys:["litchi","lemon"],
                sheetName:"水果的味道",
              ]
            ]
          }
        }
  ...
```
#### Attributes:
参数|说明|类型|可选值|默认值
-|-|-|-|-
bookType|文件格式|string|xlsx/xls|xlsx
filename|文件名称|string|--|excel
sheet|表格数据，具体看下方[表格配置](#table-setting)|array|--|--
before-start|处理数据之前的钩子，参数为导出的文件格式，文件名，表格数据，若返回 false则停止导出|function(bookType, filename, sheet)|--|--
before-export|导出文件之前的钩子，参数为导出的文件格式，文件名，blob文件流，若返回 false则停止导出|function(bookType, filename, sheet)|--|--
on-error|导出失败的钩子，参数为错误信息|function(err)|--|--

<h5 id="table-setting">表格参数配置</h5>

参数|说明|类型|可选值|默认值
-|-|-|-|-
title|表格标题|string|--|--
tHeader|表头|array|--|--
multiHeader|多级表头,即一个数组中有多个表头数组|array|--|--
table|表格数据array|--|--
merges|合并单元格|array|--|--
keys|数据键名，需与表头内容顺序对应|array|--|--
colWidth|列宽，若不传，则列宽自适应|array|--|--
sheetName|表格名称|string|--|sheet
globalStyle|表格全局样式，具体参数查看下方[表格全局样式](#global-style)|object|--|[表格全局样式](#global-style)
cellStyle|单元格样式，具体参数查看下方[单元格样式](#cell-style)|array|--|--

<h5 id="global-style">表格全局样式</h5>

参数|说明|格式|类型|可选值|默认值
---|--|-|--|--
border|单元格边框,格式如:{top:{style: 'thin'},bottom: {
style: 'thin'},left: {style: 'thin'},right:{style:'thin'}}，可更改style值改变外边框，传{}则取消边框|object|style:thin/dotted|style:'thin'
font|合并单元格|array|--|--



<h5 id="cell-style">单元格样式</h5>

### Import:

#### Typical use:
``` html
<excel-import :on-success="onSuccess">
   <div>导入</div>
</excel-import>
```
.vue file:
``` js
  import {ExcelImport} from 'pikaz-excel-js'
  ...
  export default {
        components: {
            ExcelImport,
        },
        methods:{
          onSuccess(data, file){
            console.log(data)
          }
        }
  ...
```

#### Attributes:
