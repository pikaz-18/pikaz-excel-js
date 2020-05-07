<!--
 * @Author: zouzheng
 * @Date: 2020-04-30 11:23:12
 * @LastEditors: zouzheng
 * @LastEditTime: 2020-05-07 09:57:32
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

### Typical use:

#### Export:
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
  ...
```

#### Import:
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
  ...
```
