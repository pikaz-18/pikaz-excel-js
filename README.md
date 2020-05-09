<!--
 * @Author: zouzheng
 * @Date: 2020-04-30 11:23:12
 * @LastEditors: zouzheng
 * @LastEditTime: 2020-05-09 17:26:38
 * @Description: 这是XXX组件（页面）
 -->
## Introduction

这个项目是在工作中并没有找到一个开箱即用的excel导入导出插件，js里比较知名的[xlsx](https://github.com/SheetJS/sheetjs.git)插件免费版没办法修改样式，而[xlsx-style](https://github.com/protobi/js-xlsx.git)插件需要修改源码，都比较麻烦，所以在xlsx和xlsx-style的基础上做了简单的封装，做到开箱即用，降低使用成本。

## Features

* 支持导出excel文件，并可设置列宽，边框，字体，字体颜色，字号，对齐方式，背景色等样式。
* 支持excel文件导入，生成json数据，考虑到客户端机器性能，导出大量数据时，推荐拆分数据分成多个文件导出。

## [demo](https://pikaz-18.github.io/pikaz-excel-js/example/index.html)

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
manual|手动导出模式，设置为true时，取消点击导出，并可调用[pikaExportExcel](#export-method)方法完成导出|boolean|true/false|false
sheet|表格数据，每个表格数据对象配置具体看下方[表格配置](#table-setting)|array|--|--
before-start|处理数据之前的钩子，参数为导出的文件格式，文件名，表格数据，若返回 false则停止导出|function(bookType, filename, sheet)|--|--
before-export|导出文件之前的钩子，参数为导出的文件格式，文件名，blob文件流，若返回 false则停止导出|function(bookType, filename, sheet)|--|--
on-error|导出失败的钩子，参数为错误信息|function(err)|--|--

<h5 id="table-setting">表格参数配置</h5>

参数|说明|类型|可选值|默认值
-|-|-|-|-
title|表格标题，自动设置合并，非必须项|string|--|--
tHeader|表头，非必须项|array|--|--
multiHeader|多级表头,即一个数组中包含多个表头数组，非必须项|array|--|--
table|表格数据|array|--|--
merges|合并单元格，合并的表头和表格多余数据项以空字符串填充，非必须项|array|--|--
keys|数据键名，需与表头内容顺序对应|array|--|--
colWidth|列宽，若不传，则列宽自适应，数据量多时推荐设置列宽|array|--|--
sheetName|表格名称|string|--|sheet
globalStyle|表格全局样式，具体参数查看下方[表格全局样式](#global-style)|object|--|[表格全局样式](#global-style)
cellStyle|单元格样式，每个单元格对象配置具体参数查看下方[单元格样式](#cell-style)|array|--|--

<h5 id="global-style">表格全局样式</h5>

<table>
    <tr>
        <td>参数</td>
        <td>属性名</td>
        <td>说明</td>
        <td>类型</td>
        <td>可选值</td>
        <td>默认值</td>
    </tr>
    <tr>
        <td rowspan="4">border</td>
        <td>top</td>
        <td>格式如：{style:'thin',color:{ rgb: "000000" }}</td>
        <td>object</td>
        <td>style:thin/medium/dotted/dashed</td>
        <td>{style:'thin',color:{ rgb: "000000" }}</td>
    </tr>
    <tr>
        <td>right</td>
        <td>格式如：{style:'thin',color:{ rgb: "000000" }}</td>
        <td>object</td>
        <td>style:thin/medium/dotted/dashed</td>
        <td>{style:'thin',color:{ rgb: "000000" }}</td>
    </tr>
    <tr>
        <td>bottom</td>
        <td>格式如：{style:'thin',color:{ rgb: "000000" }}</td>
        <td>object</td>
        <td>style:thin/medium/dotted/dashed</td>
        <td>{style:'thin',color:{ rgb: "000000" }}</td>
    </tr>
    <tr>
        <td>left</td>
        <td>格式如：{style:'thin',color:{ rgb: "000000" }}</td>
        <td>object</td>
        <td>style:thin/medium/dotted/dashed</td>
        <td>{style:'thin',color:{ rgb: "000000" }}</td>
    </tr>
    <tr>
        <td rowspan="7">font</td>
        <td>name</td>
        <td>字体</td>
        <td>string</td>
        <td>宋体/黑体/Tahoma等</td>
        <td>宋体</td>
    </tr>
    <tr>
        <td>sz</td>
        <td>字号</td>
        <td>number</td>
        <td>--</td>
        <td>12</td>
    </tr>
    <tr>
        <td>color</td>
        <td>字体颜色,格式如：{ rgb: "000000" }</td>
        <td>object</td>
        <td>--</td>
        <td>{ rgb: "000000" }</td>
    </tr>
    <tr>
        <td>bold</td>
        <td>是否为粗体</td>
        <td>boolean</td>
        <td>true/false</td>
        <td>false</td>
    </tr>
    <tr>
        <td>italic</td>
        <td>是否为斜体</td>
        <td>boolean</td>
        <td>true/false</td>
        <td>false</td>
    </tr>
    <tr>
        <td>underline</td>
        <td>是否有下划线</td>
        <td>boolean</td>
        <td>true/false</td>
        <td>false</td>
    </tr>
    <tr>
        <td>shadow</td>
        <td>是否有阴影</td>
        <td>boolean</td>
        <td>true/false</td>
        <td>false</td>
    </tr>
    <tr>
        <td>fill</td>
        <td>fgColor</td>
        <td>背景色</td>
        <td>{ rgb: "ffffff" }</td>
        <td>--</td>
        <td>{ rgb: "ffffff" }</td>
    </tr>
    <tr>
        <td rowspan="3">alignment</td>
        <td>horizontal</td>
        <td>水平对齐方式</td>
        <td>string</td>
        <td>bottom/center/top</td>
        <td>center</td>
    </tr>
    <tr>
        <td>vertical</td>
        <td>垂直对齐方式</td>
        <td>string</td>
        <td>bottom/center/top</td>
        <td>center</td>
    </tr>
    <tr>
        <td>wrapText</td>
        <td>文字是否换行</td>
        <td>boolean</td>
        <td>true/false</td>
        <td>false</td>
    </tr>
</table>

<h5 id="cell-style">单元格样式</h5>

<table>
    <tr>
        <td>参数</td>
        <td>说明</td>
        <td>类型</td>
        <td>可选值</td>
        <td>默认值</td>
    </tr>
    <tr>
        <td>cell</td>
        <td>单元格名称，如A1</td>
        <td>string</td>
        <td>--</td>
        <td>--</td>
    </tr>
</table>

其他属性与[表格全局样式](#global-style)设置方式一致

<div id="export-method"></div>

#### Methods:

方法名|说明|参数
-|-|-
pikaExportExcel|导出函数|--

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
参数|说明|类型|可选值|默认值
-|-|-|-|-
sheetNames|需要查询的表名，如['插件信息']|Array|--|--
before-import|文件导入前的钩子，参数file为导入文件|function(file)|--|--
on-progress|文件导入时的钩子|function(event,file)|--|--
on-change|文件状态改变时的钩子，导入文件、导入成功和导入失败时都会被调用|function(file)|--|--
on-success|文件导入成功的钩子，参数response为生成的json数据|function(response, file)|--|--
on-error|文件导入失败的钩子，参数error为错误信息|function(error, file)|--|--

## Reference
[https://www.jianshu.com/p/31534691ed53](https://www.jianshu.com/p/31534691ed53)

[https://www.cnblogs.com/yinxingen/p/11052184.html](https://www.cnblogs.com/yinxingen/p/11052184.html)