<!--
 * @Author: zouzheng
 * @Date: 2020-04-30 11:23:12
 * @LastEditors: Please set LastEditors
 * @LastEditTime: 2025-01-02 14:47:00
 * @Description: 这是XXX组件（页面）
 -->

## 介绍

导入导出 excel 的 js 插件，在 xlsx 和 xlsx-style 的基础上做了简单的封装，开箱即用。

## 特性

- 支持导出 excel 文件，并可设置列宽，边框，字体，字体颜色，字号，对齐方式，背景色等样式。
- 支持 excel 文件导入，生成 json 数据，考虑到客户端机器性能，导入大量数据时，推荐拆分数据分成多个文件导入。

## 版本更新

本插件库已更新至 1.x 版本，历史版本 0.2.x 文档请看[这里](https://github.com/pikaz-18/pikaz-excel-js/blob/master/version/0.2.16-README.md)

- 新版本改为纯 js 库，支持多种框架如 vue2, vue3, react 及无其他依赖的 html 中使用
- 合并项与单元格格式中的单元格名称，现在支持传入数字，而非只能使用 excel 单元格名称，如第一行第三列，可使用 A3 或 3-1
- 新增 esmodule 模块化，支持 vite 等使用

## [demo 示例点击这里体验](https://pikaz-18.github.io/pikaz-excel-js/example/index.html)

## [demo 代码点击这里一键 copy](https://github.com/pikaz-18/pikaz-excel-js/blob/master/example/index.html)

## 安装

### 使用 npm 或 yarn

```bash
yarn add pikaz-excel-js

npm i -S pikaz-excel-js
```

```js
import { excelExport, excelImport } from "pikaz-excel-js";
```

### 使用 cdn 引入

```html
<script
  type="text/javascript"
  src="https://cdn.jsdelivr.net/npm/pikaz-excel-js"
></script>
或者
<script type="text/javascript" src="https://unpkg.com/pikaz-excel-js"></script>
```

```js
const { excelExport, excelImport } = window.pikazExcelJs;
```

### 导出函数

#### 函数示例

```js
import { excelExport } from "pikaz-excel-js";
excelExport({
  sheet: [
    {
      // 表格标题
      title: "水果的味道1",
      // 表头
      tHeader: ["种类", "味道"],
      // 数据键名
      keys: ["name", "taste"],
      // 表格数据
      table: [
        {
          name: "荔枝",
          taste: "甜",
        },
        {
          name: "菠萝蜜",
          taste: "甜",
        },
      ],
      sheetName: "水果的味道1",
    },
  ],
});
```

#### 函数参数:

| 参数         | 说明                                                                                | 类型                                | 可选值 | 默认值 |
| ------------ | ----------------------------------------------------------------------------------- | ----------------------------------- | ------ | ------ |
| bookType     | 文件格式                                                                            | string                              | xlsx   | xlsx   |
| filename     | 文件名称                                                                            | string                              | --     | excel  |
| sheet        | 表格数据，每个表格数据对象配置具体看下方[表格配置](#table-setting)                  | object[]                            | --     | --     |
| beforeStart  | 处理数据之前的钩子，参数为导出的文件格式，文件名，表格数据，若抛出 Error 则停止导出 | function(bookType, filename, sheet) | --     | --     |
| beforeExport | 导出文件之前的钩子，参数为 blob 文件流，文件格式，文件名，若抛出 Error 则停止导出   | function(blob, bookType, filename)  | --     | --     |
| onError      | 导出失败的钩子，参数为错误信息                                                      | function(err)                       | --     | --     |

<h5 id="table-setting">表格参数配置</h5>

| 参数        | 说明                                                                                                                                                                     | 类型     | 可选值 | 默认值                        |
| ----------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------ | -------- | ------ | ----------------------------- |
| title       | 表格标题，自动设置合并，非必须项                                                                                                                                         | string   | --     | --                            |
| tHeader     | 表头, 非必须项                                                                                                                                                           | string[] | --     | --                            |
| table       | 表格数据，如果无数据，设置为空字符""，避免使用 null 或 undefined                                                                                                         | object[] | --     | --                            |
| merges      | 合并两个单元格之间所有的单位格，支持 excel 行列格式或数字格式（如合并第一排第一列至第一排第三列为'A1: A3'或'1-1:3-1'），合并的表格单元多余数据项以空字符串填充，非必须项 | string[] | --     | --                            |
| keys        | 数据键名，需与表头内容顺序对应                                                                                                                                           | string[] | --     | --                            |
| colWidth    | 列宽，若不传，则列宽自适应（自动列宽时数据类型必须为 string，如有其他数据类型，请手动设置列宽）                                                                          | number[] | --     | --                            |
| sheetName   | 表格名称                                                                                                                                                                 | string   | --     | sheet                         |
| globalStyle | 表格全局样式，具体参数查看下方[表格全局样式](#global-style)                                                                                                              | object   | --     | [表格全局样式](#global-style) |
| cellStyle   | 单元格样式，每个单元格对象配置具体参数查看下方[单元格样式](#cell-style)                                                                                                  | object[] | --     | --                            |

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
        <td>left/center/right</td>
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
        <td>单元格名称，支持excel单元格名称与数字行列格式，如'A3'或'3-1'</td>
        <td>string</td>
        <td>--</td>
        <td>--</td>
    </tr>
</table>

其他属性与[表格全局样式](#global-style)设置方式一致

<div id="export-method"></div>

### 导入函数

#### 函数示例

```js
import { excelImport } from "pikaz-excel-js";
excelImport().then((res) => {
  console.log(res);
});
```

#### 函数参数:

| 参数              | 说明                                                                                                           | 类型                     | 可选值     | 默认值 |
| ----------------- | -------------------------------------------------------------------------------------------------------------- | ------------------------ | ---------- | ------ |
| file              | 导入的文件，若不传，则自动调起上传功能                                                                         | file                     | --         | null   |
| sheetNames        | 需要导入表的表名，如['插件信息']，若不传则读取 excel 中所有表格，非必传                                        | string[]                 | --         | --     |
| removeBlankspace  | 是否移除数据中字符串的空格                                                                                     | Boolean                  | true/false | false  |
| removeSpecialchar | 是否移除不同版本及环境下 excel 数据中出现的特殊不可见字符，如 u202D 等, 使用此功能，返回的数据将被转化为字符串 | Boolean                  | true/false | true   |
| beforeImport      | 文件导入前的钩子，参数 file 为导入文件                                                                         | function(file)           | --         | --     |
| onProgress        | 文件导入时的钩子                                                                                               | function(event, file)    | --         | --     |
| onChange          | 文件状态改变时的钩子，导入文件、导入成功和导入失败时都会被调用                                                 | function(file)           | --         | --     |
| onSuccess         | 文件导入成功的钩子，参数 response 为生成的 json 数据                                                           | function(response, file) | --         | --     |
| onError           | 文件导入失败的钩子，参数 error 为错误信息                                                                      | function(error)          | --         | --     |
