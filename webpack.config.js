/*
 * @Description: 这是***页面（组件）
 * @Date: 2022-08-14 18:47:21
 * @Author: zouzheng
 * @LastEditors: zouzheng
 * @LastEditTime: 2022-08-14 18:50:24
 */
const path = require('path');

module.exports = {
    mode: 'production',
    entry: path.resolve(__dirname, 'src', 'index.js'),
    output: {
        publicPath: path.resolve(__dirname, 'lib'),
        filename: 'pikazExcel.js',
        path: path.resolve(__dirname, 'lib'),
        libraryTarget: 'umd',
        library: "pikazExcelJs"
    }
};