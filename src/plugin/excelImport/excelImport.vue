<!--
 * @Author: zouzheng
 * @Date: 2020-04-30 15:05:31
 * @LastEditors: zouzheng
 * @LastEditTime: 2020-10-28 00:47:07
 * @Description: 这是excel导入组件（页面）
 -->
<template>
    <div
        class="excel-import"
        @click="importFileClick"
    >
        <input
            type="file"
            @change="importFile(this)"
            :id="id"
            style="display: none"
            accept=".xls,.xlsx"
        />
        <slot></slot>
    </div>
</template>

<script>
import XLSX from "xlsx";
export default {
    props: {
        // 表名
        sheetNames: Array,
        // 是否移除空格
        removeBlankspace: {
            type: Boolean,
            default: false
        },
        // 是否移出特殊字符
        removeSpecialchar: {
            type: Boolean,
            default: true
        },
        //  导入前
        beforeImport: {
            type: Function,
            // file为导入文件
            default: file => {}
        },
        // 导入时
        onProgress: {
            type: Function,
            default: (event, file) => {}
        },
        // 状态改变
        onChange: {
            type: Function,
            default: file => {}
        },
        onSuccess: {
            type: Function,
            // response为生成的json数据
            default: (response, file) => {}
        },
        onError: {
            type: Function,
            // err为错误信息
            default: (err, file) => {}
        }
    },
    components: {},
    data() {
        return {
            imFile: "",
            // 枚举类
            enum: {
                // 文件类型
                bookType: ["xlsx", "xls"]
            },
            id: ""
        };
    },
    created() {
        this.id = Math.random().toString();
    },
    mounted() {
        this.imFile = document.getElementById(this.id);
    },
    methods: {
        /**
         * @name: 点击导入按钮
         * @param {type}
         * @return:
         */
        importFileClick() {
            this.imFile.click();
        },
        /**
         * @name: 导入文件
         * @param {type}
         * @return:
         */
        importFile() {
            // 导入excel
            const obj = this.imFile;
            // 无导入文件
            if (!obj.files) {
                this.onError("No imported file");
                return;
            }
            const file = obj.files[0];
            // 导入前
            const beforeImport = this.beforeImport(file);
            this.onChange(file);
            if (beforeImport === false) {
                return;
            }
            // 文件类型必须为xlsx或者xls
            const bookType = file.name.substr(
                file.name.length - 4,
                file.name.length - 1
            );
            const type = this.enum.bookType.some(e => {
                if (bookType.indexOf(e) !== -1) {
                    return true;
                }
                return false;
            });
            if (!type) {
                this.onError('The file type must be "xlsx" or "xls"', file);
                return;
            }
            const reader = new FileReader();
            const that = this;
            // 导入时
            reader.onprogress = e => {
                this.onProgress(e, file);
            };
            // 导入完成
            reader.onload = e => {
                const data = e.target.result;
                if (that.rABS) {
                    that.wb = XLSX.read(btoa(this.fixdata(data)), {
                        // 手动转化
                        type: "base64"
                    });
                } else {
                    that.wb = XLSX.read(data, {
                        type: "binary"
                    });
                }
                let json = [];
                // 查询对应表名数据
                if (that.sheetNames) {
                    that.sheetNames.forEach(name => {
                        const sheetIndex = that.wb.SheetNames.findIndex(
                            s => s === name
                        );
                        if (sheetIndex !== -1) {
                            const data = XLSX.utils.sheet_to_json(
                                that.wb.Sheets[that.wb.SheetNames[sheetIndex]]
                            );
                            json.push({ sheetName: name, data });
                        }
                    });
                } else {
                    // 查询全部数据
                    that.wb.SheetNames.forEach(item => {
                        const data = XLSX.utils.sheet_to_json(
                            that.wb.Sheets[item]
                        );
                        json.push({ sheetName: item, data });
                    });
                }
                json = this.dealData(json);
                that.importData(json, file);
            };
            if (this.rABS) {
                reader.readAsArrayBuffer(file);
            } else {
                reader.readAsBinaryString(file);
            }
        },
        /**
         * @name: 处理导入数据
         * @param {type}
         * @return:
         */
        dealData(data) {
            if (this.removeBlankspace || this.removeSpecialchar) {
                const json = data.map(item => {
                    const itemData = item.data.map(i => {
                        Object.keys(i).forEach(key => {
                            if (
                                this.removeBlankspace &&
                                Object.prototype.toString.call(i[key]) ===
                                    "[object String]"
                            ) {
                                // 字符串去除空格
                                i[key] = i[key].replace(/\s*/g, "");
                            }
                            // 去除特殊字符
                            if (
                                this.removeSpecialchar &&
                                i[key] &&
                                Object.prototype.toString.call(i[key]) !==
                                    "[object Boolean]"
                            ) {
                                i[key] = i[key]
                                    .toString()
                                    .replace(
                                        /[\u200b-\u200f\uFEFF\u202a-\u202e]/g,
                                        ""
                                    );
                            }
                        });
                        return i;
                    });
                    return { ...item, data: itemData };
                });
                return json;
            }
            return data;
        },
        /**
         * @name: 导入数据
         * @param {type}
         * @return:
         */
        importData(data, file) {
            this.imFile.value = "";
            if (data.length <= 0) {
                // 导入失败
                this.onChange(file);
                this.onError("The import failed", file);
                return;
            } else {
                //导入成功
                this.onChange(file);
                this.onSuccess(data, file);
                return;
            }
        },
        /**
         * @name: 文件流转BinaryString
         * @param {type}
         * @return:
         */
        fixdata(data) {
            const o = "";
            const l = 0;
            const w = 10240;
            for (; l < data.byteLength / w; ++l) {
                o += String.fromCharCode.apply(
                    null,
                    new Uint8Array(data.slice(l * w, l * w + w))
                );
            }
            o += String.fromCharCode.apply(
                null,
                new Uint8Array(data.slice(l * w))
            );
            return o;
        }
    },
    computed: {},
    watch: {}
};
</script>

<style scoped>
</style>