/*
 * @Description: 这是excel导入页面（组件）
 * @Date: 2022-07-16 16:07:41
 * @Author: zouzheng
 * @LastEditors: zouzheng
 * @LastEditTime: 2022-08-07 18:36:29
 */
import XLSX from "xlsx";
import { checkFileType } from "../common/index";

const config = {
    // excel文件
    file: null,
    // 表名
    sheetNames: [],
    // 是否移除空格
    removeBlankspace: false,
    // 是否移出特殊字符
    removeSpecialchar: true,
    //  导入前
    beforeImport: file => { },
    // 导入时
    onProgress: (event, file) => { },
    // 状态改变
    onChange: file => { },
    // 成功
    onSuccess: (response, file) => { },
    // 失败
    onError: (err) => { }
}

/**
 * @description: 获取文件
 * @return {*}
 */
const createUpload = () => {
    return new Promise((resolve, reject) => {
        const input = document.createElement("input")
        input.type = "file"
        input.style.display = "none"
        input.accept = ".xls,.xlsx"
        input.onchange = async (e) => {
            if (e.target.files.length) {
                const file = e.target.files[0]
                resolve(file)
            }
            reject("cancel")
        }
        input.click()
    })
}

/**
 * @name: 处理导入数据
 * @param {type}
 * @return:
 */
const dealData = ({ data, removeBlankspace, removeSpecialchar }) => {
    if (removeBlankspace || removeSpecialchar) {
        const json = data.map(item => {
            const itemData = item.data.map(i => {
                Object.keys(i).forEach(key => {
                    if (
                        removeBlankspace &&
                        Object.prototype.toString.call(i[key]) ===
                        "[object String]"
                    ) {
                        // 字符串去除空格
                        i[key] = i[key].replace(/\s*/g, "");
                    }
                    // 去除特殊字符
                    if (
                        removeSpecialchar &&
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
}

/**
 * @description: 导入处理
 * @return {*}
 */
const fileImport = ({ file, onProgress, sheetNames, onChange, onSuccess, removeBlankspace, removeSpecialchar }) => {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        // 导入时
        reader.onprogress = e => {
            onProgress(e, file);
        };
        // 导入失败
        reader.onerror = reject;
        // 导入完成
        reader.onload = e => {
            const data = e.target.result;
            const wb = XLSX.read(data, {
                type: "binary"
            });
            const json = [];
            // 查询对应表名数据
            if (sheetNames.length) {
                sheetNames.forEach(name => {
                    const sheetIndex = wb.SheetNames.findIndex(
                        s => s === name
                    );
                    if (sheetIndex !== -1) {
                        const data = XLSX.utils.sheet_to_json(
                            wb.Sheets[wb.SheetNames[sheetIndex]]
                        );
                        json.push({ sheetName: name, data });
                    }
                });
            } else {
                // 查询全部数据
                wb.SheetNames.forEach(item => {
                    const data = XLSX.utils.sheet_to_json(
                        wb.Sheets[item]
                    );
                    json.push({ sheetName: item, data });
                });
            }
            const result = dealData({ data: json, removeBlankspace, removeSpecialchar });
            if (result.length <= 0) {
                // 导入失败
                onChange(file);
                reject("The import failed")
            } else {
                //导入成功
                onChange(file);
                onSuccess(result, file)
                resolve(result)
            }
        };
        reader.readAsBinaryString(file);
    })
}


/**
 * @description: 导入文件
 * @param {*} obj/入参
 * @return {*}
 */
const excelImport = async (obj = {}) => {
    const { file: excel, sheetNames, removeBlankspace, removeSpecialchar, beforeImport, onProgress, onChange, onSuccess, onError } = { ...config, ...obj }
    try {
        let file = excel
        // 未传入file则调起上传
        if (!file) {
            file = await createUpload()
        }
        // 文件类型必须未xls/xlsx
        await checkFileType(file)
        // 导入前
        await beforeImport(file)
        await onChange(file);
        const result = await fileImport({ file, onProgress, sheetNames, onChange, onSuccess, removeBlankspace, removeSpecialchar })
        return result
    } catch (error) {
        onError(error)
        throw error
    }
}

export default excelImport