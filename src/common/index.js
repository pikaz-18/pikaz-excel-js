/*
 * @Description: 这是公共函数页面（组件）
 * @Date: 2022-07-30 15:21:55
 * @Author: zouzheng
 * @LastEditors: zouzheng
 * @LastEditTime: 2022-08-20 23:01:07
 */
const fileType = [
    { type: "xls", val: "d0cf11e0" },
    { type: "xlsx", val: "504b0304" }
]
/**
 * @description: 识别文件类型
 * @param {*} file
 * @return {*}
 */
export const checkFileType = (file) => {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => {
            const buffers = reader.result
            const uint8Array = new Uint8Array(buffers);
            const result = []
            for (let index = 0; index < uint8Array.length; index++) {
                const n = uint8Array[index].toString(16)
                result.push("00".substring(n.length) + n);
            }
            const type = fileType.find(item => item.val === result.join("").toLowerCase())
            if (type) {
                resolve(type.type);
            }
            reject("not xls or xlsx")
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(file.slice(0, 4));
    });
}