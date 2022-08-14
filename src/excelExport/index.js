/*
 * @Description: 这是***页面（组件）
 * @Date: 2022-07-16 16:07:35
 * @Author: zouzheng
 * @LastEditors: zouzheng
 * @LastEditTime: 2022-08-15 01:10:08
 */
import { createLetter } from "../common/index";
import { saveAs } from 'file-saver'
import XLSX from 'pikaz-xlsx-style'

// 默认配置
const config = {
    // 文件类型
    bookType: "xlsx",
    // 文件名
    filename: "excel",
    // 表格配置
    sheet: [],
    // 处理数据前// bookType:文件类型,filename:文件名,sheet:表格数据
    beforeStart: (bookType, filename, sheet) => { },
    // 导出前// filename:文件名,sheet:表格数据,blob:文件流
    beforeExport: (bookType, filename, blob) => { },
    // 导出错误
    onError: (err) => { },
}

// 默认全局样式
const defaultGlobalStyle = {
    border: {
        top: {
            style: "thin",
            color: { rgb: "000000" },
        },
        bottom: {
            style: "thin",
            color: { rgb: "000000" },
        },
        left: {
            style: "thin",
            color: { rgb: "000000" },
        },
        right: {
            style: "thin",
            color: { rgb: "000000" },
        },
    },
    font: {
        name: "宋体",
        sz: 12,
        color: { rgb: "000000" },
        bold: false,
        italic: false,
        underline: false,
        shadow: false,
    },
    alignment: {
        horizontal: "center",
        vertical: "center",
        wrapText: false,
    },
    fill: {
        fgColor: { rgb: "ffffff" },
    }
}

// 默认表格配置
const defaultSheet = {
    // 标题
    title: "",
    // 多级表头
    tHeader: [],
    // 表格数据
    table: [],
    // 合并项
    merges: [],
    // 数据键值
    keys: [],
    // 列宽
    colWidth: "",
    // 表名
    sheetName: "excel",
    // 全局样式
    globalStyle: {},
    // 单元格样式
    cellStyle: [],
}

// 全字母
const letterArr = createLetter()

/**
     * @name:导出excel
     * @param {type}
     * @return:
     */
const exportExcel = async (obj = {}) => {
    const { bookType, filename, sheet, beforeStart, beforeExport, onError } = { ...config, ...obj }
    // 处理数据前
    await beforeStart(
        bookType,
        filename,
        sheet
    );
    if (!sheet.length) {
        onError("Table data cannot be empty");
        return;
    }
    // excel表内容
    const wb = { SheetNames: [], Sheets: [] };
    sheet.forEach((item, index) => {
        let {
            // 标题
            title,
            // 多级表头
            tHeader,
            // 表格数据
            table,
            // 合并项
            merges,
            // 数据键值
            keys,
            // 列宽
            colWidth,
            // 表名
            sheetName,
            // 全局样式
            globalStyle,
            // 单元格样式
            cellStyle,
        } = { ...JSON.parse(JSON.stringify(defaultSheet)), ...item };
        // 全局样式
        const dgStyle = { ...JSON.parse(JSON.stringify(defaultGlobalStyle)), ...globalStyle };
        // 处理合并项
        merges = merges.map(merge => {
            let result = merge
            const mergeArr = merge.replace(/-/g, "").split(":")
            // 若为纯数字，则转换为excel格式
            if (mergeArr.every(m => /^\d+$/.test(m))) {
                result = numToLetter(merge)
            }
            return result
        })
        console.log(merges);
        let titleArr = []
        // 处理标题格式
        if (title) {
            // 取表头中的最大值
            const tHeaderLength = tHeader && tHeader.length || 0
            const titleLength = Math.max(
                tHeaderLength,
                keys.length
            );
            // 第一个元素为title，剩余以空字符串填充
            titleArr = [title, ...Array(titleLength - 1).fill("")]
            // 处理标题的合并
            const titleMerge = numToLetter(`1-1:${titleLength}-1`);
            if (!merges) {
                merges = [titleMerge];
            } else {
                if (merges.indexOf(titleMerge) === -1) {
                    merges.push(titleMerge);
                }
            }
        }
        // console.log(merges);
        //表头对应字段
        const data = table.map((v) => keys.map((j) => v[j]));
        tHeader && data.unshift(tHeader);
        title && data.unshift(titleArr);
        const ws = sheet_from_array_of_arrays(data);
        if (merges && merges.length > 0) {
            if (!ws["!merges"]) ws["!merges"] = [];
            merges.forEach((merge) => {
                ws["!merges"].push(XLSX.utils.decode_range(merge));
            });
        }
        // 如果没有列宽则自适应
        if (!colWidth) {
            // 基准比例，以12为标准
            const benchmarkRate = dgStyle.font.sz / 12;
            // 空字符长度
            const nullstr = 10 * benchmarkRate + 2;
            // 单个中文字符长度
            const chinese = 2 * benchmarkRate;
            // 单个非中文字符长度
            const nChinese = benchmarkRate;
            //设置worksheet每列的最大宽度,并+2调整一点列宽
            const sheetColWidth = data.map((row) =>
                row.map((val) => {
                    //先判断是否为null/undefined
                    if (!val) {
                        return {
                            wch: nullstr,
                        };
                    } else {
                        const strArr = val.toString().split("");
                        const pattern = new RegExp("[\u4E00-\u9FA5]+");
                        let re = strArr.map((str) => {
                            // 是否为中文
                            if (pattern.test(str)) {
                                return chinese;
                            } else {
                                return nChinese;
                            }
                        });
                        re = re.reduce((total, r) => total + r, 0);
                        return {
                            wch: re + 2,
                        };
                    }
                })
            );
            /*以第一行为初始值*/
            let result = sheetColWidth[0];
            for (let i = 1; i < sheetColWidth.length; i++) {
                for (let j = 0; j < sheetColWidth[i].length; j++) {
                    if (result[j]["wch"] < sheetColWidth[i][j]["wch"]) {
                        result[j]["wch"] = sheetColWidth[i][j]["wch"];
                    }
                }
            }
            ws["!cols"] = result;
        } else {
            ws["!cols"] = colWidth.map((i) => {
                return { wch: i };
            });
        }
        // 添加工作表
        wb.SheetNames.push(sheetName);
        wb.Sheets[sheetName] = ws;
        const dataInfo = wb.Sheets[wb.SheetNames[index]];
        //全局样式
        (function () {
            Object.keys(dataInfo).forEach((i) => {
                if (i == "!ref" || i == "!merges" || i == "!cols") {
                } else {
                    dataInfo[i.toString()].s = dgStyle;
                }
            });
        })();
        // 单个样式
        (function () {
            if (!cellStyle || cellStyle.length <= 0) {
                return;
            }
            cellStyle.forEach((s) => {
                const { border, font, alignment, fill } = s;
                dataInfo[s.cell].s = {
                    border: border === {} ? border : border || dgStyle.border,
                    font: font || dgStyle.font,
                    alignment: alignment || dgStyle.alignment,
                    fill: fill || dgStyle.fill,
                };
            });
        })();
    });
    // 类型默认为xlsx
    await writeExcel({ wb, bookType, filename, beforeExport });
}

/**
 * @description: 纯数字转换为excel格式行列合并项
 * @param {*} num
 * @return {*}
 */
const numToLetter = (merge = "1-1:2-1") => {
    const mergeArr = merge.split(":").map(item => {
        const [row, column] = item.split("-").map(s => Number(s))
        const stringArray = [];
        const numToLetters = (num) => {
            const result = num - 1;
            const first = parseInt(result / 26);
            const second = result % 26;
            stringArray.push(String.fromCharCode(64 + parseInt(second + 1)));
            if (first > 0) {
                numToLetters(first);
            }
        }
        numToLetters(row);
        return stringArray.reverse().join("") + column;
    })
    return mergeArr.join(":")
}

/**
 * @name: 导出excel文件
 * @param {type}
 * @return:
 */
const writeExcel = async ({ wb, bookType, filename, beforeExport }) => {
    const wbout = XLSX.write(wb, {
        bookType: bookType,
        bookSST: false,
        type: "binary",
    });
    const blob = new Blob([s2ab(wbout)], {
        type: "application/octet-stream",
    });
    await beforeExport(blob, bookType, filename);
    saveAs(blob, `${filename}.${bookType}`);
}

/**
 * @name: 转化时间格式
 * @param {type}
 * @return:
 */
const datenum = (v, date1904) => {
    if (date1904) v += 1462;
    const epoch = Date.parse(v);
    return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
}

/**
 * @name: 设置数据类型
 * @param {type}
 * @return:
 */
const sheet_from_array_of_arrays = (data) => {
    let ws = {};
    const range = {
        s: {
            c: 1000000000,
            r: 1000000000,
        },
        e: {
            c: 0,
            r: 0,
        },
    };
    for (let R = 0; R != data.length; ++R) {
        for (let C = 0; C != data[R].length; ++C) {
            if (range.s.r > R) range.s.r = R;
            if (range.s.c > C) range.s.c = C;
            if (range.e.r < R) range.e.r = R;
            if (range.e.c < C) range.e.c = C;
            let cell = {
                v: data[R][C],
            };
            if (cell.v == null) continue;
            let cell_ref = XLSX.utils.encode_cell({
                c: C,
                r: R,
            });
            if (typeof cell.v === "number") cell.t = "n";
            else if (typeof cell.v === "boolean") cell.t = "b";
            else if (cell.v instanceof Date) {
                cell.t = "n";
                cell.z = XLSX.SSF._table[14];
                cell.v = datenum(cell.v);
            } else cell.t = "s";
            ws[cell_ref] = cell;
        }
    }
    if (range.s.c < 1000000000) ws["!ref"] = XLSX.utils.encode_range(range);
    return ws;
}

/**
 * @name: 转换格式
 * @param {type}
 * @return:
 */
const s2ab = (s) => {
    const b = new ArrayBuffer(s.length);
    const v = new Uint8Array(b);
    for (let i = 0; i < s.length; i++) {
        v[i] = s.charCodeAt(i) & 0xff;
    }
    return b;
}

export default exportExcel