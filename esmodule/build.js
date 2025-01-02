const path = require("path");
const fs = require("fs");
const createModule = require("@pikaz/transform-package-to-ESModule");

createModule({
  jsPath: path.join(__dirname, "excelExport.js"),
  outPath: path.join(__dirname, "lib", "moduleExcelExport.js"),
  keys: ["excelExport"],
});

createModule({
  jsPath: path.join(__dirname, "excelImport.js"),
  outPath: path.join(__dirname, "lib", "moduleExcelImport.js"),
  keys: ["excelImport"],
});

fs.cp(
  path.join(__dirname, "lib"),
  path.join(__dirname, "../lib"),
  { recursive: true },
  (err) => {
    if (err) {
      console.log(err);
    }
  }
);
