const path = require("path");

module.exports = {
  mode: "production",
  entry: {
    excelExport: path.resolve(__dirname, "src", "excelExport", "index.js"),
    excelImport: path.resolve(__dirname, "src", "excelImport", "index.js"),
  },
  output: {
    filename: "[name].js",
    path: path.resolve(__dirname, "esmodule"),
    libraryTarget: "window",
    library: "[name]",
    libraryExport: "default",
  },
};
