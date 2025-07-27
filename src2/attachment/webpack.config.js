// webpack.config.js
const path = require("path");

module.exports = {
  target: "node18", // or node16 if needed
  mode: "production",
  entry: "./src/main.js",
  output: {
    filename: "__main__.js",
    path: path.resolve(__dirname, "dist"),
    libraryTarget: "commonjs2",
  },
  externals: {
    // Avoid bundling built-in node modules like fs, path, etc.
  },
  module: {
    rules: [],
  },
  resolve: {
    extensions: [".js"],
  },
};
