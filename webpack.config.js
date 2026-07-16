const path = require("path");
const os = require("os");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

module.exports = {
  entry: {
    taskpane: "./src/taskpane/index.tsx",
    commands: "./src/commands/commands.ts",
  },
  resolve: {
    extensions: [".tsx", ".ts", ".js", ".html"],
    fullySpecified: false,
  },
  module: {
    rules: [
      {
        test: /\.tsx?$/,
        use: "ts-loader",
        exclude: /node_modules/,
      },
      {
        test: /\.css$/,
        use: ["style-loader", "css-loader"],
      },
    ],
  },
  optimization: {
    usedExports: false,
    providedExports: false,
  },
  plugins: [
    new HtmlWebpackPlugin({
      template: "./src/taskpane/taskpane.html",
      filename: "taskpane/taskpane.html",
      chunks: ["taskpane"],
    }),
    new CopyWebpackPlugin({
      patterns: [
        { from: "assets", to: "assets", noErrorOnMissing: true },
        { from: "knowledge", to: "knowledge", noErrorOnMissing: true },
      ],
    }),
  ],
  devServer: {
    static: {
      directory: path.join(__dirname, "dist"),
    },
    headers: {
      "Access-Control-Allow-Origin": "*",
    },
    server: {
      type: "https",
      options: {
        key: path.join(
          os.homedir(),
          ".office-addin-dev-certs",
          "localhost.key",
        ),
        cert: path.join(
          os.homedir(),
          ".office-addin-dev-certs",
          "localhost.crt",
        ),
      },
    },
    port: 3000,
    hot: true,
  },
  output: {
    path: path.resolve(__dirname, "dist"),
    filename: "[name]/[name].js",
    clean: true,
  },
};
