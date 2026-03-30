const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

module.exports = {
  entry: {
    taskpane: "./src/taskpane/taskpane.js",
    commands: "./src/commands/commands.js",
    content: "./src/content/content.js",
    presenter: "./src/presenter/presenter.js",
  },
  output: {
    path: path.resolve(__dirname, "dist"),
    filename: "[name].js",
    clean: true,
  },
  resolve: { extensions: [".js"] },
  module: {
    rules: [
      {
        test: /\.js$/,
        exclude: /node_modules/,
        use: { loader: "babel-loader", options: { presets: ["@babel/preset-env"] } },
      },
      { test: /\.css$/, use: ["style-loader", "css-loader"] },
    ],
  },
  plugins: [
    new HtmlWebpackPlugin({
      filename: "taskpane.html",
      template: "./src/taskpane/taskpane.html",
      chunks: ["taskpane"],
    }),
    new HtmlWebpackPlugin({
      filename: "commands.html",
      template: "./src/commands/commands.html",
      chunks: ["commands"],
    }),
    new HtmlWebpackPlugin({
      filename: "content.html",
      template: "./src/content/content.html",
      chunks: ["content"],
    }),
    new HtmlWebpackPlugin({
      filename: "presenter.html",
      template: "./src/presenter/presenter.html",
      chunks: ["presenter"],
    }),
    new CopyWebpackPlugin({
      patterns: [{ from: "assets", to: "assets" }],
    }),
  ],
  devServer: {
    static: { directory: path.resolve(__dirname, "dist") },
    server: "https",
    port: 3001,
    headers: { "Access-Control-Allow-Origin": "*" },
  },
};
