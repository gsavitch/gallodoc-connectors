const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const path = require("path");
const fs = require("fs");
const os = require("os");

module.exports = {
  context: __dirname,
  entry: {
    taskpane: "./src/taskpane/taskpane.ts",
  },
  resolve: {
    extensions: [".ts", ".tsx", ".html", ".js"],
  },
  module: {
    rules: [
      {
        test: /\.ts$/,
        exclude: /node_modules/,
        use: {
          loader: "ts-loader",
        },
      },
      {
        test: /\.html$/,
        exclude: /node_modules/,
        use: "html-loader",
      },
      {
        test: /\.(png|jpg|jpeg|gif|svg)$/,
        type: "asset/resource",
      },
      {
        test: /\.css$/,
        use: ["style-loader", "css-loader"],
      },
    ],
  },
  plugins: [
    new HtmlWebpackPlugin({
      template: "./src/taskpane/taskpane.html",
      filename: "taskpane.html",
      chunks: ["taskpane"],
    }),
    new CopyWebpackPlugin({
      patterns: [
        {
          from: "public",
          to: ".",
          noErrorOnMissing: true,
        },
        {
          from: "manifest.prod.xml",
          to: "manifest.prod.xml",
          noErrorOnMissing: true,
        },
      ],
    }),
  ],
  output: {
    filename: "[name].bundle.js",
    path: path.resolve(__dirname, "dist/word-addin"),
    clean: true,
  },
  devServer: {
    headers: {
      "Access-Control-Allow-Origin": "*",
    },
    static: {
      directory: path.join(__dirname, "public"),
    },
    server: {
      type: "https",
      options: (function () {
        const certPath = path.join(os.homedir(), ".office-addin-dev-certs", "localhost.crt");
        const keyPath = path.join(os.homedir(), ".office-addin-dev-certs", "localhost.key");
        if (fs.existsSync(certPath) && fs.existsSync(keyPath)) {
          return {
            key: fs.readFileSync(keyPath),
            cert: fs.readFileSync(certPath),
          };
        }
        return {};
      })(),
    },
    port: 3000,
    hot: true,
  },
};
