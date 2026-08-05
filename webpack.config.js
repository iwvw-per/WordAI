const path = require("path");
const fs = require("fs");
const os = require("os");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const devCerts = require("office-addin-dev-certs");

// 本地调试证书：优先使用 ~/.office-addin-dev-certs 下的有效证书
// （office-addin-dev-certs 的 getHttpsServerOptions() 在部分环境会因证书重装流程挂起）
function loadLocalHttpsOptions() {
  const certDir = path.join(os.homedir(), ".office-addin-dev-certs");
  const certFile = path.join(certDir, "localhost.crt");
  const keyFile = path.join(certDir, "localhost.key");
  const caFile = path.join(certDir, "ca.crt");
  if (fs.existsSync(certFile) && fs.existsSync(keyFile)) {
    const options = {
      cert: fs.readFileSync(certFile),
      key: fs.readFileSync(keyFile),
    };
    if (fs.existsSync(caFile)) {
      options.ca = fs.readFileSync(caFile);
    }
    return options;
  }
  return undefined;
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const httpsOptions = dev ? (loadLocalHttpsOptions() || await devCerts.getHttpsServerOptions()) : {};

  return {
    entry: {
      taskpane: "./src/taskpane/taskpane.js",
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      filename: dev ? "[name].bundle.js" : "[name].[contenthash].js",
      clean: true,
    },
    resolve: {
      extensions: [".js"],
    },
    module: {
      rules: [
        {
          test: /\.css$/,
          use: ["style-loader", "css-loader"],
        },
        {
          test: /\.(png|jpg|jpeg|gif|svg)$/,
          type: "asset/resource",
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
            from: "src/assets",
            to: "assets",
            noErrorOnMissing: true,
          },
          {
            from: "src/assets/favicon.png",
            to: "favicon.png",
            noErrorOnMissing: true,
          },
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
        options: httpsOptions,
      },
      port: 3000,
      hot: true,
    },
  };
};
