const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

const urlDev = "https://localhost:3000";

module.exports = (env, argv) => {
  const isDev = argv.mode === "development";

  return {
    entry: {
      taskpane: "./src/taskpane/index.tsx",
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "[name].bundle.js",
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".tsx", ".js", ".jsx"],
      alias: {
        "@core": path.resolve(__dirname, "src/core"),
        "@services": path.resolve(__dirname, "src/services"),
        "@models": path.resolve(__dirname, "src/models"),
        "@utils": path.resolve(__dirname, "src/utils"),
      },
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
        {
          test: /\.(png|jpg|jpeg|gif|svg)$/,
          type: "asset/resource",
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        template: "./src/taskpane/index.html",
        filename: "taskpane.html",
        chunks: ["taskpane"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets",
            to: "assets",
            noErrorOnMissing: true,
          },
          {
            from: "manifest.xml",
            to: "manifest.xml",
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
      },
      port: 3000,
      hot: true,
    },
    devtool: isDev ? "source-map" : false,
  };
};
