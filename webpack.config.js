/* eslint-disable @typescript-eslint/no-var-requires */
const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const fs = require("fs");

module.exports = (env, options) => {
    const isDev = options.mode !== "production";

    return {
        entry: {
            taskpane: "./src/taskpane/index.tsx",
        },
        output: {
            path: path.resolve(__dirname, "dist"),
            filename: "[name].[contenthash].js",
            clean: true,
            publicPath: isDev ? "/" : "./",
        },
        resolve: {
            extensions: [".ts", ".tsx", ".js", ".jsx"],
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
                    test: /\.(png|svg|jpg|jpeg|gif)$/i,
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
                    { from: "assets", to: "assets", noErrorOnMissing: true },
                    { from: "manifest.xml", to: "manifest.xml" },
                ],
            }),
        ],
        devServer: {
            port: 3000,
            server: "https",
            headers: { "Access-Control-Allow-Origin": "*" },
            client: { overlay: false },
        },
        devtool: isDev ? "source-map" : false,
    };
};
