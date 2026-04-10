const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

// Try to load dev certs for trusted HTTPS (needed for Office Add-in sideloading)
let devCerts;
try {
  devCerts = require("office-addin-dev-certs");
} catch (e) {
  // Not installed — fall back to default self-signed
}

module.exports = async (env, argv) => {
  let httpsOptions = true; // default: webpack self-signed cert

  // Use office-addin-dev-certs for trusted HTTPS in development
  if (devCerts && argv.mode === "development") {
    try {
      httpsOptions = await devCerts.getHttpsServerOptions();
    } catch (e) {
      console.log("Could not load dev certs, using default HTTPS:", e.message);
    }
  }

  return {
    entry: {
      taskpane: "./src/index.js",
      // Separate bundle for the hidden function-command runtime loaded by
      // commands.html. Keeps it isolated from the task pane so Copilot can
      // invoke commands without the full UI bundle evaluating.
      commands: "./src/copilot/commandsEntry.js",
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "[name].bundle.js",
      clean: true,
    },
    module: {
      rules: [
        {
          test: /\.js$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
            options: {
              presets: ["@babel/preset-env"],
            },
          },
        },
        {
          test: /\.css$/,
          use: ["style-loader", "css-loader"],
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        template: "./src/ui/taskpane.html",
        filename: "taskpane.html",
        chunks: ["taskpane"],
      }),
      new HtmlWebpackPlugin({
        template: "./src/copilot/commands.html",
        filename: "commands.html",
        chunks: ["commands"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: "assets", to: "assets" },
          { from: "manifest.xml", to: "manifest.xml" },
          { from: "copilot-package", to: "copilot-package" },
        ],
      }),
    ],
    devServer: {
      static: {
        directory: path.join(__dirname, "dist"),
      },
      port: 3000,
      server: {
        type: "https",
        options: httpsOptions === true ? {} : httpsOptions,
      },
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      allowedHosts: "all",
    },
    resolve: {
      extensions: [".js"],
      fallback: {
        fs: false,
        stream: false,
        crypto: false,
        path: false,
      },
    },
  };
};
