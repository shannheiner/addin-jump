const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');

module.exports = {
  entry: {
    taskpane: './taskpane.js',
  },
  output: {
    path: path.resolve(__dirname, 'dist'),
    filename: '[name].js',
    clean: true
  },
  resolve: {
    extensions: ['.ts', '.js']
  },
  module: {
    rules: [
      {
        test: /\.ts$/,
        use: 'ts-loader',
        exclude: /node_modules/
      },
      {
        test: /\.js$/,
        use: 'source-map-loader',
        enforce: 'pre'
      }
    ]
  },
  plugins: [
    new HtmlWebpackPlugin({
      filename: 'taskpane.html',
      template: './taskpane.html',
      chunks: ['taskpane']
    }),
    new HtmlWebpackPlugin({
      filename: 'function-file.html',
      template: './function-file.html',
      chunks: []
    }),
    new CopyWebpackPlugin({
      patterns: [
        {
          from: 'manifest.xml',
          to: 'manifest.xml'
        },
        {
          from: 'taskpane.css',
          to: 'taskpane.css'
        }
      ]
    })
  ],
  devServer: {
    static: {
      directory: path.join(__dirname, 'dist')
    },
    headers: {
      'Access-Control-Allow-Origin': '*'
    },
    server: {
      type: 'https',
      options: {
        key: path.resolve(__dirname, 'certs/server.key'),
        cert: path.resolve(__dirname, 'certs/server.crt'),
        ca: path.resolve(__dirname, 'certs/ca.crt')
      }
    },
    port: 3000,
    hot: true
  }
};