const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');

module.exports = {
    mode: 'development',
    entry: './src/taskpane/taskpane.js',
    output: {
        path: path.resolve(__dirname, 'dist'),
        filename: 'taskpane.bundle.js',
    },
    resolve: {
        extensions: ['.js'],
    },
    plugins: [
        new HtmlWebpackPlugin({
            template: './src/taskpane/taskpane.html',
            filename: 'taskpane.html'
        }),
    ],
    module: {
        rules: [
            {
                test: /\.js$/,
                exclude: /node_modules/,
                use: {
                    loader: 'babel-loader',
                    options: {
                        presets: ['@babel/preset-env'],
                    },
                },
            },
        ],
    },
    devtool: 'source-map',
    devServer: {
        static: {
            directory: path.join(__dirname, 'dist'),
        },
        compress: true,
        port: 3000,
    },
};