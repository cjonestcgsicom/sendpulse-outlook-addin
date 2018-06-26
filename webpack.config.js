const HtmlWebpackPlugin = require('html-webpack-plugin');
const sendpulse = require('./modules/sendpulse_api_custom.js');
module.exports = {
    entry: {
        app: './src/index.ts',
        'function-file': './function-file/function-file.ts',
        'greenrope_api': './src/sendpulse_api.ts'
    },
    resolve: {
        extensions: ['.ts', '.tsx', '.html', '.js']
    },

    devServer: {
        setup: function(app) {

            var bodyParser = require('body-parser');
            app.use(bodyParser.json());

            app.post('/token', bodyParser.json(), function(req, res) {
                console.log(req.body);

                var data = req.body;
                sendpulse.init(data.client_id, data.client_secret, '/tmp/');
                sendpulse.getToken((result) => {
                    res.send(result);
                });

            });

        }
    },

    module: {
        rules: [
            {
                test: /\.tsx?$/,
                exclude: /node_modules/,
                use: 'ts-loader'
            },
            {
                test: /\.html$/,
                exclude: /node_modules/,
                use: 'html-loader'
            },
            {
                test: /\.(png|jpg|jpeg|gif)$/,
                use: 'file-loader'
            }
        ]
    },
    plugins: [
        new HtmlWebpackPlugin({
            template: './index.html',
            chunks: ['app']
        }),
        new HtmlWebpackPlugin({
            template: './function-file/function-file.html',
            filename: 'function-file/function-file.html',
            chunks: ['function-file']
        })
    ]
};