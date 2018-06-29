const HtmlWebpackPlugin = require('html-webpack-plugin');
const sendpulse = require('./modules/sendpulse_api_custom.js');
var postman_request = require('postman-request');
module.exports = {
    entry: {
        app: './src/index.ts',
        'function-file': './function-file/function-file.ts',
        'greenrope_api': './src/sendpulse_api.ts',
        'synchdialog' : './src/synchdialog.ts'
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
                var url = 'https://login.microsoftonline.com/common/oauth2/v2.0/token';
                postman_request.post({url: url, form: req.body}, function optionalCallback(err, httpResponse, body) {
                    if (err) {
                        res.send(err);
                        console.error('failed:', err);
                    }
                    console.log('Call successful!  Server responded with:', JSON.stringify(body));
                    res.send(body);
                });
            });

            app.post('/sendpulsetoken', bodyParser.json(), function(req, res) {
                console.log(req.body);

                var data = req.body;
                sendpulse.init(data.client_id, data.client_secret, '/tmp/');
                sendpulse.getToken((result) => {
                    res.send(result);
                });

            });

            app.post('/sendpulse', bodyParser.json(), function(req, res) {
                console.log(req.body);
                var parameters = req.body;
                if(!parameters)
                {
                    res.send({message: "request parameters undefined!", error: "no_parameters", error_code: 406});
                    return;
                }
                if(!parameters.url)
                {
                    res.send({message: "request function undefined!", error: "no_parameters", error_code: 407});
                    return;
                }

                sendpulse.generalCall(parameters, function(response){
                    res.send(response);
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
        }),
        new HtmlWebpackPlugin({
            template: './synchdialog.html',
            filename: 'synchdialog.html',
            chunks: ['synchdialog']
        })
    ]
};