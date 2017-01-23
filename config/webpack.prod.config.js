let webpack = require('webpack');
let path = require('path');
let CopyWebpackPlugin = require('copy-webpack-plugin');

module.exports = {
    entry: {
        app: 'index.tsx',
        vendor: ['react', 'react-dom']
    },
    output: {
        filename: 'bundle.js',
        publicPath: '/',
        path: path.resolve('dist')
    },
    resolve: {
        extensions: ['', '.ts', '.tsx', '.js', '.jsx'],
        modulesDirectories: ['src/public', 'node_modules'],
    },
    module: {
        loaders: [
            { test: /\.tsx?$/, loaders: ['babel', 'ts-loader'] },
            {
                test: /\.css$/,
                loaders: [
                    'style?sourceMap',
                    'css?modules&importLoaders=1&localIdentName=[path]___[name]__[local]___[hash:base64:5]'
                ]
            }
        ]
    },
    plugins: [
        new CopyWebpackPlugin([
            { from: 'node_modules/office-ui-fabric-react/dist/css/fabric.min.css'},
            { from: 'node_modules/office-ui-fabric-react/dist/css/fabric.rtl.min.css'},
            { from: 'src/static' }            
        ]),
        new webpack.DefinePlugin({
            'process.env': {
                'NODE_ENV': JSON.stringify('production')
            }
        }),
        new webpack.optimize.CommonsChunkPlugin({
            name: 'vendor',
            filename: 'vendor.js',
            minChunks: Infinity
        }),
        new webpack.NamedModulesPlugin(),
        new webpack.optimize.UglifyJsPlugin({
            compress: { warnings: false },
            comments: false,
            mangle: false,
            minimize: false
        })
    ]
};