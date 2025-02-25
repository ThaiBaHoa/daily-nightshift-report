const path = require('path');

module.exports = function override(config, env) {
  // Thêm alias cho thư mục data
  config.resolve.alias = {
    ...config.resolve.alias,
    '@data': path.resolve(__dirname, 'data')
  };

  return config;
};
