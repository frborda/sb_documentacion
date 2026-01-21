module.exports = {
  apps: [{
    namespace: 'recinto',
    name: 'recinto-replier',
    script: './dist/index.js',
    exec_mode: 'fork',
    watch: false,
    error_file: './recinto-replier_err.log',
    out_file: './recinto-replier_out.log'
  }]
}
