module.exports = {
  entries: {
    main: 'src/index.js',
    chart: 'src/expendPlugins/chart/plugin.js',
    export: 'src/expendPlugins/exportXlsx/plugin.js',
    print: 'src/expendPlugins/print/plugin.js'
  },
  splitChunks: 'split',
  chunkNames: {
    vendor: 'luckysheet-vendor',
    core: 'luckysheet-core'
  },
  vendors: [
    'lodash',
    'jstat',
    'numeral',
    'dayjs',
    'jquery',
    'flatpickr'
  ]
};