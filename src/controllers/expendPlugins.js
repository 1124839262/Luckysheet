import {chart} from '../expendPlugins/chart/plugin';
import {print} from '../expendPlugins/print/plugin';
import {exportXlsx} from '../expendPlugins/exportXlsx/plugin';

const pluginsObj = {
  'chart':chart,
  'print':print,
  'exportXlsx':exportXlsx
};

const isDemo = true;

/**
 * Register plugins
 *
 * plugins:[
 * {name:'chart'},
 * {name:'print'},
 * {name:'exportXlsx',config:{url:''}}
 * ]
 */
function initPlugins(plugins , options){
  if(plugins.length){
    plugins.forEach(plugin => {
      if (pluginsObj[plugin.name]) {
        pluginsObj[plugin.name](options, plugin.config, isDemo);
      } else {
        console.warn(`Luckysheet plugin "${plugin.name}" is not found. Please make sure the plugin is properly registered.`);
      }
    });
  }
}

export {
  initPlugins
};