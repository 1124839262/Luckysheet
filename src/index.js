import './utils/math'
import {luckysheet} from './core'
import __firefox from './utils/polyfill'

if (window.addEventListener && (navigator.userAgent.indexOf("Firefox") > 0)) {
    __firefox();
}

export default luckysheet;