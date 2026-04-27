import {getObjType} from '../utils/util';
import Store from '../store';

/**
 * 获取合并单元格边界信息
 * @param {Array} d - 表格数据
 * @param {Number} row_index - 行索引
 * @param {Number} col_index - 列索引
 * @returns {Object|null} 合并单元格的边界信息
 */
function mergeborer(d, row_index, col_index) {
  if (d == null || d[row_index] == null) {
    return null;
  }
  const value = d[row_index][col_index];

  if (getObjType(value) == 'object' && 'mc' in value) {
    const margeMaindata = value['mc'];
    if (margeMaindata == null) {
      return null;
    }
    col_index = margeMaindata.c;
    row_index = margeMaindata.r;

    if (d[row_index][col_index] == null) {
      return null;
    }
    const col_rs = d[row_index][col_index].mc.cs;
    const row_rs = d[row_index][col_index].mc.rs;

    const margeMain = d[row_index][col_index].mc;

    let start_r, end_r, row, row_pre;
    for (let r = row_index; r < margeMain.rs + row_index; r++) {
      if (r == 0) {
        start_r = -1;
      } else {
        start_r = Store.visibledatarow[r - 1] - 1;
      }

      end_r = Store.visibledatarow[r];

      if (row_pre == null) {
        row_pre = start_r;
        row = end_r;
      } else {
        row += end_r - start_r - 1;
      }
    }

    let start_c, end_c, col, col_pre;
    for (let c = col_index; c < margeMain.cs + col_index; c++) {
      if (c == 0) {
        start_c = 0;
      } else {
        start_c = Store.visibledatacolumn[c - 1];
      }

      end_c = Store.visibledatacolumn[c];

      if (col_pre == null) {
        col_pre = start_c;
        col = end_c;
      } else {
        col += end_c - start_c;
      }
    }

    return {
      row: [row_pre, row, row_index, row_index + row_rs - 1],
      column: [col_pre, col, col_index, col_index + col_rs - 1],
    };
  } else {
    return null;
  }
}

export {mergeborer};
