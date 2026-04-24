import Store from '../store';
import sheetmanage from './sheetmanage';
import {changeSheetContainerSize} from './resize';
import {jfrefreshgrid_rhcw} from '../global/refresh';
import server from './server';
import luckysheetPostil from './postil';
import imageCtrl from './imageCtrl';

let luckysheetZoomTimeout = null;

// 修复：声明未定义变量
let currentWheelZoom = null;

export function zoomChange(ratio) {
  if (!Store.flowdata || Store.flowdata.length === 0) {
    return;
  }

  clearTimeout(luckysheetZoomTimeout);
  luckysheetZoomTimeout = setTimeout(() => {
    if (Store.clearjfundo) {
      Store.jfredo.push({
        type: 'zoomChange',
        zoomRatio: Store.zoomRatio,
        curZoomRatio: ratio,
        sheetIndex: Store.currentSheetIndex,
      });
    }

    currentWheelZoom = null;
    Store.zoomRatio = ratio;

    const currentSheet = sheetmanage.getSheetByIndex();

    // 批注
    luckysheetPostil.buildAllPs(currentSheet.data);

    // 图片
    imageCtrl.images = currentSheet.images;
    imageCtrl.allImagesShow();
    imageCtrl.init();

    if (!currentSheet.config) {
      currentSheet.config = {};
    }

    if (!currentSheet.config.sheetViewZoom) {
      currentSheet.config.sheetViewZoom = {};
    }

    let type = currentSheet.config.curentsheetView;
    // 修复：安全判断 undefined
    if (type === undefined || type === null || type === '') {
      type = 'viewNormal';
    }

    currentSheet.config.sheetViewZoom[`${type}ZoomScale`] = ratio;

    server.saveParam('all', Store.currentSheetIndex, Store.zoomRatio, { k: 'zoomRatio' });
    server.saveParam('cg', Store.currentSheetIndex, currentSheet.config.sheetViewZoom, { k: 'sheetViewZoom' });

    zoomRefreshView();
  }, 100);
}

export function zoomRefreshView() {
  jfrefreshgrid_rhcw(Store.flowdata.length, Store.flowdata[0].length);
  changeSheetContainerSize();
}

export function zoomInitial() {
  const ZOOM_WHEEL_STEP = 0.02;
  const ZOOM_STEP = 0.1;
  const MAX_ZOOM_RATIO = 4;
  const MIN_ZOOM_RATIO = 0.1;

  // 修复：所有 === 风险、ESLint、未声明变量
  $('#luckysheet-zoom-minus').click(function () {
    let currentRatio;

    // 修复：安全判断 undefined
    if (Store.zoomRatio === undefined || Store.zoomRatio === null) {
      currentRatio = 1;
      Store.zoomRatio = 1;
    } else {
      currentRatio = Math.ceil(Store.zoomRatio * 10) / 10;
    }

    currentRatio = currentRatio - ZOOM_STEP;

    // 修复：浮点数 === 高危风险
    if (Math.abs(currentRatio - Store.zoomRatio) < 0.001) {
      currentRatio -= ZOOM_STEP;
    }

    if (currentRatio <= MIN_ZOOM_RATIO) {
      currentRatio = MIN_ZOOM_RATIO;
    }

    zoomChange(currentRatio);
    zoomNumberDomBind(currentRatio);
  });

  $('#luckysheet-zoom-plus').click(function () {
    let currentRatio;

    if (Store.zoomRatio === undefined || Store.zoomRatio === null) {
      currentRatio = 1;
      Store.zoomRatio = 1;
    } else {
      currentRatio = Math.floor(Store.zoomRatio * 10) / 10;
    }

    currentRatio += ZOOM_STEP;

    if (Math.abs(currentRatio - Store.zoomRatio) < 0.001) {
      currentRatio += ZOOM_STEP;
    }

    if (currentRatio >= MAX_ZOOM_RATIO) {
      currentRatio = MAX_ZOOM_RATIO;
    }

    zoomChange(currentRatio);
    zoomNumberDomBind(currentRatio);
  });

  $('#luckysheet-zoom-slider').mousedown(function (e) {
    const xoffset = $(this).offset().left;
    const pageX = e.pageX;
    const currentRatio = positionToRatio(pageX - xoffset);

    zoomChange(currentRatio);
    zoomNumberDomBind(currentRatio);
  });

  $('#luckysheet-zoom-cursor').mousedown(function (e) {
    const curentX = e.pageX;
    const cursorLeft = parseFloat($('#luckysheet-zoom-cursor').css('left'));

    $('#luckysheet-zoom-cursor').css('transition', 'none');

    $(document).off('mousemove.zoomCursor').on('mousemove.zoomCursor', function (event) {
      const moveX = event.pageX;
      const offsetX = moveX - curentX;
      let pos = cursorLeft + offsetX;
      let currentRatio = positionToRatio(pos);

      if (currentRatio > MAX_ZOOM_RATIO) {
        currentRatio = MAX_ZOOM_RATIO;
        pos = 100;
      }

      if (currentRatio < MIN_ZOOM_RATIO) {
        currentRatio = MIN_ZOOM_RATIO;
        pos = 0;
      }

      zoomChange(currentRatio);
      const r = `${Math.round(currentRatio * 100)}%`;
      $('#luckysheet-zoom-ratioText').html(r);
      $('#luckysheet-zoom-cursor').css('left', pos - 4);
    });

    $(document).off('mouseup.zoomCursor').on('mouseup.zoomCursor', function () {
      $(document).off('.zoomCursor');
      $('#luckysheet-zoom-cursor').css('transition', 'all 0.3s');
    });

    e.stopPropagation();
  }).click(function (e) {
    e.stopPropagation();
  });

  $('#luckysheet-zoom-ratioText').click(function () {
    zoomChange(1);
    zoomNumberDomBind(1);
  });

  zoomNumberDomBind(Store.zoomRatio);

  currentWheelZoom = null;

  document.addEventListener(
    'wheel',
    function (ev) {
      if (!ev.ctrlKey || !ev.deltaY) {
        return;
      }

      if (currentWheelZoom === null) {
        currentWheelZoom = Store.zoomRatio || 1;
      }

      currentWheelZoom += ev.deltaY < 0 ? ZOOM_WHEEL_STEP : -ZOOM_WHEEL_STEP;

      if (currentWheelZoom >= MAX_ZOOM_RATIO) {
        currentWheelZoom = MAX_ZOOM_RATIO;
      } else if (currentWheelZoom < MIN_ZOOM_RATIO) {
        currentWheelZoom = MIN_ZOOM_RATIO;
      }

      zoomChange(currentWheelZoom);
      zoomNumberDomBind(currentWheelZoom);
      ev.preventDefault();
      ev.stopPropagation();
    },
    { capture: true, passive: false }
  );

  document.addEventListener(
    'keydown',
    function (ev) {
      if (!ev.ctrlKey) {return;}

      let handled = false;
      let zoom = Store.zoomRatio || 1;

      if (ev.key === '-' || ev.which === 189) {
        zoom -= ZOOM_STEP;
        handled = true;
      } else if (ev.key === '+' || ev.which === 187) {
        zoom += ZOOM_STEP;
        handled = true;
      } else if (ev.key === '0' || ev.which === 48) {
        zoom = 1;
        handled = true;
      }

      if (handled) {
        ev.preventDefault();

        if (zoom >= MAX_ZOOM_RATIO) {
          zoom = MAX_ZOOM_RATIO;
        } else if (zoom < MIN_ZOOM_RATIO) {
          zoom = MIN_ZOOM_RATIO;
        }

        zoomChange(zoom);
        zoomNumberDomBind(zoom);
      }
    },
    { capture: true }
  );
}

function positionToRatio(pos) {
  let ratio = 1;
  if (pos < 50) {
    ratio = Math.round((pos * 1.8 / 100 + 0.1) * 100) / 100;
  } else if (pos > 50) {
    ratio = Math.round(((pos - 50) * 6 / 100 + 1) * 100) / 100;
  }
  return ratio;
}

function zoomSlierDomBind(ratio) {
  let domPos = 50;
  if (ratio < 1) {
    domPos = Math.round((ratio - 0.1) * 100 / 0.18) / 10;
  } else if (ratio > 1) {
    domPos = Math.round((ratio - 1) * 100 / 0.6) / 10 + 50;
  }
  $('#luckysheet-zoom-cursor').css('left', domPos - 4);
}

export function zoomNumberDomBind(ratio) {
  const r = `${Math.round(ratio * 100)}%`;
  $('#luckysheet-zoom-ratioText').html(r);
  zoomSlierDomBind(ratio);
}
