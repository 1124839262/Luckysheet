/**
 * polyfill event for modern browsers (replaced old IE / Firefox legacy code)
 */
function __firefox() {
  // 替换 __defineGetter__ 废弃写法 → 现代 Object.defineProperty
  if (!HTMLElement.prototype.runtimeStyle) {
    Object.defineProperty(HTMLElement.prototype, 'runtimeStyle', {
      get: __element_style
    });
  }

  if (!window.event) {
    Object.defineProperty(window, 'event', {
      get: __window_event
    });
  }

  if (!Event.prototype.srcElement) {
    Object.defineProperty(Event.prototype, 'srcElement', {
      get: __event_srcElement
    });
  }
}

function __element_style() {
  return this.style;
}

function __window_event() {
  return __window_event_constructor();
}

function __event_srcElement() {
  return this.target;
}

// ========== 核心修复：移除 arguments / caller ==========
function __window_event_constructor(e) {
  // 现代标准方式：优先从调用参数获取
  if (e instanceof Event) {return e;}

  // 兼容旧 IE
  if (window.event) {return window.event;}

  return null;
}

export default __firefox;
