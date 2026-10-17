/* -*- Mode: JS; tab-width: 2; indent-tabs-mode: nil; js-indent-level: 2; fill-column: 100 -*- */
// SPDX-License-Identifier: MIT
// Source: Derived from ZetaOffice WebAssembly example (https://github.com/zetaoffice/zeta-wasm-examples)
// Purpose: Runs WebAssembly-based LibreOffice (Calc) in a separate Web Worker thread.

// Debugging note:
// Switch the web worker in the browsers debug tab to debug this code.
// It's the "em-pthread" web worker with the most memory usage, where "zetajs" is defined.

'use strict';

// Polyfill navigator.permissions.query for Firefox / non-standard environments
(function(scope) {
  try {
    const nav = scope.navigator;
    if (!nav || !nav.permissions) return;
    const proto = (typeof scope.Permissions !== 'undefined' && scope.Permissions.prototype)
      ? scope.Permissions.prototype
      : nav.permissions;
    if (!proto || typeof proto.query !== 'function') return;
    if (proto.query.__clipboard_polyfilled__) return;

    const originalQuery = proto.query;
    function createFakeStatus(name) {
      const target = (typeof scope.EventTarget === 'function') ? new scope.EventTarget() : {};
      return Object.assign(target, { name: name, state: 'prompt', onchange: null });
    }

    const patchedQuery = function(desc) {
      if (desc && typeof desc.name === 'string' && desc.name.startsWith('clipboard-')) {
        try {
          const res = originalQuery.call(this, desc);
          if (res && typeof res.catch === 'function') {
            return res.catch(() => createFakeStatus(desc.name));
          }
          return Promise.resolve(res);
        } catch (e) {
          return Promise.resolve(createFakeStatus(desc.name));
        }
      }
      return originalQuery.call(this, desc);
    };

    patchedQuery.__clipboard_polyfilled__ = true;
    try { proto.query = patchedQuery; } catch (e) {}
    try {
      if (nav.permissions.query !== patchedQuery) {
        nav.permissions.query = patchedQuery;
      }
    } catch (e) {}
  } catch (_) {}
})(typeof globalThis !== 'undefined' ? globalThis : self);

// global variables - zetajs environment:
let zetajs, css;

// = global variables (some are global for easier debugging) =
// common variables:
let context, desktop, xModel, ctrl;


function demo() {
  context = zetajs.getUnoComponentContext();
  const bean_overwrite = new css.beans.PropertyValue({Name: 'Overwrite', Value: true});
  const bean_odt_export = new css.beans.PropertyValue({Name: 'FilterName', Value: 'writer8'});
  desktop = css.frame.Desktop.create(context);

  zetajs.mainPort.onmessage = function (e) {
    switch (e.data.cmd) {
    case 'upload':
      loadFile(e.data.filename);
      break;
    case 'download':
      xModel.store();
      zetajs.mainPort.postMessage({cmd: 'download', id: e.data.id});
      break;
    default:
      throw Error('Unknown message command: ' + e.data.cmd);
    }
  }
  zetajs.mainPort.postMessage({cmd: 'thr_running'});
}

function loadFile(filename) {
  if (xModel) {
    try {
      const xCloseable = (css && css.util && css.util.XCloseable) ? css.util.XCloseable.query(xModel) : null;
      if (xCloseable) {
        xCloseable.close(true);
      } else if (typeof xModel.close === 'function') {
        xModel.close(true);
      } else if (typeof xModel.dispose === 'function') {
        xModel.dispose();
      }
    } catch (_) {
      try {
        if (typeof xModel.dispose === 'function') {
          xModel.dispose();
        }
      } catch (_) {}
    }
    xModel = null;
    ctrl = null;
  }

  const in_path = 'file:///tmp/office/' + filename;
  xModel = desktop.loadComponentFromURL(in_path, '_default', 0, []);
  ctrl = xModel.getCurrentController();
  ctrl.getFrame().getContainerWindow().FullScreen = true;
  zetajs.mainPort.postMessage({cmd: 'ui_ready'});
}

Module.zetajs.then(function(pZetajs) {
  // initializing zetajs environment:
  zetajs = pZetajs;
  css = zetajs.uno.com.sun.star;
  demo();  // launching demo
});

/* vim:set shiftwidth=2 softtabstop=2 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */
