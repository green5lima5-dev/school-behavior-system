/* Client compatibility adapter for the original Apps Script UI. */
(() => {
  'use strict';
  function runner(success, failure) {
    return new Proxy({}, {
      get(_obj, name) {
        if (name === 'withSuccessHandler') return fn => runner(fn, failure);
        if (name === 'withFailureHandler') return fn => runner(success, fn);
        if (name === 'then') return undefined;
        return (...args) => {
          fetch('/api/rpc', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            credentials: 'same-origin',
            body: JSON.stringify({method: String(name), args})
          }).then(async res => {
            let payload;
            try { payload = await res.json(); }
            catch { throw new Error('Backend ส่งข้อมูลไม่ถูกต้อง กรุณาตรวจการตั้งค่า API'); }
            if (!res.ok || !payload.ok) throw new Error(payload.error || 'API error');
            if (typeof success === 'function') success(payload.result);
          }).catch(error => {
            if (typeof failure === 'function') failure(error);
            else console.error('API request failed:', error);
          });
        };
      }
    });
  }
  window.google = window.google || {};
  window.google.script = window.google.script || {};
  window.google.script.run = runner();
})();
