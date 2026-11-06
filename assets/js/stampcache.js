/* Ассистент — кэш печати: хранит отредактированную печать в памяти браузера 30 минут, затем стирает сама */
(function () {
  'use strict';
  const KEY = 'assistant.stamp.cache.v1', TTL = 30 * 60 * 1000;
  let timer = null; const listeners = [];
  const emit = () => listeners.forEach(f => { try { f(get()); } catch (e) {} });
  function read() { try { return JSON.parse(localStorage.getItem(KEY) || 'null'); } catch (e) { return null; } }
  function clear() { try { localStorage.removeItem(KEY); } catch (e) {} if (timer) { clearTimeout(timer); timer = null; } emit(); }
  function schedule(rec) { if (timer) clearTimeout(timer); if (rec) timer = setTimeout(clear, Math.max(0, rec.expires - Date.now()) + 50); }
  function get() {
    const rec = read(); if (!rec) return null;
    if (!rec.dataUrl || Date.now() >= rec.expires) { try { localStorage.removeItem(KEY); } catch (e) {} return null; }
    return rec;
  }
  function save(dataUrl, w, h) {
    const rec = { dataUrl, w, h, saved: Date.now(), expires: Date.now() + TTL };
    try { localStorage.setItem(KEY, JSON.stringify(rec)); } catch (e) { return false; }
    schedule(rec); emit(); return true;
  }
  const remaining = () => { const r = get(); return r ? Math.max(0, r.expires - Date.now()) : 0; };
  function bytesToDataUrl(bytes) { let s = ''; for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000)); return 'data:image/png;base64,' + btoa(s); }
  window.addEventListener('storage', e => { if (e.key === KEY) { schedule(get()); emit(); } });
  schedule(get());
  window.StampCache = { save, get, clear, remaining, bytesToDataUrl, TTL, onChange: f => listeners.push(f) };
})();
