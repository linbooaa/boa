// ==UserScript==
// @name         社群入稿中控台 SharePoint 橋接
// @namespace    boa.console
// @version      1.0
// @description  讓中控台用你已登入的 SharePoint 身分讀取進稿表與下載素材
// @match        https://linbooaa.github.io/boa/*
// @match        file:///*/社群入稿中控台*.html
// @grant        GM_xmlhttpRequest
// @grant        unsafeWindow
// @connect      sharepoint.com
// @connect      asus-my.sharepoint.com
// @connect      asus.sharepoint.com
// @run-at       document-start
// ==/UserScript==

(function () {
  'use strict';
  function gmFetch(url, opts) {
    opts = opts || {};
    return new Promise(function (resolve, reject) {
      GM_xmlhttpRequest({
        method: opts.method || 'GET',
        url: url,
        headers: Object.assign({ 'Accept': opts.accept || '*/*' }, opts.headers || {}),
        responseType: opts.responseType || 'arraybuffer',
        anonymous: false,
        onload: function (r) {
          if (r.status >= 200 && r.status < 300) resolve({ status: r.status, data: r.response, finalUrl: r.finalUrl, headers: r.responseHeaders });
          else reject(new Error('HTTP ' + r.status + ' ' + url.slice(0, 80)));
        },
        onerror: function (e) { reject(new Error('網路錯誤：' + (e && e.error ? e.error : url.slice(0, 80)))); },
        ontimeout: function () { reject(new Error('逾時')); }
      });
    });
  }

  async function resolveFile(url) {
    var u = new URL(url);
    var sd = u.searchParams.get('sourcedoc');
    var site = (u.pathname.match(/^\/(personal|sites|teams)\/[^\/]+/) || [''])[0];
    if (sd) return { host: u.origin, site: site, fileId: sd.replace(/[{}]/g, '') };
    var r = await gmFetch(url, { responseType: 'text', accept: 'text/html' });
    var f = new URL(r.finalUrl || url);
    sd = f.searchParams.get('sourcedoc');
    site = (f.pathname.match(/^\/(personal|sites|teams)\/[^\/]+/) || [''])[0];
    if (sd) return { host: f.origin, site: site, fileId: sd.replace(/[{}]/g, '') };
    if (/\.(xlsx|xlsm|png|jpe?g|gif|mp4|mov)$/i.test(f.pathname)) return { direct: f.origin + f.pathname };
    throw new Error('解析不出檔案位置，請改用含 sourcedoc 的連結（在 Excel Online 網址列複製）');
  }

  async function fetchFile(url) {
    var info = await resolveFile(url);
    var dl;
    if (info.direct) dl = info.direct;
    else dl = info.host + info.site + "/_api/web/GetFileById('" + info.fileId + "')/$value";
    var r = await gmFetch(dl, { responseType: 'arraybuffer' });
    return r.data;
  }

  async function fetchMedia(url) {
    var u = new URL(url);
    if (/\/:[a-z]:\//.test(u.pathname)) { u.searchParams.set('download', '1'); }
    var r = await gmFetch(u.toString(), { responseType: 'arraybuffer' });
    var name = '';
    var cd = /filename\*?=(?:UTF-8'')?"?([^";\r\n]+)/i.exec(r.headers || '');
    if (cd) { try { name = decodeURIComponent(cd[1]); } catch (e) { name = cd[1]; } }
    if (!name) { var p = new URL(r.finalUrl || url).pathname; name = decodeURIComponent(p.split('/').pop() || ''); }
    var ct = (/content-type:\s*([^\r\n;]+)/i.exec(r.headers || '') || [])[1] || 'application/octet-stream';
    return { data: r.data, name: name, type: ct };
  }

  var bridge = { version: '1.0', fetchFile: fetchFile, fetchMedia: fetchMedia };
  try { unsafeWindow.__spBridge = bridge; } catch (e) { window.__spBridge = bridge; }
  try { unsafeWindow.dispatchEvent(new CustomEvent('spbridge-ready')); } catch (e) {}
})();
