// ==UserScript==
// @name         正方教务 · 毕业设计选题抓取助手
// @namespace    jwxt-topic-helper
// @version      1.0.0
// @description  在学生选题页一键抓取全部课题（列表+明细），导出 Excel/CSV。复用你自己的登录状态，不上传任何数据。
// @author       dsh
// @match        https://jwxt.sztu.edu.cn/*
// @match        https://jwxx.sztu.edu.cn/*
// @run-at       document-idle
// @grant        none
// @noframes
// ==/UserScript==

/*
 * 使用：Edge 装好 Tampermonkey → 新建脚本 → 粘贴本文件 → 保存
 *      打开 教务系统 → 毕业设计 → 学生选题 → 右下角出现面板
 *      ① 抓列表（约 10 秒） → ② 抓明细（307 条约 2~3 分钟） → ③ 导出 Excel
 *
 * 控制台命令（面板够用可以不管）：
 *      __JWH.list()  __JWH.detail()  __JWH.exportXlsx()  __JWH.exportCsv()  __JWH.status()
 */
(function () {
  'use strict';

  const VERSION = '1.0.0';

  const CFG = {
    frameHint: 'bysj/xsxt.do',            // 列表所在 iframe
    listFallback: '/jsxsd/bysj/xsxt.do',
    detailTemplate: '/jsxw/view/byjs/kt_Mx_14655.html?tid={tid}',   // 兜底（实测正确值由 to_mx 截获）
    maxPages: 40,
    concurrency: 3,
    delayMs: 180,
    retry: 2,
  };

  const S = {
    tids: [],
    detail: new Map(),
    template: null,
    running: false,
    stop: false,
    includeContact: false,
    listTotal: null,
  };

  /* ============================== 工具 ============================== */
  const norm = (s) => String(s == null ? '' : s).replace(/\s+/g, '').replace(/[*＊:：|｜]/g, '');
  const sq = (s) => String(s == null ? '' : s).replace(/\s+/g, ' ').trim();
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const abs = (u, base) => { try { return new URL(u, base || location.href).href; } catch { return String(u); } };
  const clip = (s, n) => { s = String(s == null ? '' : s); return s.length > n ? s.slice(0, n) + '…' : s; };

  function download(blob, name) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 8000);
  }

  function downloadText(text, name, mime) {
    download(new Blob([text], { type: (mime || 'text/plain') + ';charset=utf-8' }), name);
  }

  async function getText(url, init) {
    let lastErr;
    for (let attempt = 0; attempt <= CFG.retry; attempt++) {
      try {
        const res = await fetch(url, Object.assign({ credentials: 'include' }, init || {}));
        if (!res.ok) throw new Error('HTTP ' + res.status);
        return await res.text();
      } catch (e) { lastErr = e; await sleep(400 * (attempt + 1)); }
    }
    throw lastErr || new Error('fetch failed');
  }

  async function pool(items, n, fn, onTick) {
    let i = 0, done = 0, fail = 0;
    const workers = Array.from({ length: Math.max(1, n) }, async () => {
      while (i < items.length) {
        if (S.stop) return;
        const idx = i++;
        try { await fn(items[idx], idx); }
        catch (e) { fail++; }
        done++;
        if (onTick) onTick(done, items.length, fail);
        await sleep(CFG.delayMs);
      }
    });
    await Promise.all(workers);
    return { done, fail };
  }

  /* ============================== xlsx 生成（零依赖） ============================== */
  const XLSX = (() => {
    const enc = new TextEncoder();
    const T = (() => {
      const t = new Uint32Array(256);
      for (let n = 0; n < 256; n++) {
        let c = n;
        for (let k = 0; k < 8; k++) c = c & 1 ? (0xEDB88320 ^ (c >>> 1)) >>> 0 : c >>> 1;
        t[n] = c >>> 0;
      }
      return t;
    })();
    const crc = (b) => {
      let c = 0xFFFFFFFF;
      for (let i = 0; i < b.length; i++) c = (T[(c ^ b[i]) & 0xFF] ^ (c >>> 8)) >>> 0;
      return (c ^ 0xFFFFFFFF) >>> 0;
    };

    function zipStore(files) {
      const parts = [], central = [];
      let offset = 0;
      const now = new Date();
      const dosTime = ((now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1)) & 0xFFFF;
      const dosDate = (((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate()) & 0xFFFF;

      for (const f of files) {
        const name = enc.encode(f.name);
        const data = f.data;
        const c = crc(data);

        const lh = new Uint8Array(30 + name.length);
        const dv = new DataView(lh.buffer);
        dv.setUint32(0, 0x04034B50, true);
        dv.setUint16(4, 20, true);
        dv.setUint16(6, 0x0800, true);
        dv.setUint16(8, 0, true);
        dv.setUint16(10, dosTime, true);
        dv.setUint16(12, dosDate, true);
        dv.setUint32(14, c, true);
        dv.setUint32(18, data.length, true);
        dv.setUint32(22, data.length, true);
        dv.setUint16(26, name.length, true);
        dv.setUint16(28, 0, true);
        lh.set(name, 30);
        parts.push(lh, data);

        const cd = new Uint8Array(46 + name.length);
        const cv = new DataView(cd.buffer);
        cv.setUint32(0, 0x02014B50, true);
        cv.setUint16(4, 20, true);
        cv.setUint16(6, 20, true);
        cv.setUint16(8, 0x0800, true);
        cv.setUint16(10, 0, true);
        cv.setUint16(12, dosTime, true);
        cv.setUint16(14, dosDate, true);
        cv.setUint32(16, c, true);
        cv.setUint32(20, data.length, true);
        cv.setUint32(24, data.length, true);
        cv.setUint16(28, name.length, true);
        cv.setUint32(42, offset, true);
        cd.set(name, 46);
        central.push(cd);
        offset += lh.length + data.length;
      }

      const cdSize = central.reduce((s, c) => s + c.length, 0);
      const eocd = new Uint8Array(22);
      const ev = new DataView(eocd.buffer);
      ev.setUint32(0, 0x06054B50, true);
      ev.setUint16(8, files.length, true);
      ev.setUint16(10, files.length, true);
      ev.setUint32(12, cdSize, true);
      ev.setUint32(16, offset, true);

      const all = parts.concat(central, [eocd]);
      const out = new Uint8Array(all.reduce((s, p) => s + p.length, 0));
      let pos = 0;
      for (const p of all) { out.set(p, pos); pos += p.length; }
      return out;
    }

    const esc = (s) => String(s == null ? '' : s)
      .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

    const colName = (n) => {
      let s = '';
      while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
      return s;
    };

    function build(sheetName, headers, rows, widths) {
      const files = [];
      const add = (name, text) => files.push({ name, data: enc.encode(text) });

      add('[Content_Types].xml',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
        '</Types>');

      add('_rels/.rels',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
        '</Relationships>');

      add('xl/workbook.xml',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheets><sheet name="' + esc(sheetName).slice(0, 28) + '" sheetId="1" r:id="rId1"/></sheets></workbook>');

      add('xl/_rels/workbook.xml.rels',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
        '</Relationships>');

      add('xl/styles.xml',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        '<fonts count="2"><font><sz val="10"/><name val="微软雅黑"/></font>' +
        '<font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="微软雅黑"/></font></fonts>' +
        '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
        '<fill><patternFill patternType="solid"><fgColor rgb="FF1F4E79"/><bgColor indexed="64"/></patternFill></fill></fills>' +
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
        '<cellXfs count="2">' +
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>' +
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1">' +
        '<alignment horizontal="center" vertical="center" wrapText="1"/></xf>' +
        '</cellXfs>' +
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
        '</styleSheet>');

      const cols = widths && widths.length
        ? '<cols>' + widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('') + '</cols>'
        : '';

      const rowXml = ['<row r="1" ht="26" customHeight="1">' +
        headers.map((h, i) => `<c r="${colName(i + 1)}1" s="1" t="inlineStr"><is><t>${esc(h)}</t></is></c>`).join('') +
        '</row>'];

      rows.forEach((row, ri) => {
        const r = ri + 2;
        const cells = row.map((v, ci) => {
          if (v == null || v === '') return '';
          const ref = colName(ci + 1) + r;
          if (typeof v === 'number' && isFinite(v)) return `<c r="${ref}"><v>${v}</v></c>`;
          return `<c r="${ref}" t="inlineStr"><is><t>${esc(v)}</t></is></c>`;
        }).join('');
        rowXml.push(`<row r="${r}">${cells}</row>`);
      });

      add('xl/worksheets/sheet1.xml',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        '<sheetViews><sheetView workbookViewId="0"><pane xSplit="3" ySplit="1" topLeftCell="D2" activePane="bottomRight" state="frozen"/></sheetView></sheetViews>' +
        '<sheetFormatPr defaultRowHeight="14"/>' + cols +
        '<sheetData>' + rowXml.join('') + '</sheetData>' +
        `<autoFilter ref="A1:${colName(headers.length)}${rows.length + 1}"/>` +
        '</worksheet>');

      return zipStore(files);
    }

    return { build };
  })();

  /* ============================== 列表帧定位 ============================== */
  function walk(win, depth, visit) {
    let doc;
    try { doc = win.document; } catch { return; }
    if (!doc || !doc.querySelectorAll) return;
    try { if (visit(win, doc) === false) return; } catch (e) { }
    let frames = [];
    try { frames = [...doc.querySelectorAll('iframe,frame')]; } catch { return; }
    for (const f of frames) {
      let w;
      try { w = f.contentWindow; } catch { continue; }
      if (w && depth < 6) walk(w, depth + 1, visit);
    }
  }

  function findListFrame() {
    let found = null;
    walk(window, 0, (win, doc) => {
      if (found) return false;
      let href = '';
      try { href = win.location.href; } catch { return; }
      if (href.includes(CFG.frameHint) && doc.querySelectorAll('tr').length > 10) {
        found = { win, doc, href };
        return false;
      }
      return true;
    });
    if (!found) {
      walk(window, 0, (win, doc) => {
        if (found) return false;
        if (doc.querySelectorAll('tr').length > 10) {
          let href = ''; try { href = win.location.href; } catch (e) { }
          found = { win, doc, href };
          return false;
        }
        return true;
      });
    }
    return found;
  }

  /* ============================== 列表抓取 ============================== */
  function buildListRequest(frame, page) {
    const doc = frame.doc;
    let form = null;
    try { form = doc.querySelector('form'); } catch (e) { }

    let action = CFG.listFallback, method = 'GET';
    const params = new URLSearchParams();

    if (form) {
      action = form.getAttribute('action') || action;
      method = (form.getAttribute('method') || 'GET').toUpperCase();
      for (const el of form.querySelectorAll('input[name],select[name],textarea[name]')) {
        if (el.type === 'submit' || el.type === 'button') continue;
        params.append(el.name, el.value == null ? '' : el.value);
      }
    }
    if (params.has('pageIndex')) params.set('pageIndex', String(page));
    else params.append('pageIndex', String(page));

    let base;
    try { base = abs(action, frame.win.location.href); } catch (e) { base = abs(action); }

    if (method === 'GET') {
      const u = new URL(base);
      params.forEach((v, k) => u.searchParams.set(k, v));
      return { url: u.href, init: { method: 'GET' } };
    }
    return {
      url: base,
      init: {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8' },
        body: params.toString(),
      },
    };
  }

  function parseListHtml(html) {
    const doc = new DOMParser().parseFromString(html, 'text/html');
    const table = [...doc.querySelectorAll('table')]
      .map((t) => ({ t, rows: t.querySelectorAll('tr').length }))
      .sort((a, b) => b.rows - a.rows)[0];
    if (!table) return { rows: [], total: null };

    const trs = [...table.t.querySelectorAll('tr')];
    const head = trs.length ? [...trs[0].children].map((c) => sq(c.textContent)) : [];
    const out = [];

    for (const tr of trs.slice(1)) {
      const cells = [...tr.children].map((c) => sq(c.textContent));
      if (!cells.length) continue;
      const m = tr.innerHTML.match(/to_mx\('([0-9A-Fa-f]{16,})'\)/);
      if (!m) continue;
      const rec = { 课题ID: m[1] };
      head.forEach((h, i) => { if (h && h !== '操作' && cells[i] != null) rec[h] = cells[i]; });
      rec.题目名称 = rec['设计(论文)题目'] || '';
      rec.可报学生上限 = rec['可报人数上限'] || '';
      out.push(rec);
    }

    const mt = html.replace(/\s+/g, ' ').match(/共\s*(\d+)\s*条/);
    return { rows: out, total: mt ? Number(mt[1]) : null };
  }

  async function grabList(onStatus) {
    const frame = findListFrame();
    if (!frame) throw new Error('没找到选题列表：请确认停在「毕业设计 → 学生选题」页面，且列表已显示');

    const all = new Map();
    for (let page = 1; page <= CFG.maxPages; page++) {
      if (S.stop) break;
      const req = buildListRequest(frame, page);
      let html;
      try { html = await getText(req.url, req.init); }
      catch (e) { onStatus && onStatus(`第 ${page} 页失败：${e.message}`); break; }

      const { rows, total } = parseListHtml(html);
      if (total && S.listTotal == null) S.listTotal = total;
      if (!rows.length) break;

      let added = 0;
      for (const r of rows) if (!all.has(r.课题ID)) { all.set(r.课题ID, r); added++; }
      if (onStatus) onStatus(`列表 第 ${page} 页：新增 ${added}，累计 ${all.size}` + (S.listTotal ? ` / 共 ${S.listTotal}` : ''));
      if (added === 0) break;
      if (S.listTotal && all.size >= S.listTotal) break;
      await sleep(120);
    }

    S.tids = [...all.values()];
    return S.tids.length;
  }

  /* ============================== 明细抓取 ============================== */
  function learnTemplate() {
    const frame = findListFrame();
    if (!frame || !S.tids.length) return CFG.detailTemplate;

    const win = frame.win;
    let captured = null;
    let orig = null;
    const fake = { focus() { }, close() { }, closed: false, location: { href: '' } };
    try {
      orig = win.open;
      win.open = (u) => { captured = abs(u, frame.href); return fake; };
      if (typeof win.to_mx === 'function') win.to_mx(S.tids[0].课题ID);
    } catch (e) { }
    finally { try { if (orig) win.open = orig; } catch (e) { } }

    if (captured) return captured.replace(S.tids[0].课题ID, '{tid}');
    return CFG.detailTemplate;
  }

  function detailUrl(tid) {
    return abs((S.template || CFG.detailTemplate).replace('{tid}', encodeURIComponent(tid)));
  }

  const DETAIL_MAP = [
    ['题目编号', 'ktbh'], ['题目名称', 'ktmc'], ['限选人数', 'xxrs'], ['年度', 'jb'],
    ['院系', 'skyxmc'], ['课题来源', 'ktly'], ['可报学生上限', 'kbxsrs'],
    ['学科门类', 'ktlb'], ['题目类型', 'tmlx'], ['论文研究方向', 'lwyjfx'],
    ['办公电话', 'zc'], ['手机号码', 'sjdc'], ['是否产业真题', 'sfcyzt'],
    ['限选专业', 'xbzymc'], ['合作单位名称', 'hzqymc'], ['产业导师姓名', 'qydsmc'],
    ['产业导师职称', 'qydszc'], ['联系方式', 'qydslxfs'],
    ['主要任务和工作内容', 'rwmb'], ['进度安排', 'jdap'], ['备注信息', 'bzxx'],
    ['送审单位', 'shyx'],
  ];

  function ctrlValue(el) {
    if (!el) return '';
    const tag = el.tagName.toLowerCase();
    if (tag === 'input') {
      if (el.type === 'checkbox') return el.checked || el.value === 'on' ? '是' : '否';
      if (el.type === 'radio') return el.checked ? (el.value || '是') : '';
      return el.value == null ? '' : String(el.value).trim();
    }
    if (tag === 'textarea') return (el.value || el.textContent || '').replace(/\s+/g, ' ').trim();
    if (tag === 'select') {
      const opts = [...(el.options || [])];
      const sel = opts[el.selectedIndex];
      if (sel) return sq(sel.textContent);
      const v = el.getAttribute('value') || el.value;
      const hit = opts.find((o) => o.getAttribute('value') === v || o.value === v);
      return hit ? sq(hit.textContent) : '';
    }
    return sq(el.textContent);
  }

  function parseDetailHtml(html) {
    const doc = new DOMParser().parseFromString(html, 'text/html');
    const fields = {};
    for (const [label, name] of DETAIL_MAP) {
      const el = doc.querySelector('[name="' + name + '"]') || doc.querySelector('#' + name);
      if (!el) continue;
      const v = ctrlValue(el);
      if (v) fields[label] = v;
    }
    return fields;
  }

  async function grabDetail(tid) {
    const html = await getText(detailUrl(tid));
    if (html.length < 3000 && /登录|用户名|login/i.test(html)) throw new Error('登录态可能已失效');
    const fields = parseDetailHtml(html);
    S.detail.set(tid, fields);
    return fields;
  }

  /* ============================== 合并 / 导出 ============================== */
  const CONTACT_KEYS = ['办公电话', '手机号码', '联系方式'];

  function merged() {
    return S.tids.map((r) => {
      const d = S.detail.get(r.课题ID) || {};
      const out = {};
      for (const [k, v] of Object.entries(r)) {
        if (k === '课题ID' || k === '设计(论文)题目' || k === '可报人数上限') continue;
        out[k] = v;
      }
      for (const [k, v] of Object.entries(d)) if (!out[k]) out[k] = v;
      out.详情链接 = detailUrl(r.课题ID);
      if (!S.includeContact) for (const k of CONTACT_KEYS) if (out[k]) out[k] = '（已隐藏）';
      return out;
    });
  }

  const COLS = ['序号', '题目编号', '题目名称', '院系', '年度', '课题来源', '限选人数', '可报学生上限',
    '已选人数', '课题选中人数', '是否辅修课题', '限选专业', '学科门类', '题目类型', '论文研究方向',
    '是否产业真题', '指导教师', '职称', '办公电话', '手机号码', '合作单位名称', '产业导师姓名',
    '产业导师职称', '联系方式', '主要任务和工作内容', '备注信息', '详情链接'];

  const WIDTHS = [5, 13, 42, 17, 7, 11, 9, 11, 9, 11, 11, 26, 9, 10, 16, 11, 10, 11, 13, 13,
    17, 13, 13, 13, 70, 22, 26];

  function toMatrix() {
    const rows = merged();
    return rows.map((r, i) => COLS.map((c) => (c === '序号' ? i + 1 : (r[c] == null ? '' : r[c]))));
  }

  function exportXlsx() {
    const rows = toMatrix();
    if (!rows.length) throw new Error('还没有数据，先抓列表');
    const bytes = XLSX.build('选题清单', COLS, rows, WIDTHS);
    download(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), '选题清单.xlsx');
    return rows.length;
  }

  function exportCsv() {
    const rows = toMatrix();
    if (!rows.length) throw new Error('还没有数据，先抓列表');
    const esc = (v) => {
      const s = v == null ? '' : String(v);
      return /[",\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
    };
    const text = '\ufeff' + [COLS.join(',')].concat(rows.map((r) => r.map(esc).join(','))).join('\r\n');
    downloadText(text, '选题清单.csv', 'text/csv');
    return rows.length;
  }

  /* ============================== 面板 ============================== */
  let elStatus = null;

  function setStatus(text, color) {
    if (elStatus) {
      elStatus.textContent = text;
      elStatus.style.color = color || '#333';
    }
    console.log('[选题助手]', text);
  }

  function buildPanel() {
    if (document.getElementById('jwh-panel')) return;

    const style = document.createElement('style');
    style.textContent = `
      #jwh-panel{position:fixed;right:16px;bottom:16px;width:250px;z-index:2147483647;
        font:12px/1.5 "Microsoft YaHei",system-ui,sans-serif;background:#fff;color:#222;
        border:1px solid #c9d3e0;border-radius:10px;box-shadow:0 8px 28px rgba(15,40,80,.22);overflow:hidden}
      #jwh-panel .hd{display:flex;align-items:center;justify-content:space-between;
        padding:7px 10px;background:#1f4e79;color:#fff;font-weight:600;cursor:default}
      #jwh-panel .hd small{font-weight:400;opacity:.75;margin-left:6px}
      #jwh-panel .hd button{border:0;background:rgba(255,255,255,.18);color:#fff;width:20px;height:20px;
        border-radius:4px;cursor:pointer;line-height:1}
      #jwh-panel .bd{padding:9px 10px 11px}
      #jwh-panel .st{font-size:11.5px;min-height:32px;margin-bottom:8px;color:#333;word-break:break-all}
      #jwh-panel .btns{display:grid;grid-template-columns:1fr 1fr;gap:6px}
      #jwh-panel button.act{border:1px solid #c9d3e0;background:#f7f9fc;color:#123;border-radius:6px;
        padding:6px 4px;cursor:pointer;font-size:12px}
      #jwh-panel button.act:hover{background:#eef4fb;border-color:#8fb2d9}
      #jwh-panel button.act:disabled{opacity:.45;cursor:not-allowed}
      #jwh-panel button.warn{color:#a33;border-color:#e6c3c3;background:#fdf5f5}
      #jwh-panel .row{display:flex;align-items:center;gap:5px;margin-top:8px;font-size:11.5px;color:#4a5568}
      #jwh-panel .tip{margin-top:7px;font-size:10.5px;color:#8a94a6;line-height:1.45}
      #jwh-panel.collapsed .bd{display:none}
    `;
    document.head.appendChild(style);

    const div = document.createElement('div');
    div.id = 'jwh-panel';
    div.innerHTML = `
      <div class="hd"><span>选题抓取助手<small>v${VERSION}</small></span><button id="jwh-min" title="收起">–</button></div>
      <div class="bd">
        <div class="st" id="jwh-status">就绪：先点「① 抓列表」</div>
        <div class="btns">
          <button class="act" data-act="list">① 抓列表</button>
          <button class="act" data-act="detail">② 抓明细</button>
          <button class="act" data-act="xlsx">③ 导出 Excel</button>
          <button class="act" data-act="csv">导出 CSV</button>
          <button class="act warn" data-act="stop" style="grid-column:1/3">停止</button>
        </div>
        <label class="row"><input type="checkbox" id="jwh-contact"> 导出教师联系方式</label>
        <div class="tip">只在本页运行，复用你的登录状态，不上传任何数据。抓完请自行保存，退出登录后失效。</div>
      </div>`;
    document.body.appendChild(div);

    elStatus = div.querySelector('#jwh-status');
    div.querySelector('#jwh-min').onclick = () => div.classList.toggle('collapsed');
    div.querySelector('#jwh-contact').onchange = (e) => { S.includeContact = e.target.checked; };

    div.querySelectorAll('button[data-act]').forEach((btn) => {
      btn.onclick = async () => {
        const act = btn.dataset.act;
        try {
          if (act === 'list') {
            S.stop = false; S.tids = []; S.detail.clear(); S.listTotal = null;
            setStatus('正在抓列表…');
            const n = await grabList((t) => setStatus(t));
            setStatus(`列表完成：${n} 条` + (S.listTotal ? `（服务端报 ${S.listTotal}）` : ''), '#0a7a3d');
          }
          if (act === 'detail') {
            if (!S.tids.length) {
              setStatus('先抓列表…');
              await grabList((t) => setStatus(t));
            }
            S.stop = false;
            if (!S.template) S.template = learnTemplate();
            setStatus(`开始抓明细：0/${S.tids.length}`);
            const r = await pool(S.tids, CFG.concurrency,
              (x) => grabDetail(x.课题ID).catch(() => { S.detail.set(x.课题ID, {}); }),
              (d, total, fail) => setStatus(`明细 ${d}/${total}${fail ? `（失败 ${fail}）` : ''}`));
            setStatus(`明细完成：${S.detail.size}/${S.tids.length}` + (r.fail ? `，失败 ${r.fail}` : ''), '#0a7a3d');
          }
          if (act === 'xlsx') {
            const n = exportXlsx();
            setStatus(`已导出 Excel：${n} 条`, '#0a7a3d');
          }
          if (act === 'csv') {
            const n = exportCsv();
            setStatus(`已导出 CSV：${n} 条`, '#0a7a3d');
          }
          if (act === 'stop') {
            S.stop = true;
            setStatus('已请求停止…', '#a33');
          }
        } catch (e) {
          setStatus('出错：' + (e && e.message ? e.message : e), '#c00');
        }
      };
    });
  }

  /* ============================== 启动 ============================== */
  const api = {
    version: VERSION,
    async list() { S.tids = []; const n = await grabList((t) => console.log(t)); setStatus(`列表 ${n} 条`); return n; },
    async detail() {
      if (!S.tids.length) await grabList(() => { });
      if (!S.template) S.template = learnTemplate();
      const r = await pool(S.tids, CFG.concurrency, (x) => grabDetail(x.课题ID), (d, t2) => setStatus(`明细 ${d}/${t2}`));
      setStatus(`明细完成 ${S.detail.size}/${S.tids.length}`);
      return r;
    },
    exportXlsx, exportCsv,
    status: () => ({ list: S.tids.length, detail: S.detail.size, template: S.template }),
    stop() { S.stop = true; },
  };
  window.__JWH = api;

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', buildPanel);
  else buildPanel();

  console.log(`%c[选题抓取助手 v${VERSION}] 已加载 · 面板在右下角 · 控制台可用 __JWH`, 'color:#1f4e79;font-weight:bold');
})();