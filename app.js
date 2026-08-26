/* global docx, supabase */
(() => {
  'use strict';

  const APP = { name: '航空城蘭园保洁服务日志', area: '航空城蘭园', taskNames: ['晨会', '', '', '', '夕会', '', '', ''] };
  const $ = (selector) => document.querySelector(selector);
  const state = { date: localDate(new Date()), log: null, taskIndex: 0, previewUrl: null, cardUrls: [] };
  let db;
  let backupInProgress = false;
  let cloudClient = null;
  let cloudReadyPromise = null;

  function localDate(value) {
    const date = new Date(value);
    const offset = date.getTimezoneOffset() * 60000;
    return new Date(date.getTime() - offset).toISOString().slice(0, 10);
  }
  function makeTasks() { return APP.taskNames.map((suggested) => ({ suggested, text: '', image: null, updatedAt: null })); }
  function blankLog(date) { return { date, tasks: makeTasks(), updatedAt: null }; }
  function escapeHtml(value = '') { return String(value).replace(/[&<>'"]/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[char])); }
  function dateParts(dateString) { const [year, month, day] = dateString.split('-').map(Number); return { year, month, day, date: new Date(year, month - 1, day) }; }
  function formatFullDate(dateString) { const { year, month, day } = dateParts(dateString); return `${year}年${month}月${day}日`; }
  function countComplete(log) { return log.tasks.filter((task) => task.image && (task.suggested ? task.text.trim() : true)).length; }
  function isRecorded(log) { return log.tasks.some((task) => task.text.trim() || task.image); }

  function openDatabase() {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open('cleaning-log-pwa', 1);
      request.onupgradeneeded = () => request.result.createObjectStore('logs', { keyPath: 'date' });
      request.onsuccess = () => { db = request.result; resolve(); };
      request.onerror = () => reject(request.error);
    });
  }
  function getLog(date) {
    return new Promise((resolve, reject) => {
      const request = db.transaction('logs', 'readonly').objectStore('logs').get(date);
      request.onsuccess = () => resolve(request.result || blankLog(date));
      request.onerror = () => reject(request.error);
    });
  }
  function putLocal(log) {
    return new Promise((resolve, reject) => {
      const request = db.transaction('logs', 'readwrite').objectStore('logs').put(log);
      request.onsuccess = resolve;
      request.onerror = () => reject(request.error);
    });
  }
  async function putLog(log) {
    const saved = { ...log, updatedAt: new Date().toISOString() };
    await putLocal(saved);
    // Never make a successful local save depend on network availability.  The
    // same record stays on this phone and is retried on the next cloud sync.
    void syncOneLogToCloud(saved).catch((error) => console.warn('Cloud sync will retry later.', error));
  }
  function deleteLocal(date) {
    return new Promise((resolve, reject) => {
      const request = db.transaction('logs', 'readwrite').objectStore('logs').delete(date);
      request.onsuccess = resolve;
      request.onerror = () => reject(request.error);
    });
  }
  function allLogs() {
    return new Promise((resolve, reject) => {
      const request = db.transaction('logs', 'readonly').objectStore('logs').getAll();
      request.onsuccess = () => resolve(request.result.sort((a, b) => a.date.localeCompare(b.date)));
      request.onerror = () => reject(request.error);
    });
  }

  function cloudSettings() {
    const settings = window.CLEANING_LOG_CLOUD || {};
    return { url: String(settings.url || '').replace(/\/$/, ''), anonKey: String(settings.anonKey || '') };
  }
  function cloudConfigured() {
    const { url, anonKey } = cloudSettings();
    return /^https:\/\/[\w-]+\.supabase\.co$/.test(url) && anonKey.length > 30 && !!window.supabase;
  }
  function setCloudInfo(text) {
    const label = $('#cloudInfo');
    if (label) label.textContent = text;
  }
  async function ensureCloudSession() {
    if (!cloudConfigured()) throw new Error('云端资料库还没有设置完成');
    if (cloudClient) return cloudClient;
    if (!cloudReadyPromise) {
      cloudReadyPromise = (async () => {
        const { url, anonKey } = cloudSettings();
        const client = supabase.createClient(url, anonKey, { auth: { persistSession: true, autoRefreshToken: true, detectSessionInUrl: false } });
        const { data: sessionData, error: sessionError } = await client.auth.getSession();
        if (sessionError) throw sessionError;
        if (!sessionData.session) {
          const { error } = await client.auth.signInAnonymously();
          if (error) throw error;
        }
        cloudClient = client;
        return client;
      })();
    }
    return cloudReadyPromise;
  }
  async function logForCloud(log) {
    return {
      date: log.date,
      updatedAt: log.updatedAt || new Date().toISOString(),
      tasks: await Promise.all(log.tasks.map(async (task) => ({
        suggested: task.suggested || '', text: task.text || '', updatedAt: task.updatedAt || null,
        image: task.image instanceof Blob ? await blobToDataUrl(task.image) : (typeof task.image === 'string' ? task.image : null)
      })))
    };
  }
  async function syncOneLogToCloud(log) {
    if (!cloudConfigured() || !isRecorded(log)) return false;
    const client = await ensureCloudSession();
    const payload = await logForCloud(log);
    const { error } = await client.from('cleaning_logs').upsert({ date: payload.date, payload }, { onConflict: 'device_id,date' });
    if (error) throw error;
    localStorage.setItem('cleaning-log-cloud-synced-at', payload.updatedAt);
    setCloudInfo(`云端已保存：${formatFullDate(payload.date)}`);
    return true;
  }
  async function remoteLogToLocal(payload) {
    if (!payload || typeof payload.date !== 'string' || !Array.isArray(payload.tasks) || payload.tasks.length !== 8) return null;
    return {
      date: payload.date,
      updatedAt: payload.updatedAt || null,
      tasks: payload.tasks.map((task, index) => ({
        suggested: task.suggested || APP.taskNames[index] || '', text: task.text || '', updatedAt: task.updatedAt || null,
        image: typeof task.image === 'string' && task.image.startsWith('data:image/') ? dataUrlToBlob(task.image) : null
      }))
    };
  }
  async function restoreCloudLogs() {
    if (!cloudConfigured()) { setCloudInfo('云端尚未设置：本机保存仍可正常使用'); return 0; }
    const client = await ensureCloudSession();
    const { data, error } = await client.from('cleaning_logs').select('payload').order('updated_at', { ascending: true });
    if (error) throw error;
    let restored = 0;
    for (const row of data || []) {
      const remote = await remoteLogToLocal(row.payload);
      if (!remote) continue;
      const local = await getLog(remote.date);
      if (!isRecorded(local) || String(remote.updatedAt || '') > String(local.updatedAt || '')) { await putLocal(remote); restored += 1; }
    }
    setCloudInfo(restored ? `已从云端恢复 ${restored} 天记录` : '云端连接正常，手机记录会自动保存');
    return restored;
  }
  async function syncAllToCloud() {
    if (!cloudConfigured()) throw new Error('云端还没有设置完成。请先完成电脑上的 Supabase 设置。');
    const logs = (await allLogs()).filter(isRecorded);
    if (!logs.length) { showMessage('还没有可同步的记录', '先保存至少一项文字或照片。', '＋'); return; }
    showBusy('正在保存到云端', `正在上传 0 / ${logs.length} 天记录`);
    let done = 0;
    for (const log of logs) { await syncOneLogToCloud(log); done += 1; busyProgress(Math.round(done / logs.length * 100), `已保存 ${done} / ${logs.length} 天记录`); }
    // The user explicitly chose a phone-as-inbox workflow.  Clear only after
    // every selected day has acknowledged a successful cloud upload; a failed
    // upload throws above and leaves all phone data intact for retry.
    busyProgress(96, '云端已确认，正在清空这台手机的已备份记录');
    for (const log of logs) await deleteLocal(log.date);
    await showDate(state.date);
    hideBusy(); $('#backupDialog').close();
    showMessage('云端已保存，手机已清空', `${logs.length} 天记录和照片已成功上传云端，已从这台手机清除。电脑开机后双击同步文件即可写入 Word。`, '✓');
  }

  async function deleteLog(date) {
    await deleteLocal(date);
    if (cloudConfigured()) {
      try {
        const client = await ensureCloudSession();
        const { error } = await client.from('cleaning_logs').delete().eq('date', date);
        if (error) throw error;
      } catch (error) { console.warn('Cloud deletion will be retried manually.', error); }
    }
  }
  function photoBackupEnabled() { return localStorage.getItem('cleaning-log-photo-backup') !== 'off'; }
  async function backupPhotosToComputer(log) {
    const photos = await Promise.all(log.tasks.map(async (task, index) => task.image instanceof Blob ? ({ index, image: await blobToDataUrl(task.image) }) : null));
    const response = await fetch('/api/photo-backups', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ date: log.date, photos: photos.filter(Boolean) }) });
    if (!response.ok) throw new Error('电脑没有接收照片备份');
  }

  function updateDateHeader() {
    const { year, month, day, date } = dateParts(state.date);
    $('#weekdayText').textContent = ['星期日', '星期一', '星期二', '星期三', '星期四', '星期五', '星期六'][date.getDay()];
    $('#dateText').textContent = `${month}月${day}日`;
    $('#yearText').textContent = `${year}年`;
    $('#dateInput').value = state.date;
  }
  function revokeCardUrls() { state.cardUrls.forEach((url) => URL.revokeObjectURL(url)); state.cardUrls = []; }
  function renderTasks() {
    revokeCardUrls();
    const completed = countComplete(state.log);
    $('#completedCount').textContent = completed;
    $('#progressBar').style.width = `${completed * 12.5}%`;
    $('#progressMessage').textContent = completed === 0 ? '先从第一项开始吧' : completed === 8 ? '今天的八项工作已完整记录' : `还差 ${8 - completed} 项，拍好照片就完成了`;
    $('#taskGrid').innerHTML = state.log.tasks.map((task, index) => {
      const done = task.image && (task.suggested ? task.text.trim() : true);
      const displayText = task.text.trim() || task.suggested || `第 ${index + 1} 项现场照片`;
      let photo = '<span class="photo-empty">＋</span>';
      if (task.image) { const url = URL.createObjectURL(task.image); state.cardUrls.push(url); photo = `<img src="${url}" alt="${escapeHtml(displayText)}" />`; }
      return `<button class="task-card ${done ? 'complete' : ''}" data-task-index="${index}"><span class="photo-tile">${photo}<span class="badge">${index + 1}</span>${done ? '<span class="done-check">✓</span>' : ''}</span><span class="task-copy"><strong>${escapeHtml(displayText)}</strong><small>${task.image ? (done ? '已记录' : '照片已选') : '点击填写并拍照'}</small></span></button>`;
    }).join('');
  }
  async function updateExportLabels() {
    const logs = await allLogs();
    const { year, month } = dateParts(state.date);
    const monthDays = logs.filter((log) => log.date.startsWith(`${year}-${String(month).padStart(2, '0')}`) && isRecorded(log)).length;
    const yearDays = logs.filter((log) => log.date.startsWith(`${year}-`) && isRecorded(log)).length;
    const exportText = $('#masterExportText');
    if (exportText) exportText.textContent = `${year}年已有 ${yearDays} 天记录，备份后可下载`;
  }
  async function showDate(date) {
    state.date = date;
    state.log = await getLog(date);
    updateDateHeader(); renderTasks(); await updateExportLabels();
  }
  let calendarCursor = null;
  function renderCalendar() {
    const cursor = calendarCursor || dateParts(state.date).date;
    const year = cursor.getFullYear(), month = cursor.getMonth();
    $('#calendarTitle').textContent = `${year}年${month + 1}月`;
    const labels = ['日', '一', '二', '三', '四', '五', '六'];
    const first = new Date(year, month, 1).getDay(); const days = new Date(year, month + 1, 0).getDate();
    const cells = labels.map((label) => `<span style="text-align:center;color:#6d827c;font-size:12px;padding:6px 0">${label}</span>`);
    for (let index = 0; index < first; index++) cells.push('<span></span>');
    for (let day = 1; day <= days; day++) { const value = `${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`; const chosen = value === state.date; cells.push(`<button type="button" data-calendar-date="${value}" style="border:0;border-radius:10px;padding:10px 0;background:${chosen ? '#0f766e' : '#eef5f1'};color:${chosen ? '#fff' : '#153d36'};font-weight:700">${day}</button>`); }
    $('#calendarGrid').innerHTML = cells.join('');
  }
  function openCalendar() { calendarCursor = dateParts(state.date).date; renderCalendar(); $('#dateDialog').showModal(); }
  async function normaliseExistingTasks() {
    const logs = await allLogs();
    for (const log of logs) {
      let changed = false;
      log.tasks.forEach((task, index) => {
        const oldSuggested = task.suggested || '';
        const nextSuggested = APP.taskNames[index];
        if (task.text === oldSuggested) { task.text = ''; changed = true; }
        if (task.suggested !== nextSuggested) { task.suggested = nextSuggested; changed = true; }
      });
      if (changed) await putLocal(log);
    }
  }

  function showMessage(title, body, symbol = '!', options = {}) {
    $('#messageTitle').textContent = title; $('#messageBody').textContent = body; $('#messageSymbol').textContent = symbol;
    const actionArea = $('#messageActions'); actionArea.innerHTML = '';
    if (options.cancel) { const cancel = document.createElement('button'); cancel.className = 'delete-button'; cancel.textContent = options.cancelText || '取消'; cancel.onclick = () => $('#messageDialog').close(); actionArea.append(cancel); }
    const confirm = document.createElement('button'); confirm.className = 'save-button'; confirm.textContent = options.confirmText || '知道了'; confirm.onclick = async () => { $('#messageDialog').close(); if (options.onConfirm) await options.onConfirm(); }; actionArea.append(confirm);
    $('#messageDialog').showModal();
  }
  function showBusy(title, text) { $('#busyTitle').textContent = title; $('#busyText').textContent = text; $('#busyProgress').style.width = '0%'; $('#busyLayer').hidden = false; }
  function busyProgress(percent, text) { $('#busyProgress').style.width = `${percent}%`; if (text) $('#busyText').textContent = text; }
  function hideBusy() { $('#busyLayer').hidden = true; }
  function errorMessage(error, context) {
    if (error?.name === 'QuotaExceededError') return ['手机空间不够', '照片保存在本机。请先导出备份，再清理不需要的记录或手机空间。'];
    if (error?.name === 'NotAllowedError') return ['无法读取照片', '请允许浏览器访问相机或相册后再试一次。'];
    return [context || '操作没有完成', '数据没有丢失。请稍后重试；如果仍失败，请先导出备份。'];
  }
  async function safely(work, context) { try { await work(); } catch (error) { console.error(error); hideBusy(); const [title, body] = errorMessage(error, context); showMessage(title, body, '×'); } }

  function resetPreview() { if (state.previewUrl) URL.revokeObjectURL(state.previewUrl); state.previewUrl = null; $('#photoPreview').innerHTML = '<span>还没有照片</span>'; }
  function showPreview(blob) { resetPreview(); if (!blob) return; state.previewUrl = URL.createObjectURL(blob); $('#photoPreview').innerHTML = `<img src="${state.previewUrl}" alt="待保存的现场照片" />`; }
  function openTask(index) {
    state.taskIndex = index;
    const task = state.log.tasks[index];
    $('#taskOrdinal').textContent = `第 ${index + 1} 项`;
    $('#taskDialogTitle').textContent = task.text || task.suggested || `第 ${index + 1} 项现场照片`;
    $('#taskTextInput').value = task.text || '';
    $('#charCount').textContent = $('#taskTextInput').value.length;
    showPreview(task.image); $('#taskDialog').showModal();
  }
  async function compressImage(file) {
    if (!file?.type?.startsWith('image/')) throw new Error('请选择照片文件');
    const sourceUrl = URL.createObjectURL(file);
    try {
      const image = await new Promise((resolve, reject) => { const img = new Image(); img.onload = () => resolve(img); img.onerror = reject; img.src = sourceUrl; });
      const targetW = 640, targetH = 480, targetRatio = targetW / targetH;
      const sourceRatio = image.naturalWidth / image.naturalHeight;
      let sx = 0, sy = 0, sw = image.naturalWidth, sh = image.naturalHeight;
      if (sourceRatio > targetRatio) { sw = Math.round(sh * targetRatio); sx = Math.round((image.naturalWidth - sw) / 2); } else { sh = Math.round(sw / targetRatio); sy = Math.round((image.naturalHeight - sh) / 2); }
      const canvas = document.createElement('canvas'); canvas.width = targetW; canvas.height = targetH;
      canvas.getContext('2d', { alpha: false }).drawImage(image, sx, sy, sw, sh, 0, 0, targetW, targetH);
      return await new Promise((resolve, reject) => canvas.toBlob((blob) => blob ? resolve(blob) : reject(new Error('图片压缩失败')), 'image/jpeg', .76));
    } finally { URL.revokeObjectURL(sourceUrl); }
  }
  async function handlePhoto(file) { const compressed = await compressImage(file); state.log.tasks[state.taskIndex].image = compressed; showPreview(compressed); }
  async function saveTask(event) {
    event.preventDefault();
    const task = state.log.tasks[state.taskIndex]; const text = $('#taskTextInput').value.trim();
    task.text = text; task.updatedAt = new Date().toISOString();
    await putLog(state.log); $('#taskDialog').close(); await showDate(state.date);
  }

  function twipsFromCm(cm) { return Math.round((cm / 2.54) * 1440); }
  function makeRun(text, options = {}) { return new docx.TextRun({ text, font: '宋体', size: options.size || 20, bold: options.bold || false, color: options.color || '000000', break: options.break || undefined }); }
  function textCell(text, options = {}) {
    return new docx.TableCell({ width: { size: options.width || 1660, type: docx.WidthType.DXA }, verticalAlign: docx.VerticalAlign.CENTER, rowSpan: options.rowSpan, margins: { top: 0, bottom: 0, left: 40, right: 40 }, children: [new docx.Paragraph({ alignment: docx.AlignmentType.CENTER, spacing: { before: 0, after: 0, line: 240 }, children: text ? [makeRun(text, { size: options.size || 21, color: '000000' })] : [] })] });
  }
  async function imageCell(task) {
    const children = [];
    if (task.image) {
      children.push(new docx.Paragraph({ alignment: docx.AlignmentType.CENTER, spacing: { before: 0, after: 0 }, children: [new docx.ImageRun({ data: await task.image.arrayBuffer(), type: 'jpg', transformation: { width: 86, height: 65 }, altText: { title: task.text || task.suggested, description: '保洁工作现场照片', name: '现场照片' } })] }));
    } else children.push(new docx.Paragraph({ alignment: docx.AlignmentType.CENTER, spacing: { before: 0, after: 0 }, children: [makeRun('未上传照片', { size: 16, color: 'A6A6A6' })] }));
    return new docx.TableCell({ width: { size: 1660, type: docx.WidthType.DXA }, verticalAlign: docx.VerticalAlign.CENTER, margins: { top: 65, bottom: 65, left: 65, right: 65 }, children });
  }
  function dateCell(month, day, continuation = false) {
    return new docx.TableCell({ width: { size: 1665, type: docx.WidthType.DXA }, verticalAlign: docx.VerticalAlign.CENTER, verticalMerge: continuation ? docx.VerticalMergeType.CONTINUE : docx.VerticalMergeType.RESTART, margins: { top: 70, bottom: 70, left: 70, right: 70 }, children: [new docx.Paragraph({ alignment: docx.AlignmentType.CENTER, spacing: { before: 0, after: 0, line: 270 }, children: continuation ? [] : [makeRun(`${month}月`, { size: 25 }), new docx.TextRun({ break: 1 }), makeRun(`${day}日`, { size: 25 })] })] });
  }
  async function dayRows(log, progress) {
    const { month, day } = dateParts(log.date); const tasks = log.tasks;
    const content = (start, continuation) => [dateCell(month, day, continuation), ...tasks.slice(start, start + 4).map((task) => textCell(task.text || task.suggested))];
    const photos = async (start) => [dateCell(month, day, true), ...(await Promise.all(tasks.slice(start, start + 4).map(imageCell)))];
    const firstPhotos = await photos(0); const lastPhotos = await photos(4);
    progress();
    return [new docx.TableRow({ children: content(0, false), cantSplit: true }), new docx.TableRow({ children: firstPhotos, cantSplit: true }), new docx.TableRow({ children: content(4, true), cantSplit: true }), new docx.TableRow({ children: lastPhotos, cantSplit: true })];
  }
  function monthHeader(logs, year, month, pageBreakBefore) {
    const days = new Date(year, month, 0).getDate();
    return [new docx.Paragraph({ pageBreakBefore, alignment: docx.AlignmentType.CENTER, spacing: { after: 120 }, children: [makeRun(APP.name, { size: 30, bold: true })] }), new docx.Paragraph({ alignment: docx.AlignmentType.LEFT, spacing: { after: 50 }, children: [makeRun(`服务日期：${year}年${month}月1日—${month}月${days}日`, { size: 20 })] }), new docx.Paragraph({ alignment: docx.AlignmentType.LEFT, spacing: { after: 140 }, children: [makeRun(`服务区域：${APP.area}`, { size: 20 })] })];
  }
  async function fetchWithTimeout(url, options = {}, timeout = 180000) {
    const controller = new AbortController(); const timer = setTimeout(() => controller.abort(), timeout);
    try { return await fetch(url, { ...options, signal: controller.signal }); }
    finally { clearTimeout(timer); }
  }
  async function phoneLogsForComputer(logs, onProgress) {
    let done = 0;
    return Promise.all(logs.map(async (log) => {
      const tasks = await Promise.all(log.tasks.map(async (task) => ({ ...task, image: task.image instanceof Blob ? await blobToDataUrl(task.image) : null })));
      done += 1; onProgress(done); return { ...log, tasks };
    }));
  }
  async function exportWord() {
    showBusy('正在准备手机下载', '正在确认电脑 Word 已备份');
    try {
      const response = await fetchWithTimeout('/api/master-status', {}, 20000);
      const status = await response.json().catch(() => ({}));
      if (!response.ok || !status.ready) throw new Error('请先点击底部“备份”，完成后再下载 Word');
      hideBusy(); window.location.assign(`/api/download-master?time=${Date.now()}`);
    } catch (error) {
      hideBusy(); showMessage('暂时不能下载', error.message || '请先完成备份。', '×');
    }
  }
  async function buildWordLocal(logs, scope, year, month) {
    showBusy(scope === 'year' ? '正在生成全年 Word' : '正在生成整月 Word', `正在整理 0 / ${logs.length} 天记录`);
    await new Promise((resolve) => setTimeout(resolve, 80));
    const children = []; const groups = new Map(); logs.forEach((log) => { const key = log.date.slice(0, 7); if (!groups.has(key)) groups.set(key, []); groups.get(key).push(log); });
    let done = 0;
    for (const [key, group] of groups) {
      const [groupYear, groupMonth] = key.split('-').map(Number); children.push(...monthHeader(group, groupYear, groupMonth, children.length > 0));
      const rows = [];
      for (const log of group) { rows.push(...await dayRows(log, () => { done += 1; busyProgress(Math.round((done / logs.length) * 76), `正在整理 ${done} / ${logs.length} 天记录`); })); }
      children.push(new docx.Table({ rows, width: { size: 100, type: docx.WidthType.PERCENTAGE }, columnWidths: [1665, 1660, 1660, 1660, 1660], layout: docx.TableLayoutType.FIXED, borders: { top: { style: docx.BorderStyle.SINGLE, size: 4, color: '000000' }, bottom: { style: docx.BorderStyle.SINGLE, size: 4, color: '000000' }, left: { style: docx.BorderStyle.SINGLE, size: 4, color: '000000' }, right: { style: docx.BorderStyle.SINGLE, size: 4, color: '000000' }, insideHorizontal: { style: docx.BorderStyle.SINGLE, size: 4, color: '000000' }, insideVertical: { style: docx.BorderStyle.SINGLE, size: 4, color: '000000' } } }));
    }
    busyProgress(82, '正在写入 Word 文件');
    const documentFile = new docx.Document({ creator: '航空城蘭园保洁日志', title: APP.name, sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, bottom: 1440, left: 1800, right: 1800 } } }, children }] });
    const blob = await docx.Packer.toBlob(documentFile); busyProgress(96, '正在准备下载');
    const label = scope === 'year' ? `${year}年` : `${year}年${month}月`;
    downloadBlob(blob, `${APP.name}_${label}.docx`); hideBusy(); showMessage('Word 已生成', `${label}日志已保存到手机下载文件夹。现在可以从“文件”里直接微信转发。`, '✓');
  }
  function downloadBlob(blob, filename) { const link = document.createElement('a'); const url = URL.createObjectURL(blob); link.href = url; link.download = filename; document.body.append(link); link.click(); link.remove(); setTimeout(() => URL.revokeObjectURL(url), 10000); }

  async function showArchive() {
    const logs = (await allLogs()).filter(isRecorded).reverse();
    $('#archiveList').innerHTML = logs.length ? logs.map((log) => `<button class="archive-row" data-open-date="${log.date}"><span class="archive-dot ${countComplete(log) === 8 ? 'complete' : ''}"></span><span><strong>${formatFullDate(log.date)}</strong><small>已完成 ${countComplete(log)} / 8 项</small></span></button>`).join('') : '<div class="dialog-copy">还没有历史记录。每天填写完成后，记录会自动出现在这里。</div>';
    $('#archiveDialog').showModal();
  }
  function blobToDataUrl(blob) { return new Promise((resolve, reject) => { const reader = new FileReader(); reader.onload = () => resolve(reader.result); reader.onerror = reject; reader.readAsDataURL(blob); }); }
  function dataUrlToBlob(dataUrl) { const [header, data] = dataUrl.split(','); const mime = /data:(.*?);/.exec(header)?.[1] || 'image/jpeg'; const binary = atob(data); const bytes = new Uint8Array(binary.length); for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i); return new Blob([bytes], { type: mime }); }
  async function makeBackupPayload(logs, title = '正在制作备份') {
    showBusy(title, `正在整理 0 / ${logs.length} 天记录`);
    const prepared = [];
    for (let index = 0; index < logs.length; index++) { const log = logs[index]; prepared.push({ ...log, tasks: await Promise.all(log.tasks.map(async (task) => ({ ...task, image: task.image ? await blobToDataUrl(task.image) : null }))) }); busyProgress(Math.round(((index + 1) / Math.max(logs.length, 1)) * 90), `正在整理 ${index + 1} / ${logs.length} 天记录`); }
    return { version: 2, app: 'cleaning-log-onedrive', exportedAt: new Date().toISOString(), logs: prepared };
  }
  function timestampForFilename() { return new Date().toISOString().replace(/[-:]/g, '').replace(/\.\d{3}Z$/, ''); }
  async function exportBackup() {
    const logs = (await allLogs()).filter(isRecorded);
    if (!logs.length) { showMessage('还没有可备份的记录', '请先保存至少一项文字或照片。', '＋'); return; }
    const payload = await makeBackupPayload(logs);
    downloadBlob(new Blob([JSON.stringify(payload)], { type: 'application/json' }), `保洁日志备份_${localDate(new Date())}.json`); hideBusy(); showMessage('备份已导出', '请把备份文件存到微信收藏、OneDrive 或电脑。', '✓');
  }
  async function shareBackupToOneDrive() {
    const all = (await allLogs()).filter(isRecorded);
    if (!all.length) { showMessage('还没有可同步的记录', '请先保存至少一项文字或照片。', '＋'); return; }
    const lastSynced = localStorage.getItem('cleaning-log-onedrive-synced-at') || '';
    const changed = all.filter((log) => !lastSynced || (log.updatedAt || '') > lastSynced);
    if (!changed.length) { showMessage('没有新的记录', '本机记录已同步过。新增或修改内容后再点同步即可。', '✓'); return; }
    $('#backupDialog').close();
    const payload = await makeBackupPayload(changed, '正在准备 OneDrive 同步');
    const file = new File([JSON.stringify(payload)], `保洁日志同步_${localDate(new Date())}_${timestampForFilename()}.ready.json`, { type: 'application/json' });
    hideBusy();
    const shareData = { files: [file], title: '同步保洁日志到 OneDrive', text: '请选择 OneDrive，并保存到“保洁日志同步”文件夹。' };
    if (navigator.share && (!navigator.canShare || navigator.canShare(shareData))) {
      try {
        await navigator.share(shareData);
        localStorage.setItem('cleaning-log-onedrive-synced-at', payload.exportedAt);
        showMessage('请在 OneDrive 中确认保存', '已打开手机分享面板：请选择“OneDrive”，保存位置选“保洁日志同步”。电脑下次开机后会自动写入原 Word。', '☁');
        return;
      } catch (error) {
        // Android browsers sometimes close the share panel without handing the
        // file to OneDrive.  Always leave a usable file in Downloads instead
        // of presenting a dead-end “cancelled” error.
        console.warn('OneDrive share sheet did not finish; downloading file instead.', error);
        downloadBlob(file, file.name);
        showMessage('同步文件已下载', '手机没有完成 OneDrive 分享，所以同步文件已自动保存到“下载”。请打开 OneDrive，点“+”上传，选择这个文件并保存到“保洁日志同步”文件夹。电脑随后会自动更新 Word。', '☁');
        return;
      }
    }
    downloadBlob(file, file.name);
    showMessage('同步文件已下载', '请在“文件/下载”里点分享，选择 OneDrive，并保存到“保洁日志同步”文件夹。', '☁');
  }
  async function importBackup(file) {
    if (!file) return; const payload = JSON.parse(await file.text()); if (!Array.isArray(payload.logs)) throw new Error('不是保洁日志备份文件');
    showBusy('正在恢复备份', `正在恢复 0 / ${payload.logs.length} 天记录`);
    for (let index = 0; index < payload.logs.length; index++) { const log = payload.logs[index]; log.tasks = log.tasks.map((task) => ({ ...task, image: task.image ? dataUrlToBlob(task.image) : null })); await putLog(log); busyProgress(Math.round(((index + 1) / Math.max(payload.logs.length, 1)) * 90), `正在恢复 ${index + 1} / ${payload.logs.length} 天记录`); }
    hideBusy(); $('#backupDialog').close(); await showDate(state.date); showMessage('备份已恢复', `${payload.logs.length} 天记录已保存到这台手机。`, '✓');
  }
  async function updateStorageInfo() { if (!navigator.storage?.estimate) return; const { usage = 0, quota = 0 } = await navigator.storage.estimate(); $('#storageInfo').textContent = `本机已使用 ${(usage / 1024 / 1024).toFixed(1)} MB / 可用 ${(quota / 1024 / 1024 / 1024).toFixed(1)} GB`; }

  function bindEvents() {
    $('#previousDay').onclick = () => safely(() => showDate(localDate(new Date(`${state.date}T12:00:00`).getTime() - 86400000)), '无法切换日期');
    $('#nextDay').onclick = () => safely(() => showDate(localDate(new Date(`${state.date}T12:00:00`).getTime() + 86400000)), '无法切换日期');
    $('#datePickerButton').onclick = openCalendar;
    $('#dateInput').onchange = (event) => safely(() => showDate(event.target.value), '无法读取这一天的记录');
    $('#closeCalendar').onclick = () => $('#dateDialog').close();
    $('#calendarPrev').onclick = () => { calendarCursor = new Date(calendarCursor.getFullYear(), calendarCursor.getMonth() - 1, 1); renderCalendar(); };
    $('#calendarNext').onclick = () => { calendarCursor = new Date(calendarCursor.getFullYear(), calendarCursor.getMonth() + 1, 1); renderCalendar(); };
    $('#calendarGrid').onclick = (event) => { const button = event.target.closest('[data-calendar-date]'); if (button) { $('#dateDialog').close(); safely(() => showDate(button.dataset.calendarDate), '无法读取这一天的记录'); } };
    $('#taskGrid').onclick = (event) => { const card = event.target.closest('[data-task-index]'); if (card) openTask(Number(card.dataset.taskIndex)); };
    $('#taskTextInput').oninput = (event) => { $('#charCount').textContent = event.target.value.length; $('#taskDialogTitle').textContent = event.target.value.trim() || state.log.tasks[state.taskIndex].suggested || `第 ${state.taskIndex + 1} 项现场照片`; };
    $('#cameraInput').onchange = (event) => safely(async () => { await handlePhoto(event.target.files[0]); event.target.value = ''; }, '照片没有处理成功');
    $('#galleryInput').onchange = (event) => safely(async () => { await handlePhoto(event.target.files[0]); event.target.value = ''; }, '照片没有处理成功');
    $('#taskForm').onsubmit = (event) => safely(() => saveTask(event), '这一项没有保存成功');
    $('#deleteTaskPhoto').onclick = () => { const task = state.log.tasks[state.taskIndex]; task.image = null; task.text = ''; $('#taskTextInput').value = ''; $('#charCount').textContent = '0'; showPreview(null); };
    $('#taskDialog').addEventListener('close', resetPreview);
    $('#clearDayButton').onclick = () => showMessage('清空当天记录？', `${formatFullDate(state.date)}的文字和照片都会从这台手机移除。`, '！', { cancel: true, cancelText: '不清空', confirmText: '清空当天', onConfirm: async () => { await deleteLog(state.date); await showDate(state.date); } });
    document.querySelectorAll('[data-export]').forEach((button) => button.onclick = () => safely(exportWord, 'Word 没有生成成功'));
    document.querySelectorAll('.nav-item').forEach((button) => button.onclick = () => { if (button.dataset.view === 'archive') safely(showArchive, '无法读取历史记录'); if (button.dataset.view === 'backup') safely(async () => { await updateStorageInfo(); if (!cloudConfigured()) setCloudInfo('云端尚未设置：当前仍只保存在本机'); $('#backupDialog').showModal(); }, '无法读取手机存储空间'); });
    $('#closeArchive').onclick = () => $('#archiveDialog').close(); $('#closeBackup').onclick = () => $('#backupDialog').close();
    $('#archiveList').onclick = (event) => { const row = event.target.closest('[data-open-date]'); if (row) { $('#archiveDialog').close(); safely(() => showDate(row.dataset.openDate), '无法打开这一天'); } };
    $('#backupExport').onclick = () => safely(exportBackup, '备份没有导出成功'); $('#backupCloud').onclick = () => safely(syncAllToCloud, '云端同步没有完成'); $('#backupImport').onchange = (event) => safely(() => importBackup(event.target.files[0]), '备份没有导入成功');
    $('#moreButton').onclick = () => showMessage('保洁日志', '保存当天记录后会自动上传云端。点底部“备份”里的“同步云端并清空手机记录”后，已成功上传的数据会从手机清除；电脑开机后双击同步文件即可写入 Word。', 'i');
  }
  async function init() {
    await openDatabase(); await navigator.storage?.persist?.(); await normaliseExistingTasks();
    try { await restoreCloudLogs(); } catch (error) { console.warn('Cloud restore unavailable.', error); setCloudInfo('暂时无法连接云端：本机记录没有丢失，稍后会重试'); }
    bindEvents(); await showDate(state.date);
    if ('serviceWorker' in navigator) navigator.serviceWorker.register('./sw.js?v=17').catch((error) => console.warn('Service worker unavailable', error));
  }
  safely(init, '应用没有启动成功');
})();
