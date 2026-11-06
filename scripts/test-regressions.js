// Node 22+；使用本机 Chrome / Edge 的无头模式，不依赖 npm 测试库。
// TICKETCHAN_TEST_BROWSER 可指定 Chromium 可执行文件。
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { pathToFileURL } = require('node:url');
const { spawn } = require('node:child_process');

const root = path.resolve(__dirname, '..');
const candidates = [process.env.TICKETCHAN_TEST_BROWSER,
  'C:/Program Files/Google/Chrome/Application/chrome.exe',
  'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe'];
const executable = candidates.find(p => p && fs.existsSync(p));
if (!executable) throw new Error('找不到 Chrome / Edge，请设置 TICKETCHAN_TEST_BROWSER');

const profile = fs.mkdtempSync(path.join(os.tmpdir(), 'ticketchan-regression-'));
const browser = spawn(executable, ['--headless=new', '--disable-gpu', '--no-first-run',
  '--no-default-browser-check', '--remote-debugging-port=0', '--user-data-dir=' + profile,
  pathToFileURL(path.join(root, 'src/index.html')).href], { windowsHide: true, stdio: 'ignore' });
let socket;
let checks = 0;
const delay = ms => new Promise(resolve => setTimeout(resolve, ms));

async function run() {
  let port;
  for (let i = 0; i < 100; i++) {
    try { port = Number(fs.readFileSync(path.join(profile, 'DevToolsActivePort'), 'utf8').split('\n')[0]); } catch {}
    if (port) break;
    await delay(100);
  }
  assert.ok(port, '浏览器调试端口未启动');
  const tabs = await (await fetch('http://127.0.0.1:' + port + '/json')).json();
  socket = new WebSocket(tabs.find(t => t.type === 'page').webSocketDebuggerUrl);
  await new Promise((resolve, reject) => { socket.onopen = resolve; socket.onerror = reject; });
  let sequence = 0;
  const pending = new Map();
  socket.onmessage = event => {
    const message = JSON.parse(event.data);
    if (!message.id) return;
    const item = pending.get(message.id);
    pending.delete(message.id);
    message.error ? item.reject(message.error) : item.resolve(message.result);
  };
  function cdp(method, params = {}) {
    return new Promise((resolve, reject) => {
      const id = ++sequence;
      pending.set(id, { resolve, reject });
      socket.send(JSON.stringify({ id, method, params }));
    });
  }
  async function evaluate(fn, ...args) {
    const result = await cdp('Runtime.evaluate', {
      expression: '(' + fn.toString() + ')(...' + JSON.stringify(args) + ')',
      returnByValue: true, awaitPromise: true,
    });
    if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text);
    return result.result.value;
  }
  function equal(actual, expected, label) { assert.deepEqual(actual, expected, label); checks++; }
  await cdp('Emulation.setDeviceMetricsOverride', { width: 1280, height: 800, deviceScaleFactor: 1, mobile: false });
  for (let i = 0; i < 100; i++) {
    if (await evaluate(() => document.readyState === 'complete' && typeof getSettings === 'function')) break;
    await delay(100);
  }

  // 真实样本原文与人工核对值放在仓库外；通过环境变量显式启用。
  if (process.env.TICKETCHAN_SAMPLE_INPUT || process.env.TICKETCHAN_SAMPLE_EXPECTED) {
    assert.ok(process.env.TICKETCHAN_SAMPLE_INPUT && process.env.TICKETCHAN_SAMPLE_EXPECTED, '样本输入与预期文件须同时指定');
    const inputs = JSON.parse(fs.readFileSync(process.env.TICKETCHAN_SAMPLE_INPUT, 'utf8'));
    const expected = JSON.parse(fs.readFileSync(process.env.TICKETCHAN_SAMPLE_EXPECTED, 'utf8'));
    const results = await evaluate(inputs => inputs.map(c => {
      var info = c.input ? extractByCoordinates(c.input) : c.info;
      var f = createFileObj({ name: c.name, type: c.kind === 'ocr' ? 'jpg' : c.kind,
        ow: c.input?.imgW || 600, oh: c.input?.imgH || 400 });
      if (c.kind === 'pdf') applyPdfTextResult(f, c.input);
      else if (c.kind === 'ocr') applyOcrResult(f, c.input);
      else Object.assign(f, info, { invoiceRemark: info.remark || '', _isToll: !!info.isToll, _xmlInvoice: c.kind === 'xml' });
      var extracted = Object.assign({}, info, { invoiceRemark: info.invoiceRemark || info.remark || '' });
      return { name: c.name, page: c.page, extracted: extracted, applied: Object.assign({}, f, { invoiceType: resolveInvoiceType(f) }) };
    }), inputs);
    equal(results.length, expected.length, '真实样本数量');
    let sampleChecks = 0;
    for (const c of results) {
      const golden = expected.find(e => e.name === c.name && e.page === c.page);
      assert.ok(golden, '未提供样本预期: ' + c.name + '/' + c.page);
      for (const [field, value] of Object.entries(golden.fields)) {
        equal(c.applied[field], value, c.name + '/' + c.page + '/写入/' + field);
        sampleChecks++;
        if (field !== 'invoiceType') {
          equal(c.extracted[field], value, c.name + '/' + c.page + '/提取/' + field);
          sampleChecks++;
        }
      }
    }
    if (process.env.TICKETCHAN_SAMPLE_REPORT) fs.writeFileSync(process.env.TICKETCHAN_SAMPLE_REPORT, JSON.stringify(results, null, 2));
    console.log('PASS: ' + inputs.length + ' real sample pages/results, ' + sampleChecks + ' field checks');
  }

  const names = await evaluate(() => {
    function invoice(rows) {
      return { imgW: 1000, imgH: 600, text: rows.map(r => r[0]).join('\n'),
        lines: rows.map(r => ({ confidence: 1, words: [{ text: r[0], x: r[1], y: r[2], w: r[3], h: 12 }] })) };
    }
    var legacy = invoice([['增值税普通发票', 400, 30, 250],
      ['购买方名称：购方测试有限公司', 110, 110, 300], ['91310000123456789A', 110, 140, 200],
      ['价税合计', 30, 400, 100], ['备注', 530, 485, 30], ['销方测试有限公司', 110, 485, 200],
      ['91310000987654321B', 110, 515, 200], ['销方开户银行：银行测试有限公司', 600, 485, 350]]);
    var personal = invoice([['电子发票（普通发票）', 350, 30, 300],
      ['购买方信息', 30, 100, 80], ['销售方信息', 530, 100, 80], ['名称', 100, 100, 30],
      ['购方测试有限公司', 210, 100, 200], ['名称', 620, 100, 30], ['张三', 660, 100, 50],
      ['统一社会信用代码', 100, 145, 100], ['91310000123456789A', 210, 145, 200],
      ['纳税人识别号', 620, 145, 100], ['310101199001010019', 760, 145, 200], ['价税合计', 30, 400, 100]]);
    var missing = invoice([['电子发票（普通发票）', 350, 30, 300],
      ['购买方名称：购方测试有限公司', 100, 100, 310], ['91310000987654321B', 760, 145, 200]]);
    var results = [legacy, personal, missing].map(extractByCoordinates);
    return results.map(r => ({ buyer: r.buyerName, seller: r.sellerName }));
  });
  equal(names[0], { buyer: '购方测试有限公司', seller: '销方测试有限公司' }, '旧票下部销方与右侧备注银行');
  equal(names[1], { buyer: '购方测试有限公司', seller: '张三' }, '保留个人销售方');
  equal(names[2], { buyer: '购方测试有限公司', seller: '' }, '缺失销方不能拿购方填充');

  const remarks = await evaluate(() => {
    function extract(withRemark) {
      var rows = [['价税合计（大写）', 40, 400, 120], ['（小写）', 650, 400, 60],
        ['贰佰伍拾圆整', 220, 410, 120], ['¥250.00', 720, 410, 80],
        ['备', 30, 455, 12], ['注', 30, 480, 12], ['开票人：李四', 40, 525, 120]];
      if (withRemark) rows.push(['合同金额250元，按月结算', 60, 425, 240], ['收款银行：测试银行', 60, 445, 220]);
      var input = { imgW: 1000, imgH: 600, text: '', lines: rows.map(function(r) {
        return { confidence: 1, words: [{ text: r[0], x: r[1], y: r[2], w: r[3], h: 12 }] };
      }) };
      return _extractItemNameRemarkByCoords(_buildWords(input.lines, input.imgW, input.imgH), input.imgW, input.imgH).invoiceRemark;
    }
    // issue #56：无合计锚点、竖排「备/注」标签在备注框下部、正文在标签上方、
    // 页脚取值与标签字交错（开/票/张三/人/：）——页脚既不能混进备注，也不能按 x 排序拼出错序文本
    function sample(hasLabel, withAmounts) {
      var content = '购买方名称:测试公司;纳税人识别号:91310000123456789A;地址:测试地址;开户行及账号:测试银行1234567890';
      var lines = [];
      if (withAmounts) lines.push(
        { confidence: 1, words: [
          { text: '¥', x: 1775.0, y: 1080.8, w: 16.7, h: 33.3 },
          { text: '1,250.00', x: 1808.3, y: 1080.8, w: 100.0, h: 33.3 }] },
        { confidence: 1, words: [{ text: '（壹仟贰佰伍拾圆整）', x: 222.0, y: 1149.8, w: 500.0, h: 33.3 }] });
      lines.push({ confidence: 1, words: [{ text: content, x: 187.9, y: 1226.8, w: 1600.0, h: 33.3 }] });
      if (hasLabel) lines.push({ confidence: 1, words: [
        { text: '备', x: 121.0, y: 1296.8, w: 35.4, h: 33.3 },
        { text: '注', x: 121.0, y: 1364.8, w: 35.4, h: 33.3 }] });
      lines.push({ confidence: 1, words: [
        { text: '开', x: 205.5, y: 1572.3, w: 35.4, h: 33.3 },
        { text: '票', x: 276.4, y: 1572.3, w: 35.4, h: 33.3 },
        { text: '张三', x: 333.5, y: 1576.8, w: 100.0, h: 33.3 },
        { text: '人', x: 347.3, y: 1572.3, w: 35.4, h: 33.3 },
        { text: '：', x: 417.7, y: 1572.3, w: 35.4, h: 33.3 }] });
      return _extractItemNameRemarkByCoords(_buildWords(lines, 2480, 1653), 2480, 1653).invoiceRemark;
    }
    return [extract(false), extract(true), sample(true, true), sample(false, false),
      _isFooterLineText('开票金额合计人民币'), _isFooterLineText('开票张三人：')];
  });
  equal(remarks[0], '', '空备注不混入基线较低的大写/数字金额');
  equal(remarks[1], '合同金额250元，按月结算 收款银行：测试银行', '居中备注标签上方的正文及正文金额保留');
  equal(remarks[2], '购买方名称:测试公司;纳税人识别号:91310000123456789A;地址:测试地址;开户行及账号:测试银行1234567890',
    'issue #56 无锚点时取标签上方正文，交错页脚不混入');
  equal(remarks[3], '购买方名称:测试公司;纳税人识别号:91310000123456789A;地址:测试地址;开户行及账号:测试银行1234567890',
    'issue #56 无标签兜底同样剔除交错页脚行');
  equal([remarks[4], remarks[5]], [false, true], '交错页脚识别不误伤正文');

  await evaluate(() => {
    window.reviewSetup = function(cols, reverse, ordinal) {
      var svg = 'data:image/svg+xml,' + encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" width="250" height="350"><rect width="250" height="350" fill="red"/></svg>');
      S.files = ['A', 'B', 'C'].map(name => createFileObj({ name: name, type: 'image', previewUrl: svg, ow: 250, oh: 350 }));
      S.files[1].copies = 3;
      S.layout = { cols: cols, rows: 1 };
      document.getElementById('pageOrder').value = reverse ? 'reverse' : 'normal';
      var matches = [];
      getActiveFiles().forEach(function(f, i) { if (f === S.files[1]) matches.push(i); });
      var pos = matches[ordinal];
      S.currentPage = Math.floor(pos / cols);
      S.selectedSlot = pos % cols;
      updatePreview();
    };
    window.reviewSelected = function() {
      return { name: getSelectedFileObj()?.name, ordinal: captureSlotSelection()?.ordinal,
        page: S.currentPage, slot: S.selectedSlot, active: getActiveFiles().map(f => f.name) };
    };
  });
  equal(await evaluate(() => { reviewSetup(5, false, 1); setSlotCopies(1); return reviewSelected(); }),
    { name: 'B', ordinal: 0, page: 0, slot: 1, active: ['A', 'B', 'C'] }, '减少副本后仍选中原票');
  equal(await evaluate(() => { reviewSetup(2, false, 2); setSlotCopies(1); return reviewSelected(); }),
    { name: 'B', ordinal: 0, page: 0, slot: 1, active: ['A', 'B', 'C'] }, '跨页副本减少');
  equal(await evaluate(() => { reviewSetup(2, true, 2); setSlotCopies(1); return reviewSelected(); }),
    { name: 'B', ordinal: 0, page: 0, slot: 1, active: ['C', 'B', 'A'] }, '倒序副本减少');
  equal(await evaluate(() => { reviewSetup(2, false, 1); setSlotCopies(3); return captureSlotSelection().ordinal; }), 1, '份数不变保留副本序号');
  equal(await evaluate(() => { reviewSetup(2, false, 2); _ctxIdx = 0; ctxSetCopies(3); return getSelectedFileObj().name; }), 'B', '右键修改前面的票仍保留选中');
  equal(await evaluate(() => { reviewSetup(2, false, 2); setAllCopies({ stopPropagation() {} }, 1); return getSelectedFileObj().name; }), 'B', '批量份数保持选中');
  equal(await evaluate(() => { reviewSetup(2, false, 2); openSlotInvModal(); document.getElementById('mCopies').value = 1; confirmInvModal(); return reviewSelected(); }),
    { name: 'B', ordinal: 0, page: 0, slot: 1, active: ['A', 'B', 'C'] }, '自定义弹窗份数保持选中');
  equal(await evaluate(() => { reviewSetup(5, false, 1); setSlotCopies(1); deleteSlotInvoice(); return S.files.map(f => f.name); }), ['A', 'C'], '份数减少后删除原票');

  // 单票份数的唯一入口是浮动工具条「份数」按钮（issue #60 起双击票面不再设份数）
  await evaluate(() => { reviewSetup(3, false, 0); S.files[1].copies = 1; closeSlotCopyMenu(); selectSlot(1); });
  const copiesBtn = await evaluate(() => { var r = document.getElementById('slotCopiesBtn').getBoundingClientRect(); return { x: r.x + r.width / 2, y: r.y + r.height / 2 }; });
  await cdp('Input.dispatchMouseEvent', { type: 'mousePressed', ...copiesBtn, button: 'left', clickCount: 1 });
  await cdp('Input.dispatchMouseEvent', { type: 'mouseReleased', ...copiesBtn, button: 'left', clickCount: 1 });
  equal(await evaluate(() => document.getElementById('slotCopyMenu').classList.contains('hidden')), false, '工具条份数按钮打开菜单');
  const item = await evaluate(() => { var r = document.querySelector('#slotCopyMenu [data-n="3"]').getBoundingClientRect(); return { x: r.x + r.width / 2, y: r.y + r.height / 2 }; });
  await cdp('Input.dispatchMouseEvent', { type: 'mousePressed', ...item, button: 'left', clickCount: 1 });
  await cdp('Input.dispatchMouseEvent', { type: 'mouseReleased', ...item, button: 'left', clickCount: 1 });
  equal(await evaluate(() => ({ copies: S.files[1].copies, selected: getSelectedFileObj().name })), { copies: 3, selected: 'B' }, '菜单真实鼠标点击设份数');

  // 双击票面只做选中、不再弹份数菜单（issue #60：双击入口已移除，勿回归）
  // 鼠标落点取槽位下部，避开浮动工具条骑在票面上沿的区域
  async function slotPoint(idx) {
    return evaluate(idx => {
      var r = document.querySelector('.invoice-slot[data-slot-idx="' + idx + '"]').getBoundingClientRect();
      return { x: r.x + r.width / 2, y: r.bottom - 30 };
    }, idx);
  }
  async function mouse(type, point, clickCount = 1) {
    await cdp('Input.dispatchMouseEvent', { type, ...point, button: 'left', clickCount });
  }
  async function click(point) {
    await mouse('mousePressed', point);
    await mouse('mouseReleased', point);
  }
  async function copyMenuOpen() {
    return evaluate(() => !document.getElementById('slotCopyMenu').classList.contains('hidden'));
  }

  await evaluate(() => { reviewSetup(3, false, 0); S.files[1].copies = 1; S.selectedSlot = -1; closeSlotCopyMenu(); updatePreview(); });
  const point = await slotPoint(1);
  for (let count = 1; count <= 2; count++) {
    await mouse('mousePressed', point, count);
    await mouse('mouseReleased', point, count);
    await delay(60);
  }
  equal(await copyMenuOpen(), false, '双击票面不再弹份数菜单');
  equal(await evaluate(() => S.selectedSlot), 1, '双击票面仍选中该票');

  const toolbarBtn = await evaluate(() => { var r = document.getElementById('slotCopiesBtn').getBoundingClientRect(); return { x: r.x + r.width / 2, y: r.y + r.height / 2 }; });
  await click(toolbarBtn);
  equal(await copyMenuOpen(), true, '选中票面后工具条可打开份数菜单');
  await click(await slotPoint(2));
  equal(await copyMenuOpen(), false, '点击其它票关闭份数菜单');

  // 窄格子收成图标要按「非 compact 固有宽度」判定：读上一档宽度会来回翻转，奇数格停在 395 宽压邻格
  const tbAudit = await evaluate(() => {
    var svg = 'data:image/svg+xml,' + encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" width="250" height="350"><rect width="250" height="350" fill="red"/></svg>');
    S.feat.reimburse = false; S.feat.pasteMode = false;
    S.layout = { cols: 3, rows: 3 };
    var per = getPerPage(getSettings());
    S.files = Array.from({ length: per }, (_, i) => createFileObj({ name: 'p' + i + '.png', type: 'image', previewUrl: svg, ow: 250, oh: 350 }));
    S.currentPage = 0; updatePreview();
    var out = [];
    for (var i = 0; i < per; i++) {
      selectSlot(i);
      var el = document.querySelector('.invoice-slot[data-slot-idx="' + i + '"]');
      var tb = document.getElementById('slotToolbar');
      out.push({ compact: tb.classList.contains('compact'), fits: tb.offsetWidth + 8 <= el.offsetWidth });
    }
    selectSlot(-1);
    return out;
  });
  equal(tbAudit.length, 9, '3×3 槽位数');
  equal(tbAudit.every(x => x.compact && x.fits), true, '3×3 工具条全程收成图标且不压邻格');

  // 工具条容器 pointer-events:none，点内边距会穿透成预览底：不能当成「点空白处」清掉选中
  await evaluate(() => { reviewSetup(2, false, 0); S.files.forEach(f => { f.copies = 1; }); updatePreview(); selectSlot(1); });
  const tbPad = await evaluate(() => { var r = document.getElementById('slotToolbar').getBoundingClientRect(); return { x: Math.round(r.left + 2), y: Math.round(r.top + 2) }; });
  await click(tbPad);
  equal(await evaluate(() => ({ sel: S.selectedSlot, hidden: document.getElementById('slotToolbar').classList.contains('hidden') })), { sel: 1, hidden: false }, '点工具条内边距不取消选中');
  const blankHit = await evaluate(() => { var r = document.getElementById('previewWrap').getBoundingClientRect(); return { x: Math.round(r.right - 12), y: Math.round(r.bottom - 12) }; });
  await click(blankHit);
  equal(await evaluate(() => S.selectedSlot), -1, '点预览空白仍取消选中');

  // 删除后按文件对象重定位选中：槽位序号会整体前移，沿用旧索引会指到别的票
  const rmState = await evaluate(() => {
    var svg = 'data:image/svg+xml,' + encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" width="250" height="350"><rect width="250" height="350" fill="red"/></svg>');
    function setup(n) {
      S.feat.reimburse = false; S.feat.pasteMode = false;
      S.layout = { cols: 2, rows: 1 };
      S.files = Array.from({ length: n }, (_, i) => createFileObj({ name: 'r' + i + '.png', type: 'image', previewUrl: svg, ow: 250, oh: 350 }));
      S.currentPage = 0; updatePreview();
    }
    function snap() {
      var f = getSelectedFileObj();
      return { sel: S.selectedSlot, name: f ? f.name : null, tbHidden: document.getElementById('slotToolbar').classList.contains('hidden') };
    }
    var out = {};
    setup(3); selectSlot(0); rmFile(0); out.rmSelected = snap();
    setup(1); selectSlot(0); rmFile(0); out.rmLast = snap();
    setup(3); selectSlot(1); rmFile(0); out.rmBefore = snap();
    setup(3); selectSlot(1); rmFile(2); out.rmAfter = snap();
    setup(3); selectSlot(1); deleteSelected(); out.bulkCleared = snap();
    selectSlot(-1);
    return out;
  });
  equal(rmState.rmSelected, { sel: -1, name: null, tbHidden: true }, '删除选中票后清空选中');
  equal(rmState.rmLast, { sel: -1, name: null, tbHidden: true }, '删空后工具条隐藏');
  equal(rmState.rmBefore, { sel: 0, name: 'r1.png', tbHidden: false }, '删除前面的票后选中跟随原对象');
  equal(rmState.rmAfter, { sel: 1, name: 'r1.png', tbHidden: false }, '删除后面的票不动选中');
  equal(rmState.bulkCleared, { sel: -1, name: null, tbHidden: true }, '批量删除后清空选中且工具条隐藏');

  // 排序 / 列表重排会让预览槽位换票，选中必须按对象跟随
  const reorderState = await evaluate(() => {
    var svg = 'data:image/svg+xml,' + encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" width="250" height="350"><rect width="250" height="350" fill="red"/></svg>');
    function setup(dates) {
      S.feat.reimburse = false; S.feat.pasteMode = false;
      S.layout = { cols: 3, rows: 1 };
      S.files = dates.map((d, i) => createFileObj({ name: 'o' + i + '.png', type: 'image', previewUrl: svg, ow: 250, oh: 350, invoiceDate: d }));
      S.currentPage = 0; updatePreview();
    }
    function snap() {
      var f = getSelectedFileObj();
      return { sel: S.selectedSlot, name: f ? f.name : null, order: getActiveFiles().map(o => o.name).join(',') };
    }
    var out = {};
    setup(['2026-01-03', '2026-01-01', '2026-01-02']); selectSlot(1); sortByDate(1); out.sortDate = snap();
    setup(['2026-01-03', '2026-01-01', '2026-01-02']); selectSlot(0); swapFiles(0, 2); out.swapList = snap();
    setup(['2026-01-03', '2026-01-01', '2026-01-02']); selectSlot(0); moveFileTo(0, 2, 'after'); out.moveList = snap();
    selectSlot(-1);
    return out;
  });
  equal(reorderState.sortDate, { sel: 0, name: 'o1.png', order: 'o1.png,o2.png,o0.png' }, '日期排序后选中跟随原票');
  equal(reorderState.swapList, { sel: 2, name: 'o0.png', order: 'o2.png,o1.png,o0.png' }, '列表对调后选中跟随原票');
  equal(reorderState.moveList, { sel: 2, name: 'o0.png', order: 'o1.png,o2.png,o0.png' }, '列表插入后选中跟随原票');

  // 勾选变化同样改变排版成员：取消勾选的票移出排版，选中要么跟随、要么清空
  const checkState = await evaluate(() => {
    var svg = 'data:image/svg+xml,' + encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" width="250" height="350"><rect width="250" height="350" fill="red"/></svg>');
    function setup(n) {
      S.feat.reimburse = false; S.feat.pasteMode = false;
      S.layout = { cols: 3, rows: 1 };
      S.files = Array.from({ length: n }, (_, i) => createFileObj({ name: 'c' + i + '.png', type: 'image', previewUrl: svg, ow: 250, oh: 350 }));
      S.currentPage = 0; updatePreview();
    }
    function snap() {
      var f = getSelectedFileObj();
      return { sel: S.selectedSlot, name: f ? f.name : null, tbHidden: document.getElementById('slotToolbar').classList.contains('hidden') };
    }
    var out = {};
    setup(3); selectSlot(2); togCheck(0); out.uncheckBefore = snap();
    setup(3); selectSlot(0); togCheck(0); out.uncheckSelected = snap();
    setup(3); selectSlot(1); deselectAll(); out.deselectAll = snap();
    selectSlot(-1);
    return out;
  });
  equal(checkState.uncheckBefore, { sel: 1, name: 'c2.png', tbHidden: false }, '取消勾选前面的票后选中跟随原票');
  equal(checkState.uncheckSelected, { sel: -1, name: null, tbHidden: true }, '取消勾选选中票后清空选中');
  equal(checkState.deselectAll, { sel: -1, name: null, tbHidden: true }, '全部取消勾选后清空选中');

  const updates = await evaluate(async () => {
    isTauri = true;
    APP_VERSION = '2.6.16-beta.8';
    var requests = [], shown = [], toasts = [];
    invoke = function(cmd, args) { return new Promise(function(resolve, reject) { requests.push({ channel: args.channel, resolve: resolve, reject: reject }); }); };
    showUpdateModal = function(info) { shown.push(info.channel); };
    toast = function(text) { toasts.push(text); };
    localStorage.setItem(_UPDATE_CHANNEL_KEY, 'beta');
    localStorage.removeItem(_UPDATE_CACHE_KEY);
    _updateChecking = false;
    checkForUpdates(false);
    setUpdateChannel('stable');
    checkForUpdates(false);
    requests[0].resolve({ has_update: true, channel: 'beta', latest_version: '2.6.16-beta.9' });
    await Promise.resolve();
    var stale = { shown: shown.slice(), cache: localStorage.getItem(_UPDATE_CACHE_KEY), checking: _updateChecking, requests: requests.length };
    requests[1].resolve({ has_update: false, channel: 'stable', current_version: APP_VERSION });
    await Promise.resolve();
    var accepted = { channel: JSON.parse(localStorage.getItem(_UPDATE_CACHE_KEY)).channel, checking: _updateChecking, message: toasts[toasts.length - 1] };
    setUpdateChannel('beta');
    checkForUpdates(false);
    setUpdateChannel('stable');
    checkForUpdates(false);
    requests[2].reject('旧渠道错误');
    await Promise.resolve(); await Promise.resolve();
    var rejected = { checking: _updateChecking, staleErrorShown: toasts.some(t => t.includes('旧渠道错误')) };
    requests[3].resolve({ has_update: false, channel: 'stable', current_version: APP_VERSION });
    await Promise.resolve();
    // 旧版串渠道缓存不能被静默检查复用。
    localStorage.setItem(_UPDATE_CACHE_KEY, JSON.stringify({ ts: Date.now(), ver: APP_VERSION, channel: 'stable', info: { has_update: true, channel: 'beta' } }));
    checkForUpdates(true);
    var poisonedCacheFetched = requests.length === 5;
    requests[4].resolve({ has_update: false, channel: 'stable', current_version: APP_VERSION });
    await Promise.resolve();
    // 切走又切回原渠道时，最早的响应仍然过期。
    checkForUpdates(false);
    setUpdateChannel('beta'); setUpdateChannel('stable');
    requests[5].resolve({ has_update: true, channel: 'stable', latest_version: '2.6.16' });
    await Promise.resolve();
    return { stale: stale, accepted: accepted, rejected: rejected, poisonedCacheFetched: poisonedCacheFetched,
      returnedChannelStale: shown.length === 0 && localStorage.getItem(_UPDATE_CACHE_KEY) === null };
  });
  equal(updates.stale, { shown: [], cache: null, checking: true, requests: 2 }, '旧响应不能污染缓存或解除新请求锁');
  equal(updates.accepted, { channel: 'stable', checking: false, message: '已是最新版本 v2.6.16-beta.8' }, '当前渠道响应正常生效');
  equal(updates.rejected, { checking: true, staleErrorShown: false }, '旧渠道错误不影响新请求');
  equal(updates.poisonedCacheFetched, true, '拒绝历史串渠道缓存');
  equal(updates.returnedChannelStale, true, '切走再切回仍丢弃旧响应');
  // issue #62：双面 + 逐份按源文件边界补空白页；预览 / 保存 PDF 不补
  const duplex = await evaluate(() => {
    var mk = function (p) { return { id: p, name: p, _filePath: p, previewUrl: 'data:image/png;base64,AAAA', ow: 100, oh: 100, rotation: 0, type: 'image' }; };
    var a = mk('C:/a.jpg'), b = mk('C:/b.jpg');
    var flags = function (pages) { return pages.map(function (p) { return p === null; }); };
    var s = getSettings();
    s.rows = 1; s.cols = 1; s.copies = 1; s.duplex = true; s.collate = true;
    s.reimburseMode = false; s.photoMode = false; s.pasteMode = false;
    var reqPrint = buildLayoutRequest([a, b], s, null, { forPrint: true });
    var reqSave = buildLayoutRequest([a, b], s);
    s.copies = 2;
    var reqMoreCopies = buildLayoutRequest([a, b], s, null, { forPrint: true });
    return {
      oneUp: flags(padDuplexSeparators([[a], [b]])),
      sameFile: flags(padDuplexSeparators([[a], [a]])),
      printBlanks: reqPrint.pages.filter(function (p) { return !!p.blank; }).length,
      saveBlanks: reqSave.pages.filter(function (p) { return !!p.blank; }).length,
      printPages: reqPrint.pages.length,
      savePages: reqSave.pages.length,
      copiesStable: JSON.stringify(reqMoreCopies.pages) === JSON.stringify(reqPrint.pages)
    };
  });
  equal(duplex.oneUp, [false, true, false, true], '双面逐份：跨文件边界补空白页且末尾补齐偶数页');
  equal(duplex.sameFile, [false, false], '双面逐份：同一文件多页不插空白页');
  equal({ print: duplex.printBlanks, save: duplex.saveBlanks, printPages: duplex.printPages, savePages: duplex.savePages },
    { print: 2, save: 0, printPages: 4, savePages: 2 }, '双面逐份：仅打印请求含空白分隔页');
  equal(duplex.copiesStable, true, '双面逐份：改份数不改变 pages，PDF 缓存不误失效');
  // 启动恢复期：文件列表恢复中禁止落盘，避免 filePaths 被残缺 S.files 覆盖
  const restoreLock = await evaluate(() => {
    var before = localStorage.getItem('ticketchan-settings');
    _restoringFileList = true;
    _restoreSavePending = false;
    saveSettings();
    var blocked = _restoreSavePending === true;
    var unchanged = localStorage.getItem('ticketchan-settings') === before;
    _restoringFileList = false;
    _restoreSavePending = false;
    return blocked && unchanged;
  });
  equal(restoreLock, true, '文件列表恢复期 saveSettings 挂起，不写残缺快照');
  // printMode 旧键迁移：设置体写入失败时保留旧键，下次启动可再迁移
  const migrateFallback = await evaluate(() => {
    var settingsBefore = localStorage.getItem('ticketchan-settings');
    var pmBefore = localStorage.getItem('ticketchan-print-mode');
    var selectBefore = document.getElementById('printMode').value;
    localStorage.setItem('ticketchan-print-mode', 'confirm');
    localStorage.removeItem('ticketchan-settings');
    var origSetItem = Storage.prototype.setItem;
    Storage.prototype.setItem = function () { throw new Error('quota'); };
    try { migrateLegacyPrintMode(); } finally { Storage.prototype.setItem = origSetItem; }
    var kept = localStorage.getItem('ticketchan-print-mode') === 'confirm';
    if (settingsBefore === null) localStorage.removeItem('ticketchan-settings');
    else localStorage.setItem('ticketchan-settings', settingsBefore);
    if (pmBefore === null) localStorage.removeItem('ticketchan-print-mode');
    else localStorage.setItem('ticketchan-print-mode', pmBefore);
    document.getElementById('printMode').value = selectBefore;
    return kept;
  });
  equal(migrateFallback, true, 'printMode 迁移写入失败时保留旧键');

  await cdp('Browser.close');
  console.log('PASS: ' + checks + ' regression checks');
}

run().catch(error => { console.error(error); process.exitCode = 1; }).finally(async () => {
  if (socket) socket.close();
  browser.kill();
  await delay(200);
  // 只清理本次测试创建且已核对位于系统临时目录的 profile。
  if (path.dirname(profile) === path.resolve(os.tmpdir()) && path.basename(profile).startsWith('ticketchan-regression-')) {
    try { fs.rmSync(profile, { recursive: true, force: true }); } catch {}
  }
});
