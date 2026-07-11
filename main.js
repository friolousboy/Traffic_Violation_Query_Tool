/**
 * 违章查询脚本 — 从 sc.122.gov.cn 批量查询车辆违章信息（Node.js / Playwright 版）。
 *
 * 读取 example.xlsx 中的车牌号，依次查询每个车牌的违章记录，
 * 提取详细信息后保存到 output.xlsx。
 */

'use strict';

const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');
const XLSX = require('xlsx');

// ============================================================
// 配置常量
// ============================================================
const EXCEL_INPUT = 'example.xlsx';
const EXCEL_OUTPUT = 'output.xlsx';
const EXCEL_COLUMNS = ['车牌', '时间', '违法地点', '违法行为', '罚款金额', '记分值', '是否处理'];

const HEADLESS = false;          // true=后台静默（需先登录过且Cookie有效）
const SAVE_INTERVAL = 20;        // 每处理 N 个车牌保存一次
const DEFAULT_TIMEOUT = 5000;    // Playwright 操作默认超时（毫秒），was 5 seconds
const LOGIN_TIMEOUT = 300000;    // 等待用户扫码登录超时（毫秒），was 300 seconds
const PAGE_STABLE_WAIT = 5000;   // 翻页/查询后等表格稳定的最大毫秒数，was 5 seconds
const DETAIL_DELAY = 500;        // 详情面板打开后的稳定等待（毫秒），was 0.5
const MAX_ROW_RETRIES = 3;       // 单条违章记录最大重试次数
const MAX_PLATE_RETRIES = 3;     // 单张车牌最大重新搜索次数

const CAR_TYPES_ALL = ['小型汽车', '小型新能源汽车', '大型汽车'];

const BASE_URL = 'https://sc.122.gov.cn/views/memfyy/violation.html';

// ---- 元素定位器 ----
const SELECTOR_CAR_TYPE = '#hpzl';
const SELECTOR_PLATE_INPUT = '#hphm';
const SELECTOR_SUBMIT = '.btn-submit-veh';
const SELECTOR_VIOLATION_ROWS = '.surveil_view';
const SELECTOR_CLOSE = '#bind_close';

// 违章详情面板 — 字段名 → XPath
const DETAIL_XPATHS = {
  time:     '//*[@id="view"]/div[2]/div[1]/form/div[3]/span[2]',
  location: '//*[@id="view"]/div[2]/div[1]/form/div[4]/span[2]',
  behavior: '//*[@id="view"]/div[2]/div[1]/form/div[5]/span[2]',
  fine:     '//*[@id="view"]/div[2]/div[1]/form/div[7]/span[2]',
  points:   '//*[@id="view"]/div[2]/div[1]/form/div[8]/span[2]',
};


// ============================================================
// ConsoleTee — 同时输出到控制台和日志文件
// ============================================================
class ConsoleTee {
  constructor(logStream) {
    this._log = logStream;
  }

  hook() {
    const origOut = process.stdout.write.bind(process.stdout);
    const origErr = process.stderr.write.bind(process.stderr);
    const logStream = this._log;

    process.stdout.write = function (chunk, encoding, cb) {
      const str = typeof chunk === 'string' ? chunk : chunk.toString();
      logStream.write(str);
      return origOut(chunk, encoding, cb);
    };
    process.stderr.write = function (chunk, encoding, cb) {
      const str = typeof chunk === 'string' ? chunk : chunk.toString();
      logStream.write(str);
      return origErr(chunk, encoding, cb);
    };
  }
}


// ============================================================
// 爬虫类
// ============================================================
class ViolationScraper {
  /** 初始化：创建 Chromium 浏览器实例。 */
  constructor() {
    this.browser = null;
    this.context = null;
    this.page = null;
  }

  async init() {
    this.browser = await chromium.launch({
      headless: HEADLESS,
      args: HEADLESS ? ['--disable-gpu'] : [],
    });
    this.context = await this.browser.newContext({
      locale: 'zh-CN',
      viewport: { width: 1280, height: 900 },
    });
    this.page = await this.context.newPage();
    // 设置默认超时
    this.page.setDefaultTimeout(DEFAULT_TIMEOUT);

    // 一次性注册弹窗自动处理（消除 Python 版弹窗竞态问题）
    this.page.on('dialog', async (dialog) => {
      await dialog.accept();
    });
  }

  // ---- 等待工具（带进度显示） ----
  _tick(waited, total, label) {
    process.stdout.write(`  ${label}... ${waited}/${Math.round(total / 1000)}s\r`);
  }

  async _poll(predicate, timeoutMs, label = '等待') {
    const deadline = Date.now() + timeoutMs;
    while (Date.now() < deadline) {
      try {
        if (await predicate()) {
          const elapsed = Math.round((timeoutMs - (deadline - Date.now())) / 1000);
          console.log(`  ${label}... ✓ (${elapsed}s)` + ' '.repeat(10));
          return true;
        }
      } catch (_) {
        // 忽略 predicate 中的异常，继续轮询
      }
      const waited = Math.round((timeoutMs - (deadline - Date.now())) / 1000);
      this._tick(waited, timeoutMs, label);
      await this.page.waitForTimeout(1000);
    }
    console.log(`  ${label}... 超时 (${Math.round(timeoutMs / 1000)}s)` + ' '.repeat(10));
    return false;
  }

  // ============================================================
  // 登录处理
  // ============================================================
  async _isOnQueryPage() {
    const hasPlate = (await this.page.locator(SELECTOR_PLATE_INPUT).count()) > 0;
    const hasSubmit = (await this.page.locator(SELECTOR_SUBMIT).count()) > 0;
    return hasPlate && hasSubmit;
  }

  async _navigateToViolationQuery() {
    if (await this._isOnQueryPage()) return;

    const navKeywords = ['违法查询', '违章查询', '机动车违法', '违法处理'];
    for (const keyword of navKeywords) {
      const links = this.page.locator('a').filter({ hasText: keyword });
      if ((await links.count()) > 0) {
        try {
          await links.first().click();
          await this.page.waitForTimeout(2000);
          if (await this._isOnQueryPage()) {
            console.log(`已通过导航【${keyword}】进入查询页面`);
            return;
          }
        } catch (_) {
          continue;
        }
      }
    }

    // 自动导航失败，请用户手动操作
    console.log('\n' + '='.repeat(50));
    console.log('请在浏览器中手动点击进入【机动车违法查询】页面');
    console.log('='.repeat(50));
    const ok = await this._poll(() => this._isOnQueryPage(), 120000, '等待进入查询页面');
    if (!ok) throw new Error('等待进入违章查询页面超时，请检查后重试');
    console.log('已进入查询页面 ✓');
  }

  async _ensureLogin() {
    // 1. 打开登录页
    await this.page.goto('https://gab.122.gov.cn/m/login?t=2');
    console.log('\n' + '='.repeat(50));
    console.log('请在浏览器中扫描二维码完成登录');
    console.log('='.repeat(50));

    const loggedIn = async () => {
      const url = this.page.url();
      if (url.includes('login') || url.endsWith('/')) return false;
      return true;
    };

    const loginOk = await this._poll(loggedIn, LOGIN_TIMEOUT, '等待扫码登录');
    if (!loginOk) throw new Error(`等待登录超时（${Math.round(LOGIN_TIMEOUT / 1000)} 秒），请检查网络后重试`);
    console.log('登录成功 ✓');

    // 2. 等待用户点击公司
    console.log('\n' + '='.repeat(50));
    console.log('请在浏览器中点击【公司】进入');
    console.log('='.repeat(50));
    await this.page.waitForTimeout(2000);
    const oldBody = (await this.page.locator('body').innerText()) || '';

    const bodyChanged = async () => {
      const current = (await this.page.locator('body').innerText()) || '';
      return current !== oldBody;
    };

    const clickOk = await this._poll(bodyChanged, 300000, '等待点击公司');
    if (clickOk) {
      console.log('已进入公司主页 ✓');
    } else {
      console.log('未检测到页面跳转，继续...');
    }

    // 3. 导航到违章查询页面
    await this._navigateToViolationQuery();
  }

  // ============================================================
  // 数据读写
  // ============================================================
  static readPlates(filepath = EXCEL_INPUT) {
    const workbook = XLSX.readFile(filepath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet);
    return data.map(row => row['车牌']).filter(Boolean);
  }

  static saveResults(results, filepath = EXCEL_OUTPUT) {
    if (!results || results.length === 0) {
      console.log('无数据可保存');
      return;
    }
    const ws = XLSX.utils.aoa_to_sheet([EXCEL_COLUMNS, ...results]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '违章查询结果');
    XLSX.writeFile(wb, filepath);

    // 打印结果表格
    console.log(EXCEL_COLUMNS.join('\t'));
    for (const row of results) {
      console.log(row.join('\t'));
    }
    console.log(`已保存 ${results.length} 条记录至 ${filepath}`);
  }

  // ============================================================
  // 违章详情提取
  // ============================================================
  async _extractDetailFields() {
    const fields = {};
    for (const [name, xpath] of Object.entries(DETAIL_XPATHS)) {
      try {
        const el = this.page.locator(`xpath=${xpath}`);
        await el.waitFor({ state: 'visible', timeout: DEFAULT_TIMEOUT });
        fields[name] = (await el.innerText()) || '';
      } catch (_) {
        fields[name] = '';
      }
    }
    return fields;
  }

  async _closeDetailPanel() {
    try {
      const btn = this.page.locator(SELECTOR_CLOSE);
      await btn.waitFor({ state: 'visible', timeout: DEFAULT_TIMEOUT });
      await btn.click();
    } catch (_) {
      try {
        await this.page.locator(SELECTOR_CLOSE).click({ force: true });
      } catch (_e) {
        // 面板可能已关闭
      }
    }
  }

  async _getHandledStatus(rowIndex, total) {
    let xpath;
    if (total === 1) {
      xpath = '//*[@id="my-msg-list"]/tbody/tr/td[5]';
    } else {
      xpath = `//*[@id="my-msg-list"]/tbody/tr[${rowIndex + 1}]/td[5]`;
    }
    try {
      return (await this.page.locator(`xpath=${xpath}`).innerText({ timeout: DEFAULT_TIMEOUT })) || '';
    } catch (_) {
      return '';
    }
  }

  async _waitForTableStable() {
    let prevCount = -1;
    const deadline = Date.now() + PAGE_STABLE_WAIT;
    while (Date.now() < deadline) {
      const rows = this.page.locator(SELECTOR_VIOLATION_ROWS);
      const curCount = await rows.count();
      if (curCount > 0 && curCount === prevCount) {
        return;
      }
      prevCount = curCount;
      const elapsed = Math.round((PAGE_STABLE_WAIT - Math.max(0, deadline - Date.now())) / 1000);
      this._tick(elapsed, PAGE_STABLE_WAIT, '表格稳定中');
      await this.page.waitForTimeout(500);
    }
  }

  async _clickViewButton(rowIndex) {
    const xpath = `//*[@id="my-msg-list"]/tbody/tr[${rowIndex + 1}]/td[last()]/a`;
    const btn = this.page.locator(`xpath=${xpath}`);
    await btn.waitFor({ state: 'visible', timeout: DEFAULT_TIMEOUT });
    await btn.scrollIntoViewIfNeeded();
    await this.page.waitForTimeout(200);
    try {
      await btn.click();
    } catch (_) {
      await btn.click({ force: true });
    }
    // 等待详情面板渲染完毕
    try {
      const locXpath = DETAIL_XPATHS.location;
      await this.page.locator(`xpath=${locXpath}`).waitFor({ state: 'visible', timeout: 5000 });
    } catch (_) {
      // 面板可能还没渲染好，继续
    }
    await this.page.waitForTimeout(DETAIL_DELAY);
  }

  // ============================================================
  // 单条违章记录处理
  // ============================================================
  async _processViolationRow(plate, rowIndex, totalRows) {
    const handled = await this._getHandledStatus(rowIndex, totalRows);

    for (let attempt = 0; attempt < MAX_ROW_RETRIES; attempt++) {
      try {
        await this._clickViewButton(rowIndex);
        const fields = await this._extractDetailFields();

        if (fields.location && fields.location.trim()) {
          await this._closeDetailPanel();
          return [
            plate,
            fields.time,
            fields.location,
            fields.behavior,
            fields.fine,
            fields.points,
            handled,
          ];
        }

        // 关键字段为空，关闭面板后重试
        await this._closeDetailPanel();
        await this.page.waitForTimeout(500);

      } catch (_) {
        // 超时或元素异常，重试
        if (attempt < MAX_ROW_RETRIES - 1) continue;
      }
    }

    console.log(`  违章行 ${rowIndex + 1} 提取失败（已重试 ${MAX_ROW_RETRIES} 次）`);
    return null;
  }

  // ============================================================
  // 翻页处理（含完整验证逻辑）
  // ============================================================
  static async _isNextPageDisabled(locator) {
    const cls = (await locator.getAttribute('class')) || '';
    return cls.includes('disabled') || cls.includes('nolink');
  }

  async _goToNextPage(pageNum) {
    // 翻页前记录当前页快照，用于验证翻页是否真正生效
    const rows = this.page.locator(SELECTOR_VIOLATION_ROWS);
    const oldCount = await rows.count();
    const oldFirstText = oldCount > 0 ? (await rows.first().innerText()) || '' : '';

    // 检查页面数据是否已经变化
    const dataChanged = async () => {
      const now = this.page.locator(SELECTOR_VIOLATION_ROWS);
      const nowCount = await now.count();
      if (nowCount === 0) return false;               // 尚未加载出新数据
      if (nowCount !== oldCount) return true;          // 行数变了 → 确认翻页
      // 行数相同，比较首行内容
      try {
        return ((await now.first().innerText()) || '') !== oldFirstText;
      } catch (_) {
        return true;   // DOM 已刷新
      }
    };

    const clickAndWait = async (element) => {
      await element.click();

      // 等待页面数据变化，超时说明翻页没生效
      const start = Date.now();
      let changed = false;
      while (Date.now() - start < PAGE_STABLE_WAIT) {
        if (await dataChanged()) { changed = true; break; }
        await this.page.waitForTimeout(500);
      }
      if (!changed) return false;

      // 等待新数据稳定
      await this._waitForTableStable();

      // 最终验证：稳定后数据必须真的变了
      const finalRows = this.page.locator(SELECTOR_VIOLATION_ROWS);
      const finalCount = await finalRows.count();
      if (finalCount > 0 && finalCount === oldCount) {
        try {
          if (((await finalRows.first().innerText()) || '') === oldFirstText) {
            return false;  // 数据完全没变 → AJAX 闪烁后回退
          }
        } catch (_) {
          // DOM 已刷新
        }
      }

      return true;
    };

    // 模式 1：「下一页」链接
    const nextLinks = this.page.locator('a').filter({ hasText: '下一页' });
    const linkCount = await nextLinks.count();
    for (let i = 0; i < linkCount; i++) {
      const link = nextLinks.nth(i);
      try {
        if (!(await ViolationScraper._isNextPageDisabled(link))) {
          console.log(`  → 翻至第${pageNum + 1}页（下一页）`);
          return await clickAndWait(link);
        }
      } catch (_) {
        continue;
      }
    }
    if (linkCount > 0) return false;  // 有"下一页"但被禁用 → 最后一页

    // 模式 2：pagination 组件中的 » 或 > 符号
    try {
      const pagination = this.page.locator('[class*="pagination"], [class*="pager"]');
      if ((await pagination.count()) > 0) {
        const candidates = pagination.locator('a').filter({ hasText: /»|>/ });
        const candCount = await candidates.count();
        for (let i = 0; i < candCount; i++) {
          const el = candidates.nth(i);
          try {
            if (!(await ViolationScraper._isNextPageDisabled(el))) {
              console.log(`  → 翻至第${pageNum + 1}页（»）`);
              return await clickAndWait(el);
            }
          } catch (_) {
            continue;
          }
        }
      }
    } catch (_) {
      // 没有 pagination 容器，跳过
    }

    // 模式 3：页码链接，点当前激活页的下一个兄弟
    try {
      const active = this.page.locator('[class*="current"], [class*="active"]').first();
      if ((await active.count()) > 0) {
        const parent = active.locator('..');
        const siblings = parent.locator('a');
        const sibCount = await siblings.count();
        for (let i = 0; i < sibCount; i++) {
          const sib = siblings.nth(i);
          try {
            const sibCls = (await sib.getAttribute('class')) || '';
            if ((sibCls.includes('current') || sibCls.includes('active')) && i + 1 < sibCount) {
              console.log(`  → 翻至第${pageNum + 1}页（页码）`);
              return await clickAndWait(siblings.nth(i + 1));
            }
          } catch (_) {
            continue;
          }
        }
      }
    } catch (_) {
      // 找不到 active 页码
    }

    return false;
  }

  // ============================================================
  // 单车牌处理
  // ============================================================
  async _tryCarType(carType, plate) {
    // ---- 等下拉框选项加载完毕 ----
    const dropdown = this.page.locator(SELECTOR_CAR_TYPE);
    await dropdown.waitFor({ state: 'visible', timeout: DEFAULT_TIMEOUT });

    const dropdownReady = async () => {
      const options = dropdown.locator('option');
      const count = await options.count();
      if (count === 0) return false;
      for (let i = 0; i < count; i++) {
        if (await options.nth(i).innerText()) return true;
      }
      return false;
    };
    await this._poll(dropdownReady, 5000, '等待下拉框就绪');

    // 选择号牌种类
    try {
      await dropdown.selectOption({ label: carType });
    } catch (_) {
      // 按 label 失败，尝试按文本匹配 value
      const options = dropdown.locator('option');
      const count = await options.count();
      let matchedValue = null;
      for (let i = 0; i < count; i++) {
        const opt = options.nth(i);
        const text = (await opt.innerText()).trim();
        if (text === carType) {
          matchedValue = await opt.getAttribute('value');
          break;
        }
      }
      if (matchedValue) {
        await dropdown.selectOption({ value: matchedValue });
      } else {
        throw new Error(`找不到号牌种类: ${carType}`);
      }
    }

    // ---- 记下点击前的页面快照 ----
    const oldBody = (await this.page.locator('body').innerText()) || '';

    // 点击查询按钮
    const submit = this.page.locator(SELECTOR_SUBMIT);
    await submit.waitFor({ state: 'visible' });
    await submit.click();

    // ---- 等待页面刷新 ----
    const bodyChanged = async () => {
      const current = (await this.page.locator('body').innerText()) || '';
      return current !== oldBody;
    };
    await this._poll(bodyChanged, 3000, '等待页面刷新');

    // 等待违章行或「未查询到违法记录」
    const gotRows = async () => (await this.page.locator(SELECTOR_VIOLATION_ROWS).count()) > 0;
    const gotEmpty = async () => {
      const text = (await this.page.locator('body').innerText()) || '';
      return text.includes('未查询到违法记录');
    };

    let outcome;
    if (await this._poll(gotRows, 3000, '等待违章数据加载')) {
      outcome = 'rows';
    } else if (await gotEmpty()) {
      outcome = 'empty';
    } else {
      await this._poll(gotEmpty, 3000, '确认无违章');
      outcome = (await gotEmpty()) ? 'empty' : 'timeout';
    }

    if (outcome !== 'rows') {
      if (outcome === 'empty') {
        console.log(`  ${plate}: ${carType} - 无违章`);
      } else {
        console.log(`  ${plate}: ${carType} - 查询超时，跳过`);
      }
      return [];
    }

    // 表格稳固后再处理
    await this._waitForTableStable();

    // ---- 逐页提取 ----
    const allResults = [];
    let pageNum = 1;

    while (true) {
      const rows = this.page.locator(SELECTOR_VIOLATION_ROWS);
      const rowCount = await rows.count();
      if (rowCount === 0) break;

      console.log(`  ${plate}: ${carType} - 第${pageNum}页, ${rowCount}条`);
      for (let idx = 0; idx < rowCount; idx++) {
        const rowData = await this._processViolationRow(plate, idx, rowCount);
        if (rowData) {
          allResults.push(rowData);
          console.log(JSON.stringify(rowData));
        }
      }

      const hasNext = await this._goToNextPage(pageNum);
      if (!hasNext) break;
      pageNum++;
    }

    return allResults;
  }

  async processPlate(plate) {
    const plateNumber = plate.slice(1);  // 去掉首位（省份简称）

    // 等页面从上一次查询中恢复
    await this.page.waitForTimeout(500);

    const input = this.page.locator(SELECTOR_PLATE_INPUT);
    await input.waitFor({ state: 'visible' });
    await input.clear();
    await input.fill(plateNumber);
    await this.page.waitForTimeout(300);  // 等号牌种类下拉框响应输入

    // 根据车牌长度推断号牌种类顺序（新能源 8 位，蓝牌 7 位）
    const types = plate.length === 8
      ? ['小型新能源汽车', '小型汽车', '大型汽车']
      : ['小型汽车', '大型汽车'];

    for (let retry = 0; retry < MAX_PLATE_RETRIES; retry++) {
      try {
        for (const carType of types) {
          try {
            const results = await this._tryCarType(carType, plate);
            if (results && results.length > 0) return results;
            // 弹窗或无结果，继续下一类型
          } catch (_) {
            console.log(`  ${plate}: ${carType} 触发弹窗，尝试下一类型`);
            continue;
          }
        }
        console.log(`  ${plate}: 所有类型均失败，重新搜索`);
      } catch (_) {
        console.log(`  ${plate}: 页面元素缺失，重新搜索`);
      }
    }

    console.log(`  ${plate}: 超过最大重试次数（${MAX_PLATE_RETRIES}），跳过`);
    return [];
  }

  // ============================================================
  // 生命周期
  // ============================================================
  async run() {
    const plates = ViolationScraper.readPlates();
    const allResults = [];

    await this._ensureLogin();

    for (let idx = 0; idx < plates.length; idx++) {
      const plate = plates[idx];
      console.log(`\n${'='.repeat(40)}`);
      console.log(`开始搜索 [${idx + 1}/${plates.length}] ${plate}`);
      console.log(`${'='.repeat(40)}`);

      const plateResults = await this.processPlate(plate);
      allResults.push(...plateResults);

      // 每 N 条增量保存
      if ((idx + 1) % SAVE_INTERVAL === 0) {
        ViolationScraper.saveResults(allResults);
      }
    }

    // 最终保存
    console.log(`\n${'='.repeat(40)}`);
    console.log('查询完成，最终保存...');
    console.log(`${'='.repeat(40)}`);
    ViolationScraper.saveResults(allResults);
  }

  async quit() {
    try {
      if (this.browser) await this.browser.close();
    } catch (_) {
      // ignore
    }
  }
}


// ============================================================
// 入口
// ============================================================
(async () => {
  const logStream = fs.createWriteStream(path.join(__dirname, 'log.txt'), {
    flags: 'w',
    encoding: 'utf-8',
  });
  const tee = new ConsoleTee(logStream);
  tee.hook();

  const scraper = new ViolationScraper();
  try {
    await scraper.init();
    await scraper.run();
  } catch (e) {
    console.error(`\n[FATAL] ${e.message}`);
    console.error(e.stack);
  } finally {
    await scraper.quit();
    console.log('脚本结束。');
    logStream.end();
  }
})();
