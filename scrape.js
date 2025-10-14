// Granot -> Salesmen Performance -> click each "Total In" and bucket sources
import 'dotenv/config';
import { chromium } from 'playwright';

const BASE_URL     = process.env.GRANOT_BASE || 'https://fox.hellomoving.com/vanguamos/admin.htm';
const GAS_ENDPOINT = process.env.GAS_ENDPOINT || '';

const NETWORK_ID   = process.env.GRANOT_NET_ID   || process.env.NETWORK_ID;
const NETWORK_PASS = process.env.GRANOT_NET_PASS || process.env.NETWORK_PASS;
const GRANOT_USER  = process.env.GRANOT_USER;
const GRANOT_PASS  = process.env.GRANOT_PASS;

const HEADLESS     = String(process.env.HEADLESS || 'true').toLowerCase() === 'true';
const RETRIES      = Number(process.env.RETRIES || 1);
const NET_IDLE_MS  = Number(process.env.NET_IDLE_MS || 1500);
const DEBUG        = String(process.env.DEBUG || 'false').toLowerCase() === 'true';

for (const [k,v] of [
  ['GRANOT_BASE', BASE_URL],
  ['GRANOT_NET_ID', NETWORK_ID],
  ['GRANOT_NET_PASS', NETWORK_PASS],
  ['GRANOT_USER', GRANOT_USER],
  ['GRANOT_PASS', GRANOT_PASS],
]) if (!v) throw new Error(`Missing env var: ${k}`);

const VENDOR_ORDER = ['Raw','Em Semi','Mover Matcher','MM Inbound','MF Paper','MF Calls','Semi AH','1-800-BOOK'];
const vendorTemplate = () => Object.fromEntries(VENDOR_ORDER.map(v => [v, 0]));
const SOURCE_TO_VENDOR = {
  'Equate Media - Raw Calls': 'Raw',
  'Equate Media - Raw Calls - Email': 'Em Semi',
  'Equate Media - Raw Calls — Email': 'Em Semi',
  'EM Semi': 'Em Semi',
  'Mover Matcher': 'Mover Matcher',
  'MM Inbound': 'MM Inbound',
  'MF Paper': 'MF Paper',
  'MF Calls': 'MF Calls',
  'Semi AH': 'Semi AH',
  '1-800-BOOK': '1-800-BOOK',
};

// ---------- DATE HELPERS (Eastern Time, Monday→Sunday) ----------
const TZ = 'America/New_York';

const mmddyyyy = d => {
  const mm = String(d.getMonth()+1).padStart(2,'0');
  const dd = String(d.getDate()).padStart(2,'0');
  return `${mm}/${dd}/${d.getFullYear()}`;
};

function partsInTZ(date, timeZone) {
  const fmt = new Intl.DateTimeFormat('en-US', {
    timeZone,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit'
  });
  const parts = Object.fromEntries(fmt.formatToParts(date).map(p => [p.type, p.value]));
  return { year: +parts.year, month: +parts.month, day: +parts.day };
}
function dateFromParts({ year, month, day }) {
  return new Date(year, month - 1, day, 12);
}
function getWeekRangeInTZ({ timeZone = TZ, lastWeek = false } = {}) {
  const { year, month, day } = partsInTZ(new Date(), timeZone);
  const local = dateFromParts({ year, month, day });
  const isoDow = (local.getDay() + 6) % 7; // 0=Mon .. 6=Sun
  const monday = new Date(local);
  monday.setDate(local.getDate() - isoDow - (lastWeek ? 7 : 0));
  const sunday = new Date(monday);
  sunday.setDate(monday.getDate() + 6);
  return { monday, sunday };
}
// ---------------------------------------------------------------

async function ensureAdminMenu(page) {
  await page.goto(BASE_URL, { waitUntil: 'domcontentloaded', timeout: 60000 });

  // Gate 1 (Network)
  let frame = await (await page.waitForSelector('frame[name="content"]', { timeout: 30000 })).contentFrame();
  await frame.locator('input[name="NetworkID"], input[type="text"]').first().fill(NETWORK_ID);
  await frame.locator('input[name="Password"], input[type="password"]').first().fill(NETWORK_PASS);
  await Promise.all([
    page.waitForLoadState('domcontentloaded'),
    frame.locator('input[name="LOGON"], input[type="submit"]').first().click()
  ]);
  await page.waitForTimeout(NET_IDLE_MS);

  // Gate 2 (Granot user)
  frame = await (await page.waitForSelector('frame[name="content"]', { timeout: 30000 })).contentFrame();
  await frame.locator('input[type="text"]').first().fill(GRANOT_USER);
  await frame.locator('input[type="password"]').first().fill(GRANOT_PASS);
  await Promise.all([
    page.waitForLoadState('domcontentloaded'),
    frame.locator('input.inputlogin, input[name="LOGON"], input[type="submit"]').first().click()
  ]);
  await page.waitForTimeout(NET_IDLE_MS);

  return await (await page.waitForSelector('frame[name="content"]', { timeout: 30000 })).contentFrame();
}

async function openSalesmenPerformance(context, adminFrame) {
  // Pull the URL for menu 21 from the admin script
  const perfPath = await adminFrame.evaluate(() => {
    const txt = Array.from(document.scripts).map(s => s.textContent || '').join('\n');
    const m = txt.match(/if\s*\(\s*i\s*==\s*21\s*\)\s*window\.open\('([^']+)'/i);
    return m ? m[1] : '';
  });
  if (!perfPath) throw new Error('Could not extract Salesmen Performance URL');

  const origin = new URL(adminFrame.url()).origin || 'https://fox.hellomoving.com';
  const perfUrl = origin + perfPath;

  const perfPage = await context.newPage();
  await perfPage.goto(perfUrl, { waitUntil: 'domcontentloaded', timeout: 60000 });
  return perfPage;
}

async function setDateRange(perfPage, fromDate, toDate) {
  await perfPage.fill('#Date1', fromDate);
  await perfPage.fill('#Date2', toDate);
  await perfPage.locator('input.SUBMIT, input[type="button"][value="Submit"]').click();
  await perfPage.waitForLoadState('domcontentloaded');
  await perfPage.waitForTimeout(600);
}

function vendorFromSource(srcText) {
  const s = (srcText || '').trim();
  if (SOURCE_TO_VENDOR[s]) return SOURCE_TO_VENDOR[s];
  // loose matches as backup
  if (/equate media.*raw/i.test(s) && /email/i.test(s)) return 'Em Semi';
  if (/equate media.*raw/i.test(s)) return 'Raw';
  if (/mover\s*matcher/i.test(s)) return 'Mover Matcher';
  if (/mm\s*inbound/i.test(s)) return 'MM Inbound';
  if (/mf\s*paper/i.test(s)) return 'MF Paper';
  if (/mf\s*calls/i.test(s)) return 'MF Calls';
  if (/semi\s*ah/i.test(s)) return 'Semi AH';
  if (/1[\s\-]?800[\s\-]?book/i.test(s)) return '1-800-BOOK'; // updated line
  return null;
}

async function scrapeOnce() {
  const browser = await chromium.launch({ headless: HEADLESS });
  const context = await browser.newContext();
  const page = await context.newPage();

  try {
    const adminFrame = await ensureAdminMenu(page);
    const perfPage = await openSalesmenPerformance(context, adminFrame);

    // ---------- Use Eastern Time week (Monday→Sunday) ----------
    const { monday, sunday } = getWeekRangeInTZ({ lastWeek: false });
    await setDateRange(perfPage, mmddyyyy(monday), mmddyyyy(sunday));
    const weekLabel = `${mmddyyyy(monday)}-${mmddyyyy(sunday)}`;
    // -----------------------------------------------------------

    // The grid is the first table with border="1" and width="99%"
    const grid = perfPage.locator('table[border="1"][width="99%"]').first();
    await grid.waitFor({ state: 'visible', timeout: 20000 });

    const origin = new URL(perfPage.url()).origin;
    const reps = [];

    const rows = await grid.locator(':scope > tbody > tr').all(); // header + rows + footer
    for (let r = 1; r < rows.length - 2; r++) {
      const row = rows[r];
      const tds = await row.locator('td').all();
      if (tds.length === 0) continue;

      const name = (await tds[0].innerText()).replace(/\u00a0/g,' ').trim();
      if (!name) continue;

      const counts = vendorTemplate();

      // Total In is the 2nd column with an <a> when > 0
      const link = tds[1].locator('a').first();
      if (await link.count()) {
        const text = (await link.innerText()).trim();
        if (text && /^\d+$/.test(text)) {
          const href = await link.getAttribute('href');
          const absolute = href?.startsWith('http') ? href : origin + href;

          // try to get the popup, otherwise force-open
          let detailPage;
          const popupPromise = perfPage.context().waitForEvent('page').catch(() => null);
          await link.click({ button: 'middle' }).catch(() => {});
          detailPage = await popupPromise;
          if (!detailPage) {
            detailPage = await perfPage.context().newPage();
            await detailPage.goto(absolute, { waitUntil: 'domcontentloaded', timeout: 60000 });
          } else if (detailPage.url() === 'about:blank') {
            await detailPage.goto(absolute, { waitUntil: 'domcontentloaded', timeout: 60000 });
          } else {
            await detailPage.waitForLoadState('domcontentloaded');
          }

          // Detail table: <table bgcolor="#EEEEEE" ... width="98%">
          const detTable = detailPage
            .locator('table[bgcolor="#EEEEEE"], table[width="98%"]').first();
          await detTable.waitFor({ state: 'visible', timeout: 15000 });

          const detHead = await detTable.locator('tr').first().locator('td,th').allInnerTexts();
          const detCol = {};
          detHead.forEach((h,i) => { detCol[String(h).toLowerCase().trim()] = i; });
          const srcIdx = detCol['source'] ?? (detHead.length - 1);

          const detRows = await detTable.locator('tr').count();
          if (DEBUG) console.log(`[${name}] rows: ${detRows - 1}`);

          for (let i = 1; i < detRows - 1; i++) {
            const tds2 = detTable.locator('tr').nth(i).locator('td');
            const n = await tds2.count();
            if (n < 2) continue;

            const srcText = (await tds2.nth(Math.min(srcIdx, n - 1)).innerText() || '').trim();
            const vendor = vendorFromSource(srcText);
            if (vendor && counts[vendor] != null) counts[vendor] += 1;
          }

          if (!HEADLESS) await detailPage.waitForTimeout(200);
          await detailPage.close();
        }
      }

      reps.push({ rep: name, vendors: counts });
    }

    // ✅ NEW: include a flat list of rep names (for Column A placement on the sheet)
    const repNames = reps.map(r => r.rep);

    const payload = {
      date: weekLabel,           // tab name per week
      scraped_at: new Date().toISOString(),
      repNames,                  // <--- added list of reps
      reps                       // existing detailed vendor counts
    };

    if (!GAS_ENDPOINT) {
      console.log('Payload:\n', JSON.stringify(payload, null, 2));
    } else {
      const res = await fetch(GAS_ENDPOINT, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      console.log('Posted to GAS:', await res.text());
    }
  } finally {
    await browser.close();
  }
}

(async () => {
  for (let i = 1; i <= RETRIES; i++) {
    try { await scrapeOnce(); break; }
    catch (e) {
      console.error(`Attempt ${i}/${RETRIES} failed:`, e.message || e);
      if (i === RETRIES) process.exitCode = 1;
      else await new Promise(r => setTimeout(r, 1200));
    }
  }
})();