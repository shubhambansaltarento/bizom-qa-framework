/**
 * Fetches every company setting live from /settings/allsettings/1/1 (the
 * "All Settings" page grid), using the logged-in session's cookies and XSRF token.
 *
 * Use from a test, once logged in:
 *   import { fetchSettings } from '../utils/config/settings/fetchSettings';
 *   const { settings } = await fetchSettings(page.context());
 *   const stocktaking = settings.find(s => s.name === 'stocktaking')?.value;
 *
 * Or run standalone (logs in using .env, writes reports/<out>.json and .xlsx):
 *   npm run settings
 *   npm run settings -- --search stock --print
 *
 *   --search <text>  only keep settings whose name or description contains text (case-insensitive)
 *   --print          list the settings in the terminal
 *   --out <name>     report base name in reports/ (default settings)
 *
 * settings-api.json (this folder) holds the API path and the page used as Referer.
 */
import { chromium, type BrowserContext } from '@playwright/test';
import ExcelJS from 'exceljs';
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import { parseArgs } from 'node:util';
import { authHeaders, login, SNAPSHOT_DIR } from '../permissions/fetchPermissionEntities';

export type Setting = {
  id: string;
  name: string;
  setting_type: string;
  value: string;
  description: string;
  isjson: string;
  company_id: string;
  for_device: string;
  is_active: string;
  created: string;
  modified: string;
  isuserworkflow: string;
  type: string;
  submodule_mapping_id: string;
  rule_set: string | null;
};
export type SettingsSnapshot = { fetchedAt: string; baseUrl: string; total: number; settings: Setting[] };

export const settingsApi: { api: string; referer: string } =
  JSON.parse(readFileSync(path.join(__dirname, 'settings-api.json'), 'utf8'));

const env = (k: string): string => {
  const v = process.env[k];
  if (!v) throw new Error(`Missing ${k} in .env`);
  return v;
};

/**
 * Calls the settings API with the logged-in user's cookies + X-XSRF-TOKEN and
 * returns every setting. Throws if the call fails, the session is not valid,
 * the list is empty, or a row has no id / name or a duplicate name.
 */
export async function fetchSettings(
  context: BrowserContext,
  { baseUrl = env('BIZOM_BASE_URL') }: { baseUrl?: string } = {},
): Promise<SettingsSnapshot> {
  const headers = await authHeaders(context, baseUrl, baseUrl + settingsApi.referer);
  const url = `${baseUrl}${settingsApi.api}?_=${Date.now()}`; // cache buster, as the page sends
  const res = await fetch(url, { headers });
  const text = await res.text();
  if (res.status !== 200) throw new Error(`settings: ${settingsApi.api} returned HTTP ${res.status}`);

  let body: any;
  try {
    body = JSON.parse(text);
  } catch {
    throw new Error(`settings: ${settingsApi.api} did not return JSON (session expired?): ${text.slice(0, 200)}`);
  }
  if (!Array.isArray(body?.data)) throw new Error(`settings: response has no data array: ${text.slice(0, 200)}`);

  const settings: Setting[] = body.data.map((row: any) => row?.Setting ?? {});
  if (!settings.length) throw new Error(`settings: ${settingsApi.api} returned an empty list`);

  const problems: string[] = [];
  const seen = new Set<string>();
  settings.forEach((s, i) => {
    if (!s.id) problems.push(`row ${i + 1}: no id`);
    if (!s.name) problems.push(`row ${i + 1} (id ${s.id}): no name`);
    else if (seen.has(s.name)) problems.push(`row ${i + 1}: duplicate name "${s.name}"`);
    seen.add(s.name);
  });
  if (problems.length) throw new Error(`settings: ${problems.length} bad row(s): ${problems.slice(0, 10).join('; ')}`);

  return { fetchedAt: new Date().toISOString(), baseUrl, total: settings.length, settings };
}

/** Excel workbook: every setting, plus counts by setting_type and type. */
export async function writeSettingsWorkbook(settings: Setting[], file: string) {
  const wb = new ExcelJS.Workbook();
  const sheet = (name: string, cols: [string, keyof Setting | string, number][], rows: object[]) => {
    const ws = wb.addWorksheet(name, { views: [{ state: 'frozen', ySplit: 1 }] });
    ws.columns = cols.map(([header, key, width]) => ({ header, key, width }));
    ws.addRows(rows);
    ws.getRow(1).font = { bold: true };
    ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: cols.length } };
  };

  sheet('Settings', [
    ['ID', 'id', 8], ['Name', 'name', 40], ['Value', 'value', 30], ['Description', 'description', 60],
    ['Setting Type', 'setting_type', 14], ['Type', 'type', 10], ['For Device', 'for_device', 11],
    ['Active', 'is_active', 8], ['Is JSON', 'isjson', 8], ['User Workflow', 'isuserworkflow', 13],
    ['Submodule Mapping', 'submodule_mapping_id', 18], ['Rule Set', 'rule_set', 40],
    ['Created', 'created', 20], ['Modified', 'modified', 20], ['Company ID', 'company_id', 12],
  ], settings.map(s => ({ ...s, id: Number(s.id), submodule_mapping_id: s.submodule_mapping_id ? Number(s.submodule_mapping_id) : null })));

  const counts = (key: keyof Setting) => {
    const m = new Map<string, number>();
    for (const s of settings) m.set(String(s[key] ?? '') || '(blank)', (m.get(String(s[key] ?? '') || '(blank)') ?? 0) + 1);
    return [...m].sort((a, b) => b[1] - a[1]);
  };
  sheet('Summary', [['Field', 'field', 16], ['Value', 'value', 20], ['Settings', 'count', 10]], [
    ...counts('setting_type').map(([value, count]) => ({ field: 'setting_type', value, count })),
    ...counts('type').map(([value, count]) => ({ field: 'type', value, count })),
    ...counts('for_device').map(([value, count]) => ({ field: 'for_device', value, count })),
    ...counts('is_active').map(([value, count]) => ({ field: 'is_active', value, count })),
  ]);

  await wb.xlsx.writeFile(file);
}

// Standalone run: log in, fetch all settings, write reports.
if (require.main === module) {
  process.loadEnvFile(path.resolve(__dirname, '../../../.env'));
  (async () => {
    const { values: args } = parseArgs({
      options: {
        search: { type: 'string' },
        print: { type: 'boolean', default: false },
        out: { type: 'string', default: 'settings' },
      },
    });
    const browser = await chromium.launch({ headless: process.env.HEADLESS !== 'false' });
    try {
      const context = await browser.newContext();
      await login(await context.newPage());
      const snapshot = await fetchSettings(context);
      console.log(`Logged in. ${snapshot.total} settings fetched from ${settingsApi.api}`);

      const q = args.search?.toLowerCase();
      const settings = q
        ? snapshot.settings.filter(s => s.name.toLowerCase().includes(q) || (s.description ?? '').toLowerCase().includes(q))
        : snapshot.settings;
      if (q) console.log(`${settings.length} match "${args.search}"`);

      if (args.print) {
        const pad = (s: string, n: number) => (s.length > n ? s.slice(0, n - 1) + '…' : s.padEnd(n));
        for (const s of settings)
          console.log(`  ${pad(s.id, 6)} ${pad(s.name, 40)} = ${pad(String(s.value ?? ''), 24)} [${s.setting_type}] ${(s.description ?? '').replace(/\s+/g, ' ')}`);
      }

      mkdirSync(SNAPSHOT_DIR, { recursive: true });
      const jsonFile = path.join(SNAPSHOT_DIR, `${args.out}.json`);
      const xlsxFile = path.join(SNAPSHOT_DIR, `${args.out}.xlsx`);
      writeFileSync(jsonFile, JSON.stringify({ ...snapshot, search: args.search ?? null, total: settings.length, settings }, null, 2) + '\n');
      await writeSettingsWorkbook(settings, xlsxFile);
      console.log(`Wrote ${path.relative(process.cwd(), jsonFile)} and ${path.relative(process.cwd(), xlsxFile)}`);
    } finally {
      await browser.close();
    }
  })().catch(e => { console.error(e.message ?? e); process.exit(1); });
}
