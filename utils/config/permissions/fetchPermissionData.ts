/**
 * Loads the URL permissions (Add Permissions page grid) for every
 * module × group × entity combination, using /urlpermissions/loadPermissionData.
 *
 *   moduleKey      = module name from permissions-config.json  (e.g. "User Management")
 *   entityKey      = group key                                  ("role" | "designation" | "user")
 *   entityDataKey  = entity id from the live entity APIs         (e.g. ADMIN = 1)
 *
 * Which entities are covered comes from permission-targets.json / PERMISSION_* env
 * vars (see fetchPermissionEntities.ts); entities are always fetched live.
 *
 * Use from a test, once logged in:
 *   const entities = await fetchPermissionEntities(page.context());
 *   const results = await fetchAllPermissionData(page.context(), getPermissionCombinations(entities));
 *
 * Or run standalone (logs in using .env, writes reports/<out>.json and .xlsx):
 *   npm run permissions                     targets from permission-targets.json / PERMISSION_* env
 *   npm run permissions:all                 every role, designation and user
 *   npm run permissions:admin               ADMIN role in every module, printed -> reports/admin-permissions.*
 *   npm run permissions -- --user "E2E ADMIN,47" --print --out e2e-admin
 *
 *   --role / --designation / --user   comma-separated ids or names ("*" = all); groups not named are skipped
 *   --print                           list every permission in the terminal
 *   --out <name>                      report base name in reports/ (default permission-data)
 *
 * PERMISSION_CONCURRENCY (default 5) sets how many requests run in parallel.
 */
import { chromium, type BrowserContext } from '@playwright/test';
import ExcelJS from 'exceljs';
import { mkdirSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import { parseArgs } from 'node:util';
import {
  authHeaders, fetchPermissionEntities, getPermissionCombinations, loadPermissionTargets, login,
  GROUPS, SNAPSHOT_DIR, type GroupKey, type PermissionCombination, type PermissionTargets,
} from './fetchPermissionEntities';

export const PERMISSION_DATA_API = '/urlpermissions/loadPermissionData';
const COLUMNS = ['main_action', 'description', 'isBlocked', 'url'] as const;

export type PermissionAction = { action: string; description: string; url: string; isBlocked: string };
export type PermissionDataResult = {
  module: string;
  group: GroupKey;
  entity: { key: string; label: string };
  status: number;
  ms: number;
  recordsTotal: number | null;
  blockedCount: number;
  actions: PermissionAction[];
  pass: boolean;
  errors: string[];
};

/** Same DataTables form body the page sends; length=-1 asks for every row. */
export function permissionDataBody(moduleKey: string, entityKey: GroupKey, entityDataKey: string): string {
  const p = new URLSearchParams({ draw: '1' });
  COLUMNS.forEach((c, i) => {
    p.set(`columns[${i}][data]`, c);
    p.set(`columns[${i}][name]`, '');
    p.set(`columns[${i}][searchable]`, 'true');
    p.set(`columns[${i}][orderable]`, 'true');
    p.set(`columns[${i}][search][value]`, '');
    p.set(`columns[${i}][search][regex]`, 'false');
  });
  p.set('order[0][column]', '0');
  p.set('order[0][dir]', 'asc');
  p.set('order[0][name]', '');
  p.set('start', '0');
  p.set('length', '-1');
  p.set('search[value]', '');
  p.set('search[regex]', 'false');
  p.set('moduleKey', moduleKey);
  p.set('entityKey', entityKey);
  p.set('entityDataKey', entityDataKey);
  return p.toString();
}

/** Fetches and validates the permission grid for one combination. Never throws; problems go in `errors`. */
export async function fetchPermissionData(
  headers: Record<string, string>,
  baseUrl: string,
  combo: PermissionCombination,
): Promise<PermissionDataResult> {
  const result: PermissionDataResult = {
    module: combo.module, group: combo.group.key, entity: combo.entity,
    status: 0, ms: 0, recordsTotal: null, blockedCount: 0, actions: [], pass: false, errors: [],
  };
  const body = permissionDataBody(combo.module, combo.group.key, combo.entity.key);
  const init = {
    method: 'POST',
    headers: { ...headers, 'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8', Origin: baseUrl },
    body,
  };

  let text = '';
  for (let attempt = 1; attempt <= 2; attempt++) { // one retry on network error / 5xx
    const t0 = Date.now();
    try {
      const res = await fetch(baseUrl + PERMISSION_DATA_API, init);
      text = await res.text();
      result.status = res.status;
      result.ms = Date.now() - t0;
      if (res.status < 500) break;
    } catch (e) {
      result.ms = Date.now() - t0;
      if (attempt === 2) result.errors.push(`request failed: ${(e as Error).message}`);
    }
  }
  if (result.errors.length) return result;
  if (result.status !== 200) {
    result.errors.push(`HTTP ${result.status}: ${text.slice(0, 200)}`);
    return result;
  }

  let json: any;
  try {
    json = JSON.parse(text);
  } catch {
    result.errors.push(`response is not JSON (session expired?): ${text.slice(0, 200)}`);
    return result;
  }
  if (!Array.isArray(json?.data)) {
    result.errors.push(`response has no data array: ${text.slice(0, 200)}`);
    return result;
  }

  result.recordsTotal = Number(json.recordsTotal);
  result.actions = json.data.map((r: any) => ({
    action: String(r.main_action ?? ''),
    description: String(r.description ?? ''),
    url: String(r.url ?? ''),
    isBlocked: String(r.isBlocked ?? ''),
  }));
  result.blockedCount = result.actions.filter(a => a.isBlocked === '1').length;

  if (result.actions.length !== result.recordsTotal)
    result.errors.push(`got ${result.actions.length} rows but recordsTotal is ${json.recordsTotal}`);
  result.actions.forEach((a, i) => {
    if (!a.action) result.errors.push(`row ${i + 1}: main_action is empty`);
    if (!a.url) result.errors.push(`row ${i + 1} (${a.action}): url is empty`);
    if (a.isBlocked !== '0' && a.isBlocked !== '1') result.errors.push(`row ${i + 1} (${a.action}): isBlocked is "${a.isBlocked}", expected "0" or "1"`);
  });
  result.pass = result.errors.length === 0;
  return result;
}

/** Runs fetchPermissionData for every combination with limited parallelism, preserving order. */
export async function fetchAllPermissionData(
  context: BrowserContext,
  combos: PermissionCombination[],
  {
    baseUrl = process.env.BIZOM_BASE_URL ?? '',
    referer = process.env.ADD_PERMISSION_URL ?? '',
    concurrency = Number(process.env.PERMISSION_CONCURRENCY ?? 5),
    onProgress,
  }: { baseUrl?: string; referer?: string; concurrency?: number; onProgress?: (done: number, total: number) => void } = {},
): Promise<PermissionDataResult[]> {
  if (!baseUrl || !referer) throw new Error('BIZOM_BASE_URL and ADD_PERMISSION_URL must be set');
  const headers = await authHeaders(context, baseUrl, referer);
  const results = new Array<PermissionDataResult>(combos.length);
  let next = 0;
  let done = 0;
  await Promise.all(Array.from({ length: Math.max(1, concurrency) }, async () => {
    while (next < combos.length) {
      const i = next++;
      results[i] = await fetchPermissionData(headers, baseUrl, combos[i]!);
      onProgress?.(++done, combos.length);
    }
  }));
  return results;
}

/** Excel workbook: one row per combination, one row per permission, and the failures. */
export async function writePermissionDataWorkbook(results: PermissionDataResult[], file: string) {
  const wb = new ExcelJS.Workbook();
  const sheet = (name: string, cols: [string, string, number][], rows: object[]) => {
    const ws = wb.addWorksheet(name, { views: [{ state: 'frozen', ySplit: 1 }] });
    ws.columns = cols.map(([header, key, width]) => ({ header, key, width }));
    ws.addRows(rows);
    ws.getRow(1).font = { bold: true };
    ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: cols.length } };
    return ws;
  };
  const base = (r: PermissionDataResult) => ({
    module: r.module, group: r.group, id: Number(r.entity.key), name: r.entity.label || '(blank)',
  });

  sheet('Combinations', [
    ['Module', 'module', 32], ['Group', 'group', 12], ['Entity ID', 'id', 10], ['Entity Name', 'name', 32],
    ['HTTP', 'status', 8], ['Permissions', 'total', 12], ['Allowed', 'allowed', 10], ['Blocked', 'blocked', 10],
    ['Result', 'result', 10], ['Errors', 'errors', 60],
  ], results.map(r => ({
    ...base(r), status: r.status, total: r.recordsTotal, allowed: r.actions.length - r.blockedCount,
    blocked: r.blockedCount, result: r.pass ? 'PASS' : 'FAIL', errors: r.errors.join('; '),
  })));

  sheet('Permissions', [
    ['Module', 'module', 32], ['Group', 'group', 12], ['Entity ID', 'id', 10], ['Entity Name', 'name', 32],
    ['Action', 'action', 32], ['Description', 'description', 44], ['URL', 'url', 36], ['Blocked', 'blocked', 10],
  ], results.flatMap(r => r.actions.map(a => ({
    ...base(r), action: a.action, description: a.description, url: a.url, blocked: a.isBlocked === '1' ? 'Yes' : 'No',
  }))));

  const failed = results.filter(r => !r.pass);
  if (failed.length)
    sheet('Failures', [
      ['Module', 'module', 32], ['Group', 'group', 12], ['Entity ID', 'id', 10], ['Entity Name', 'name', 32],
      ['HTTP', 'status', 8], ['Errors', 'errors', 100],
    ], failed.map(r => ({ ...base(r), status: r.status, errors: r.errors.join('; ') })));

  await wb.xlsx.writeFile(file);
}

/** Prints each combination's permissions to the terminal, grouped by module. */
export function printPermissions(results: PermissionDataResult[]) {
  const pad = (s: string, n: number) => (s.length > n ? s.slice(0, n - 1) + '…' : s.padEnd(n));
  for (const r of results) {
    const head = `${r.module} | ${r.group} | ${r.entity.label || '(blank)'} [${r.entity.key}]`;
    console.log(`\n${head}  — ${r.actions.length} permissions, ${r.blockedCount} blocked${r.pass ? '' : `  FAIL: ${r.errors.join('; ')}`}`);
    for (const a of r.actions)
      console.log(`  ${a.isBlocked === '1' ? 'BLOCKED' : 'allowed'}  ${pad(a.action, 34)} ${pad(a.url, 38)} ${a.description}`);
  }
}

// Standalone run: log in, fetch entities live, load every combination's permissions, write reports.
//   npx tsx utils/config/permissions/fetchPermissionData.ts [--role ADMIN] [--designation DSM] [--user "Guru,4"] [--print] [--out name]
//   --print  also list every permission in the terminal
//   --out    report base name in reports/ (default permission-data -> permission-data.json + .xlsx)
if (require.main === module) {
  process.loadEnvFile(path.resolve(__dirname, '../../../.env'));
  (async () => {
    const { values: args } = parseArgs({
      options: {
        role: { type: 'string' },
        designation: { type: 'string' },
        user: { type: 'string' },
        print: { type: 'boolean', default: false },
        out: { type: 'string', default: 'permission-data' },
      },
    });
    // --role / --designation / --user pick exactly those entities (comma-separated ids or names, "*" = all);
    // groups not named are skipped. With none of them, permission-targets.json / PERMISSION_* env apply.
    const cliGroups = GROUPS.filter(g => args[g] !== undefined);
    const targets = cliGroups.length
      ? Object.fromEntries(GROUPS.map(g => [g, (args[g] ?? '').split(',').map(s => s.trim()).filter(Boolean)])) as PermissionTargets
      : loadPermissionTargets();
    const browser = await chromium.launch({ headless: process.env.HEADLESS !== 'false' });
    try {
      const context = await browser.newContext();
      await login(await context.newPage());
      const entities = await fetchPermissionEntities(context, { targets });
      const combos = getPermissionCombinations(entities);
      console.log(`Logged in. ${GROUPS.map(g => `${entities[g].length} ${g}`).join(', ')} -> ${combos.length} combinations`);

      const started = Date.now();
      let lastPct = -1;
      const results = await fetchAllPermissionData(context, combos, {
        onProgress: (done, total) => {
          const pct = Math.floor((done / total) * 10) * 10;
          if (pct !== lastPct) { lastPct = pct; console.log(`  ${done}/${total}`); }
        },
      });

      const failed = results.filter(r => !r.pass);
      const report = {
        runAt: new Date().toISOString(),
        baseUrl: entities.baseUrl,
        entitiesFetchedAt: entities.fetchedAt,
        total: results.length,
        passed: results.length - failed.length,
        failed: failed.length,
        results,
      };
      mkdirSync(SNAPSHOT_DIR, { recursive: true });
      const jsonFile = path.join(SNAPSHOT_DIR, `${args.out}.json`);
      const xlsxFile = path.join(SNAPSHOT_DIR, `${args.out}.xlsx`);
      writeFileSync(jsonFile, JSON.stringify(report, null, 2) + '\n');
      await writePermissionDataWorkbook(results, xlsxFile);

      if (args.print) printPermissions(results);

      console.log(`Done in ${Math.round((Date.now() - started) / 1000)}s: ${report.passed} passed, ${report.failed} failed`);
      for (const r of failed.slice(0, 20)) console.log(`  FAIL ${r.module} / ${r.group} / ${r.entity.label} [${r.entity.key}]: ${r.errors[0]}`);
      if (failed.length > 20) console.log(`  … ${failed.length - 20} more in the report`);
      console.log(`Wrote ${path.relative(process.cwd(), jsonFile)} and ${path.relative(process.cwd(), xlsxFile)}`);
      if (failed.length) process.exitCode = 1;
    } finally {
      await browser.close();
    }
  })().catch(e => { console.error(e.message ?? e); process.exit(1); });
}
