/**
 * Fetches the role / designation / user lists used by the Add Permissions page
 * and saves them to permission-entities.json in this folder.
 *
 * Use from a test, once the user is logged in:
 *   import { fetchPermissionEntities } from '../utils/config/permissions/fetchPermissionEntities';
 *   const entities = await fetchPermissionEntities(page.context());
 *
 * Or run standalone (logs in using .env, then fetches):
 *   npx tsx utils/config/permissions/fetchPermissionEntities.ts
 *
 * Files in this folder:
 *   permissions-config.json   module -> { groupKey: label }
 *   entity-apis.json          groupKey -> API path
 *   permission-entities.json  groupKey -> [{ id, name }]  (written by this script)
 */
import { chromium, type BrowserContext, type Page } from '@playwright/test';
import { readFileSync, writeFileSync } from 'node:fs';
import path from 'node:path';

export type GroupKey = 'role' | 'designation' | 'user';
export type Entity = { id: string; name: string };
export type PermissionEntities = { fetchedAt: string; baseUrl: string } & Record<GroupKey, Entity[]>;

const DIR = __dirname;
const readJson = (file: string) => JSON.parse(readFileSync(path.join(DIR, file), 'utf8'));

export const moduleGroups: Record<string, Record<GroupKey, string>> = readJson('permissions-config.json');
export const entityApis: Record<GroupKey, string> = readJson('entity-apis.json');
export const ENTITIES_FILE = path.join(DIR, 'permission-entities.json');

// Each API returns a different shape; normalise all to [{ id, name }].
const parsers: Record<GroupKey, (body: any) => Entity[]> = {
  // [{ id, name }]
  role: body => body.map((r: any) => ({ id: String(r.id), name: r.name })),
  // { data: [{ Designation: { id, name, ... } }] }
  designation: body => body.data.map((d: any) => ({ id: String(d.Designation.id), name: d.Designation.name })),
  // { data: { "<id>": "<name>" } }
  user: body => Object.entries(body.data).map(([id, name]) => ({ id, name: String(name ?? '') })),
};

const env = (k: string): string => {
  const v = process.env[k];
  if (!v) throw new Error(`Missing ${k} in .env`);
  return v;
};

/** Request headers built from the logged-in context's cookies (CAKEPHP + XSRF-TOKEN). */
export async function authHeaders(context: BrowserContext, baseUrl: string, referer: string) {
  const cookies = await context.cookies(baseUrl);
  const xsrf = cookies.find(c => c.name === 'XSRF-TOKEN');
  if (!xsrf) throw new Error('XSRF-TOKEN cookie not found — is the user logged in?');
  if (!cookies.some(c => c.name === 'CAKEPHP')) throw new Error('CAKEPHP session cookie not found — is the user logged in?');

  return {
    Cookie: cookies.map(c => `${c.name}=${c.value}`).join('; '),
    'X-XSRF-TOKEN': decodeURIComponent(xsrf.value),
    Accept: 'application/json, text/javascript, */*; q=0.01',
    'X-Requested-With': 'XMLHttpRequest',
    Referer: referer,
    'Cache-Control': 'no-cache',
    Pragma: 'no-cache',
  };
}

/**
 * Calls the 3 entity APIs with the logged-in user's cookies, saves the result
 * to permission-entities.json and returns it. Throws if any call fails.
 */
export async function fetchPermissionEntities(
  context: BrowserContext,
  { baseUrl = env('BIZOM_BASE_URL'), referer = env('ADD_PERMISSION_URL'), save = true } = {},
): Promise<PermissionEntities> {
  const headers = await authHeaders(context, baseUrl, referer);
  const groups = Object.keys(entityApis) as GroupKey[];

  const lists = await Promise.all(groups.map(async group => {
    const url = `${baseUrl}${entityApis[group]}?_=${Date.now()}`; // cache buster, as the page sends
    const res = await fetch(url, { headers });
    const text = await res.text();
    if (res.status !== 200) throw new Error(`${group}: ${entityApis[group]} returned HTTP ${res.status}`);
    let entities: Entity[];
    try {
      entities = parsers[group](JSON.parse(text));
    } catch (e) {
      throw new Error(`${group}: could not parse ${entityApis[group]} response (${(e as Error).message}): ${text.slice(0, 200)}`);
    }
    if (!entities.length) throw new Error(`${group}: ${entityApis[group]} returned an empty list`);
    return [group, entities] as const;
  }));

  const result = { fetchedAt: new Date().toISOString(), baseUrl, ...Object.fromEntries(lists) } as PermissionEntities;
  if (save) writeFileSync(ENTITIES_FILE, JSON.stringify(result, null, 2) + '\n');
  return result;
}

async function login(page: Page) {
  await page.goto(env('BIZOM_BASE_URL') + env('LOGIN_PATH'));
  await page.fill('input[name="User[username]"]', env('BIZOM_USERNAME'));
  await page.fill('#password', env('BIZOM_PASSWORD'));
  await Promise.all([
    page.waitForURL(u => !u.pathname.includes('/users/login'), { timeout: Number(process.env.LOGIN_TIMEOUT ?? 60000) }),
    page.click('#loginBtn'),
  ]);
}

// Standalone run: log in, then fetch.
if (require.main === module) {
  process.loadEnvFile(path.resolve(DIR, '../../../.env'));
  (async () => {
    const browser = await chromium.launch({ headless: process.env.HEADLESS !== 'false' });
    try {
      const context = await browser.newContext();
      await login(await context.newPage());
      console.log('Logged in. Fetching permission entities...');
      const result = await fetchPermissionEntities(context);
      for (const g of Object.keys(entityApis) as GroupKey[])
        console.log(`  ${g.padEnd(11)} ${entityApis[g]} -> ${result[g].length} items`);
      console.log(`Saved to ${path.relative(process.cwd(), ENTITIES_FILE)}`);
    } finally {
      await browser.close();
    }
  })().catch(e => { console.error(e.message ?? e); process.exit(1); });
}
