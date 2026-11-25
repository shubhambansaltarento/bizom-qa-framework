/**
 * Loads the roles / designations / users for the Add Permissions page at run time.
 * They change per environment and over time, so nothing here is read from a
 * committed list: every call hits the live APIs with the logged-in user's session.
 *
 * Only the entities named in permission-targets.json (or the env overrides below)
 * are returned, and the run fails if any of them does not exist on the server.
 *
 * Use from a test, once the user is logged in:
 *   import { fetchPermissionEntities, getPermissionCombinations } from '../utils/config/permissions/fetchPermissionEntities';
 *   const entities = await fetchPermissionEntities(page.context());
 *   const combos = getPermissionCombinations(entities);
 *
 * Or run standalone (logs in using .env, then fetches):
 *   npx tsx utils/config/permissions/fetchPermissionEntities.ts
 *
 * Files in this folder (all hand-edited config):
 *   permissions-config.json   module -> { groupKey: label }
 *   entity-apis.json          groupKey -> API path
 *   permission-targets.json   groupKey -> [id or name, ...]  which entities to test
 *
 * Choosing targets:
 *   permission-targets.json   { "role": ["ADMIN"], "designation": ["DSM"], "user": ["Pradeep Kumar", "42"] }
 *     - each value is an entity id or its exact name (case-insensitive)
 *     - ["*"] selects every entity in that group
 *     - [] or a missing key skips that group (its API is not called)
 *   Env overrides, comma-separated, replace the file's list for that group:
 *     PERMISSION_ROLES, PERMISSION_DESIGNATIONS, PERMISSION_USERS
 *     e.g. PERMISSION_USERS="Guru,Arun Salesman" npx tsx ...   (set to "" to skip the group)
 *
 * The standalone run writes a snapshot of what it loaded to reports/ for
 * debugging. Those files are git-ignored and are never read back.
 */
import { chromium, type BrowserContext, type Page } from '@playwright/test';
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import path from 'node:path';

export type GroupKey = 'role' | 'designation' | 'user';
export type Entity = { id: string; name: string };
export type PermissionEntities = { fetchedAt: string; baseUrl: string } & Record<GroupKey, Entity[]>;
export type PermissionTargets = Record<GroupKey, string[]>;
export type PermissionCombination = {
  module: string;
  group: { key: GroupKey; label: string };
  entity: { key: string; label: string };
  entityApiCalled: string[];
};

const DIR = __dirname;
const readJson = (file: string) => JSON.parse(readFileSync(path.join(DIR, file), 'utf8'));

export const GROUPS: readonly GroupKey[] = ['role', 'designation', 'user'];
export const moduleGroups: Record<string, Record<GroupKey, string>> = readJson('permissions-config.json');
export const entityApis: Record<GroupKey, string> = readJson('entity-apis.json');
export const TARGETS_FILE = path.join(DIR, 'permission-targets.json');
export const SNAPSHOT_DIR = path.resolve(DIR, '../../../reports');

const TARGET_ENV: Record<GroupKey, string> = {
  role: 'PERMISSION_ROLES',
  designation: 'PERMISSION_DESIGNATIONS',
  user: 'PERMISSION_USERS',
};

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

/**
 * Which entities to test, per group: permission-targets.json, with any
 * PERMISSION_ROLES / PERMISSION_DESIGNATIONS / PERMISSION_USERS env var replacing
 * that group's list. Throws on a malformed file or if every group is empty.
 */
export function loadPermissionTargets(): PermissionTargets {
  const file: Record<string, unknown> = JSON.parse(readFileSync(TARGETS_FILE, 'utf8'));
  const unknown = Object.keys(file).filter(k => !GROUPS.includes(k as GroupKey));
  if (unknown.length) throw new Error(`permission-targets.json: unknown key(s) ${unknown.join(', ')}; expected ${GROUPS.join(', ')}`);

  const targets = {} as PermissionTargets;
  for (const g of GROUPS) {
    const fromEnv = process.env[TARGET_ENV[g]];
    const list = fromEnv !== undefined ? fromEnv.split(',') : (file[g] ?? []);
    if (!Array.isArray(list)) throw new Error(`permission-targets.json: "${g}" must be an array of ids or names`);
    targets[g] = list.map(v => String(v).trim()).filter(Boolean);
  }
  if (GROUPS.every(g => !targets[g].length))
    throw new Error(`No permission targets configured. Add ids or names to ${path.basename(TARGETS_FILE)} or set ${Object.values(TARGET_ENV).join(' / ')}.`);
  return targets;
}

/** Picks the wanted entities (by id, else by case-insensitive name) from a live list. Throws if one is missing or ambiguous. */
export function selectEntities(group: GroupKey, all: Entity[], wanted: string[]): Entity[] {
  if (wanted.includes('*')) return all;
  const picked = new Map<string, Entity>();
  for (const w of wanted) {
    const byId = all.filter(e => e.id === w);
    const matches = byId.length ? byId : all.filter(e => e.name.trim().toLowerCase() === w.toLowerCase());
    if (!matches.length) {
      const names = all.map(e => `${e.name || '(blank)'} [${e.id}]`);
      const shown = names.slice(0, 30).join(', ') + (names.length > 30 ? `, … ${names.length - 30} more` : '');
      throw new Error(`${group} "${w}" not found on the server (${entityApis[group]}). Available: ${shown}`);
    }
    if (matches.length > 1)
      throw new Error(`${group} "${w}" matches ${matches.length} entities (ids ${matches.map(e => e.id).join(', ')}); use the id instead`);
    picked.set(matches[0]!.id, matches[0]!);
  }
  return [...picked.values()];
}

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
 * Calls the entity APIs live with the logged-in user's cookies and returns only
 * the configured targets. Groups with no targets are not fetched and come back
 * as []. Throws if a call fails or a configured target is not on the server.
 */
export async function fetchPermissionEntities(
  context: BrowserContext,
  {
    baseUrl = env('BIZOM_BASE_URL'),
    referer = env('ADD_PERMISSION_URL'),
    targets = loadPermissionTargets(),
  }: { baseUrl?: string; referer?: string; targets?: PermissionTargets } = {},
): Promise<PermissionEntities> {
  const headers = await authHeaders(context, baseUrl, referer);
  const groups = GROUPS.filter(g => targets[g].length);

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
    return [group, selectEntities(group, entities, targets[group])] as const;
  }));

  const result = { fetchedAt: new Date().toISOString(), baseUrl, role: [], designation: [], user: [] } as PermissionEntities;
  for (const [group, entities] of lists) result[group] = entities;
  return result;
}

/** Every module × group × entity to test, from entities loaded by fetchPermissionEntities. */
export function getPermissionCombinations(entities: PermissionEntities): PermissionCombination[] {
  return Object.entries(moduleGroups).flatMap(([module, labels]) =>
    GROUPS.flatMap(group => entities[group].map(e => ({
      module,
      group: { key: group, label: labels[group] },
      entity: { key: e.id, label: e.name },
      entityApiCalled: [entityApis[group]],
    }))));
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

// Standalone run: log in, fetch, and write a debugging snapshot to reports/.
if (require.main === module) {
  process.loadEnvFile(path.resolve(DIR, '../../../.env'));
  (async () => {
    const targets = loadPermissionTargets(); // fail on bad config before opening a browser
    const browser = await chromium.launch({ headless: process.env.HEADLESS !== 'false' });
    try {
      const context = await browser.newContext();
      await login(await context.newPage());
      console.log('Logged in. Fetching permission entities...');
      const entities = await fetchPermissionEntities(context, { targets });
      const combos = getPermissionCombinations(entities);
      for (const g of GROUPS) {
        const picked = entities[g].map(e => `${e.name} [${e.id}]`).join(', ');
        console.log(`  ${g.padEnd(11)} ${targets[g].length ? picked : '(skipped)'}`);
      }
      mkdirSync(SNAPSHOT_DIR, { recursive: true });
      writeFileSync(path.join(SNAPSHOT_DIR, 'permission-entities.json'), JSON.stringify(entities, null, 2) + '\n');
      writeFileSync(path.join(SNAPSHOT_DIR, 'permission-combinations.json'), JSON.stringify(combos, null, 2) + '\n');
      console.log(`${combos.length} combinations. Snapshot written to ${path.relative(process.cwd(), SNAPSHOT_DIR)}/`);
    } finally {
      await browser.close();
    }
  })().catch(e => { console.error(e.message ?? e); process.exit(1); });
}
