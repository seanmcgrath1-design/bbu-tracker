// run-handoff-auto.js
// One-command automated handoff:
//   1. Ask Apps Script which sites are ready for handoff
//   2. Download every CQ in one browser session (Playwright)
//   3. Poll until each CQ is visible in Drive (cloud) — kills the local→cloud sync race
//   4. Trigger the handoff drafts (created with CQs attached)
//
// Usage: npm run handoff
// Requires a gitignored .handoff.env at the repo root (or env vars):
//   HANDOFF_API_URL=https://script.google.com/macros/s/XXXX/exec
//   HANDOFF_API_TOKEN=<same token set via setupHandoffApiToken in Apps Script>
// Set VZ_PASSWORD as an env var to skip the password prompt.

const readline = require('readline');
const { spawn } = require('child_process');
const path = require('path');
const fs = require('fs');
const os = require('os');

function loadEnvFile() {
  const p = path.join(__dirname, '..', '.handoff.env');
  if (!fs.existsSync(p)) return;
  for (const line of fs.readFileSync(p, 'utf8').split(/\r?\n/)) {
    const m = line.match(/^\s*([A-Za-z0-9_]+)\s*=\s*(.*?)\s*$/);
    if (m && !process.env[m[1]]) process.env[m[1]] = m[2];
  }
}

function promptPassword(prompt) {
  return new Promise((resolve) => {
    process.stdout.write(prompt);
    process.stdin.setRawMode(true);
    process.stdin.resume();
    process.stdin.setEncoding('utf8');
    let password = '';
    process.stdin.on('data', function handler(char) {
      if (char === '\r' || char === '\n' || char === '\x04') {
        process.stdin.setRawMode(false);
        process.stdin.pause();
        process.stdin.removeListener('data', handler);
        process.stdout.write('\n');
        resolve(password);
      } else if (char === '\x03') {
        process.exit();
      } else if (char === '\x7f' || char === '\x08') {
        if (password.length > 0) {
          password = password.slice(0, -1);
          process.stdout.clearLine(0);
          process.stdout.cursorTo(0);
          process.stdout.write(prompt + '*'.repeat(password.length));
        }
      } else {
        password += char;
        process.stdout.write('*');
      }
    });
  });
}

// The web app is domain-restricted (Verizon disables anonymous access), so calls must be authenticated.
// Mint a Google access token from the refresh token clasp stored at login (~/.clasprc.json).
let _accessToken = null;
async function getAccessToken() {
  if (_accessToken) return _accessToken;
  const p = path.join(os.homedir(), '.clasprc.json');
  if (!fs.existsSync(p)) throw new Error('clasp creds not found (~/.clasprc.json). Run `npx clasp login`.');
  const cr = JSON.parse(fs.readFileSync(p, 'utf8')).tokens.default;
  const r = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ client_id: cr.client_id, client_secret: cr.client_secret, refresh_token: cr.refresh_token, grant_type: 'refresh_token' }),
  });
  const tk = await r.json();
  if (!tk.access_token) throw new Error('Could not refresh Google token (run `npx clasp login`): ' + JSON.stringify(tk));
  _accessToken = tk.access_token;
  return _accessToken;
}

async function api(action, extra) {
  const accessToken = await getAccessToken();
  const res = await fetch(process.env.HANDOFF_API_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + accessToken },
    body: JSON.stringify(Object.assign({ token: process.env.HANDOFF_API_TOKEN, action }, extra || {})),
    redirect: 'follow',
  });
  const text = await res.text();
  let data;
  try { data = JSON.parse(text); }
  catch { throw new Error(`Non-JSON response (HTTP ${res.status}). Check the URL/access: ${text.slice(0, 200)}`); }
  if (data.error) throw new Error(`API: ${data.error}`);
  return data;
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

function runPlaywright(fuzeIds, password) {
  return new Promise((resolve) => {
    const child = spawn(
      'npx',
      ['playwright', 'test', '--grep', 'CQ Retrieval', '--headed', '--project=chromium'],
      {
        env: { ...process.env, VZ_PASSWORD: password, FUZE_IDS: fuzeIds.join(',') },
        stdio: 'inherit',
        shell: true,
        cwd: path.join(__dirname, '..'),
      }
    );
    child.on('exit', (code) => resolve(code ?? 0));
  });
}

async function main() {
  loadEnvFile();
  if (!process.env.HANDOFF_API_URL || !process.env.HANDOFF_API_TOKEN) {
    console.error('Missing HANDOFF_API_URL / HANDOFF_API_TOKEN. Set them in .handoff.env (see header).');
    process.exit(1);
  }

  console.log('\n=== Automated Handoff ===\n');

  let password = process.env.VZ_PASSWORD;
  if (password) console.log('Using VZ_PASSWORD from environment.\n');
  else password = (await promptPassword('Verizon Password: ')).trim();

  console.log('Fetching ready handoff sites...');
  const { ready } = await api('ready');
  if (!ready || ready.length === 0) {
    console.log('No sites ready for handoff. Nothing to do.');
    process.exit(0);
  }
  const fuzeIds = ready.map((r) => String(r.fuze).trim()).filter(Boolean);
  console.log(`Ready (${ready.length}): ${ready.map((r) => `${r.site} (${r.fuze})`).join(', ')}\n`);

  console.log('Downloading CQs via Playwright...\n');
  await runPlaywright(fuzeIds, password);

  console.log('\nConfirming CQs are visible in Drive (cloud)...');
  const deadline = Date.now() + 3 * 60 * 1000;
  let pending = fuzeIds.slice();
  while (pending.length && Date.now() < deadline) {
    const { status } = await api('cqStatus', { fuzeIds: pending });
    const present = status.filter((s) => s.present).map((s) => s.fuze);
    if (present.length) console.log(`  synced: ${present.join(', ')}`);
    pending = status.filter((s) => !s.present).map((s) => s.fuze);
    if (pending.length) await sleep(10000);
  }
  if (pending.length) {
    console.warn(`  WARNING: not visible after wait: ${pending.join(', ')} — their drafts may lack a CQ.`);
  }

  console.log('\nGenerating handoff drafts...');
  const result = await api('generate');
  console.log(`\nDone. ${result.created} draft(s) created.`);
  if (result.missing && result.missing.length) {
    console.log(`No CQ attached for: ${result.missing.join(', ')}`);
  }
  console.log('Review the drafts in Gmail before sending.');
  process.exit(0);
}

main().catch((err) => { console.error('\nError:', (err && err.message) || err); process.exit(1); });
