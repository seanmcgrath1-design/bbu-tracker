// run-Ericsson V1 SFP.js
// Ericsson V1 SFP — clones EO E000274573
// Usage: node "Bulk EO Creation/run-Ericsson V1 SFP.js"

const readline = require('readline');
const { spawn } = require('child_process');
const path = require('path');

async function main() {
  console.log('\n=== EO Clone Tool — Ericsson V1 SFP ===\n');

  // Prompt for password
  const rl1 = readline.createInterface({ input: process.stdin, output: process.stdout });
  const password = await new Promise((resolve) => rl1.question('Verizon Password: ', resolve));
  rl1.close();

  // Collect IDs — paste one per line, comma/space separated, or one continuous string (1–20 IDs)
  console.log('\nPaste your Project IDs (1–20), then press Enter twice:\n');
  const rl2 = readline.createInterface({ input: process.stdin, output: process.stdout });

  const collected = [];
  await new Promise((resolve) => {
    rl2.on('line', (line) => {
      const parts = line.trim().split(/[\s,]+/).filter(Boolean);
      collected.push(...parts);
      if (collected.length >= 20) resolve();
    });
    rl2.on('close', resolve);
  });
  rl2.close();

  // Handle one continuous string of 8-char IDs
  const stripped = collected.join('');
  let ids;
  if (stripped.length % 8 === 0 && collected.length === 1) {
    ids = stripped.match(/.{8}/g);
  } else {
    ids = collected;
  }

  if (!ids || ids.length < 1 || ids.length > 20) {
    console.error(`\nExpected 1–20 Project IDs but got ${ids?.length ?? 0}. Please try again.\n`);
    process.exit(1);
  }

  console.log(`\nUsing ${ids.length} IDs: ${ids.join(', ')}`);
  console.log('\nLaunching Playwright...\n');

  const env = { ...process.env, VZ_PASSWORD: password.trim(), SPM_COUNT: String(ids.length) };
  ids.forEach((id, i) => { env[`SPMID${i + 1}`] = id; });

  const child = spawn('npx', ['playwright', 'test', 'Ericsson V1 SFP.spec.ts', '--headed', '--project=chromium'], {
    env,
    stdio: 'inherit',
    shell: true,
    cwd: path.join(__dirname, '..'),
  });

  child.on('exit', (code) => process.exit(code ?? 0));
}

main().catch((err) => { console.error(err); process.exit(1); });
