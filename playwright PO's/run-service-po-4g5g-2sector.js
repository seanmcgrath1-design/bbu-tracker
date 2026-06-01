// run-service-po-4g5g-2sector.js
// Service PO — 4G/5G 2 Sector
// Usage: node "playwright-tests/run-service-po-4g5g-2sector.js"

const readline = require('readline');
const { spawn } = require('child_process');
const path = require('path');

function promptPassword(prompt) {
  return new Promise((resolve) => {
    process.stdout.write(prompt);
    process.stdin.setRawMode(true);
    process.stdin.resume();
    process.stdin.setEncoding('utf8');
    let password = '';
    process.stdin.on('data', function handler(char) {
      if (char === '\r' || char === '\n' || char === '') {
        process.stdin.setRawMode(false);
        process.stdin.pause();
        process.stdin.removeListener('data', handler);
        process.stdout.write('\n');
        resolve(password);
      } else if (char === '') {
        process.exit();
      } else if (char === '' || char === '\b') {
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

async function prompt(question) {
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  const answer = await new Promise((resolve) => rl.question(question, resolve));
  rl.close();
  return answer.trim();
}

async function main() {
  console.log('\n=== Service PO — 4G/5G 2 Sector ===\n');

  const password    = await promptPassword('Verizon Password: ');
  const projectNum  = await prompt('Project Number: ');
  const dueDate     = await prompt('Requested Due Date (MM/DD/YYYY): ');

  console.log(`\nProject: ${projectNum}  |  Due Date: ${dueDate}`);
  console.log('\nLaunching Playwright...\n');

  const env = {
    ...process.env,
    VZ_PASSWORD:  password,
    PROJECT_NUM:  projectNum,
    DUE_DATE:     dueDate,
  };

  const child = spawn(
    'npx',
    ['playwright', 'test', 'service-po-4g5g-2sector.spec.ts', '--headed', '--project=chromium'],
    { env, stdio: 'inherit', shell: true, cwd: path.join(__dirname, '..') }
  );

  child.on('exit', (code) => process.exit(code ?? 0));
}

main().catch((err) => { console.error(err); process.exit(1); });
