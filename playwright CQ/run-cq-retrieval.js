// run-cq-retrieval.js
// CQ Retrieval — downloads one CQ (by Fuze Project ID) to Downloads, then moves it
// into the synced Google Drive folder bbu-tracker/CQ.
// Usage: npm run cq-retrieval
// For automated runs, set VZ_PASSWORD as a Windows environment variable to skip the prompt.

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

async function prompt(question) {
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  const answer = await new Promise((resolve) => rl.question(question, resolve));
  rl.close();
  return answer.trim();
}

async function main() {
  console.log('\n=== CQ Retrieval ===\n');

  let password = process.env.VZ_PASSWORD;
  if (password) {
    console.log('Using VZ_PASSWORD from environment.\n');
  } else {
    password = await promptPassword('Verizon Password: ');
  }

  const fuzeId = await prompt('Fuze Project ID: ');
  if (!fuzeId) {
    console.error('Fuze Project ID is required.');
    process.exit(1);
  }

  console.log(`\nFuze ID: ${fuzeId}`);
  console.log('\nLaunching Playwright...\n');

  const child = spawn(
    'npx',
    ['playwright', 'test', '--grep', 'CQ Retrieval', '--headed', '--project=chromium'],
    {
      env: { ...process.env, VZ_PASSWORD: password.trim(), FUZE_ID: fuzeId },
      stdio: 'inherit',
      shell: true,
      cwd: path.join(__dirname, '..'),
    }
  );

  child.on('exit', (code) => process.exit(code ?? 0));
}

main().catch((err) => { console.error(err); process.exit(1); });
