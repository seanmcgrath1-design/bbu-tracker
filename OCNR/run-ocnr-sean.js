// run-ocnr-sean.js
// OCNR-Sean — downloads and date-stamps the OCNR report
// Usage: npm run ocnr-sean
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

async function main() {
  console.log('\n=== OCNR-Sean ===\n');

  let password = process.env.VZ_PASSWORD;

  if (password) {
    console.log('Using VZ_PASSWORD from environment.\n');
  } else {
    password = await promptPassword('Verizon Password: ');
  }

  const child = spawn(
    'npx',
    ['playwright', 'test', '--grep', 'OCNR-Sean', '--headed', '--project=chromium'],
    {
      env: { ...process.env, VZ_PASSWORD: password.trim() },
      stdio: 'inherit',
      shell: true,
      cwd: path.join(__dirname, '..'),
    }
  );

  child.on('exit', (code) => process.exit(code ?? 0));
}

main().catch((err) => { console.error(err); process.exit(1); });
