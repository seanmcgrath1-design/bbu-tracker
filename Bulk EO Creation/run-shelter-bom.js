// run-shelter-bom.js
// Shelter BOM — generates and downloads the RFDS BOM for a Fuze project
// Usage: npm run shelter-bom

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

async function main() {
  console.log('\n=== Shelter BOM Tool ===\n');

  const password = await promptPassword('Verizon Password: ');

  const rl2 = readline.createInterface({ input: process.stdin, output: process.stdout });
  const projectId = await new Promise((resolve) => rl2.question('\nFuze Project Number: ', resolve));
  rl2.close();

  if (!projectId.trim()) {
    console.error('\nNo project ID provided. Please try again.\n');
    process.exit(1);
  }

  const rl3 = readline.createInterface({ input: process.stdin, output: process.stdout });
  const dateInput = await new Promise((resolve) => rl3.question('\nForecast Date (MM/DD/YYYY, or Enter to auto-calculate): ', resolve));
  rl3.close();
  const forecastDate = dateInput.trim();

  if (forecastDate) {
    console.log(`\nUsing provided Forecast Date: ${forecastDate}`);
  } else {
    console.log('\nForecast Date will be auto-calculated after login.');
  }

  console.log(`\nProject: ${projectId.trim()}`);
  console.log('\nLaunching Playwright...\n');

  const env = {
    ...process.env,
    VZ_PASSWORD: password.trim(),
    FUZE_PROJECT_ID: projectId.trim(),
    FORECAST_DATE: forecastDate,
  };

  const child = spawn('npx', ['playwright', 'test', '--grep', 'Shelter BOM', '--headed', '--project=chromium'], {
    env,
    stdio: 'inherit',
    shell: true,
    cwd: path.join(__dirname, '..'),
  });

  child.on('exit', (code) => process.exit(code ?? 0));
}

main().catch((err) => { console.error(err); process.exit(1); });
