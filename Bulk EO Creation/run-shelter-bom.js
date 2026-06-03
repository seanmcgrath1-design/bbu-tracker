// run-shelter-bom.js
// Shelter BOM — generates and downloads the RFDS BOM for one or more Fuze projects
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

function runProject(projectId, password) {
  return new Promise((resolve) => {
    const env = {
      ...process.env,
      VZ_PASSWORD: password,
      FUZE_PROJECT_ID: projectId,
      // FORECAST_DATE only set if provided as env override; otherwise the spec reads it from the page
      ...(process.env.FORECAST_DATE ? { FORECAST_DATE: process.env.FORECAST_DATE } : {}),
    };

    const child = spawn('npx', ['playwright', 'test', '--grep', 'Shelter BOM', '--headed', '--project=chromium'], {
      env,
      stdio: 'inherit',
      shell: true,
      cwd: path.join(__dirname, '..'),
    });

    child.on('exit', (code) => resolve(code ?? 0));
  });
}

async function main() {
  console.log('\n=== Shelter BOM Tool ===\n');

  const password = await promptPassword('Verizon Password: ');

  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  const input = await new Promise((resolve) =>
    rl.question('\nFuze Project Numbers (up to 5, comma-separated): ', resolve)
  );
  rl.close();

  // Parse: split on commas/spaces, strip blanks, deduplicate, cap at 5
  const projectIds = [...new Set(
    input.split(/[\s,]+/).map(s => s.trim()).filter(Boolean)
  )].slice(0, 5);

  if (projectIds.length === 0) {
    console.error('\nNo project IDs provided. Please try again.\n');
    process.exit(1);
  }

  console.log(`\nProjects to process (${projectIds.length}): ${projectIds.join(', ')}\n`);

  const results = [];

  for (let i = 0; i < projectIds.length; i++) {
    const projectId = projectIds[i];
    console.log(`${'='.repeat(50)}`);
    console.log(`  [${i + 1}/${projectIds.length}] Project: ${projectId}`);
    console.log(`${'='.repeat(50)}\n`);

    const exitCode = await runProject(projectId, password.trim());
    results.push({ projectId, passed: exitCode === 0 });

    console.log(`\n  → ${projectId}: ${exitCode === 0 ? 'PASSED ✓' : 'FAILED ✗'}\n`);
  }

  // Summary
  console.log(`${'='.repeat(50)}`);
  console.log('  SUMMARY');
  console.log(`${'='.repeat(50)}`);
  for (const { projectId, passed } of results) {
    console.log(`  ${projectId}: ${passed ? 'PASSED ✓' : 'FAILED ✗'}`);
  }
  console.log(`${'='.repeat(50)}\n`);

  const failCount = results.filter(r => !r.passed).length;
  process.exit(failCount > 0 ? 1 : 0);
}

main().catch((err) => { console.error(err); process.exit(1); });
