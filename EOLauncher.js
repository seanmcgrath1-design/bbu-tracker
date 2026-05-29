// launcher.js
// EO Bulk Order Launcher — interactive menu to select and run an EO clone script
// Usage: node launcher.js            (interactive menu)
//        node launcher.js DWDM       (direct by exact name)
//        node launcher.js 2          (direct by menu number)

const { spawn } = require('child_process');
const path = require('path');

// Add new EO types here as scripts are created
const EO_TYPES = [
  { label: 'DWDM',            script: 'Bulk EO Creation/run-DWDM.js' },
  { label: 'Ericsson V1 SFP', script: 'Bulk EO Creation/run-Ericsson V1 SFP.js' },
  { label: 'Ericsson V2 SFP', script: 'Bulk EO Creation/run-Ericsson V2 SFP.js' },
  { label: 'ProLabs V1 SFP',  script: 'Bulk EO Creation/run-prolabs-v1-sfp.js' },
];

function showMenu() {
  console.log('\n=== EO Bulk Order Launcher ===\n');
  EO_TYPES.forEach((t, i) => console.log(`  ${i + 1}. ${t.label}`));
  console.log('');
}

function resolveType(input) {
  const trimmed = input.trim();
  const num = parseInt(trimmed, 10);
  if (!isNaN(num) && num >= 1 && num <= EO_TYPES.length) {
    return EO_TYPES[num - 1];
  }
  return EO_TYPES.find(t => t.label.toLowerCase() === trimmed.toLowerCase()) ?? null;
}

function launch(eoType) {
  console.log(`\nStarting: ${eoType.label}\n`);
  const child = spawn('node', [eoType.script], {
    stdio: 'inherit',
    cwd: __dirname,
  });
  child.on('exit', (code) => process.exit(code ?? 0));
}

async function main() {
  const arg = process.argv[2];

  if (arg) {
    const eoType = resolveType(arg);
    if (!eoType) {
      console.error(`\nUnknown EO type: "${arg}"\n`);
      showMenu();
      process.exit(1);
    }
    launch(eoType);
    return;
  }

  showMenu();
  const input = await new Promise((resolve) => {
    process.stdout.write('Enter number or EO type name: ');
    process.stdin.setEncoding('utf8');
    process.stdin.resume();
    process.stdin.once('data', (data) => {
      process.stdin.pause();
      resolve(data.toString().trim());
    });
  });

  const eoType = resolveType(input);
  if (!eoType) {
    console.error(`\nUnknown EO type: "${input.trim()}". Enter a number or exact name.\n`);
    process.exit(1);
  }

  launch(eoType);
}

main().catch((err) => { console.error(err); process.exit(1); });
