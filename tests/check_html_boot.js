const fs = require('fs');
const path = require('path');
const assert = require('assert');
const { execSync } = require('child_process');
const { JSDOM, VirtualConsole } = require('jsdom');

const repoRoot = path.resolve(__dirname, '..');
const htmlPath = path.resolve(repoRoot, 'Tubex.html');
const updateScriptPath = path.resolve(repoRoot, 'Scripts', 'update_html.py');

let check1Pass = false;
let check2Pass = false;
let check3Pass = false;
const failDetails = [];

// ============================================================================
// CHECK 2: Run update_html.py first, then check DASH_DATA in freshly generated Tubex.html
// Task 7 check: petYest must be 15755 for latest date 2026-09-16 (or 17400 for 2026-09-21),
// and tubeYest must be unchanged.
// ============================================================================
try {
  // Run update_html.py
  execSync(`python "${updateScriptPath}"`, { cwd: repoRoot, stdio: 'pipe' });

  const freshHtml = fs.readFileSync(htmlPath, 'utf8');
  const dashMatch = freshHtml.match(/const DASH_DATA = (\{[\s\S]*?\});\s*\/\* DATA_END \*\//);
  if (!dashMatch) {
    throw new Error('DASH_DATA block not found in Tubex.html');
  }
  const dashData = JSON.parse(dashMatch[1]);
  const kpi = dashData.kpi || {};
  const latestDate = dashData.latestDate || {};

  // Check based on latest date in workbook
  if (latestDate.iso === '2026-09-16') {
    assert.strictEqual(kpi.petYest, 15755, `Expected petYest=15755 on 2026-09-16, got ${kpi.petYest}`);
    assert.strictEqual(kpi.tubeYest, 24990, `Expected tubeYest=24990 on 2026-09-16, got ${kpi.tubeYest}`);
  } else if (latestDate.iso === '2026-09-21') {
    assert.strictEqual(kpi.petYest, 17400, `Expected petYest=17400 on 2026-09-21, got ${kpi.petYest}`);
    assert.strictEqual(kpi.tubeYest, 25942, `Expected tubeYest=25942 on 2026-09-21, got ${kpi.tubeYest}`);
  } else {
    assert(kpi.petYest > 0, `Expected petYest > 0 on ${latestDate.iso}, got ${kpi.petYest}`);
    assert(kpi.tubeYest > 0, `Expected tubeYest > 0 on ${latestDate.iso}, got ${kpi.tubeYest}`);
  }

  // Also confirm historical 16-Sep row in PRODUCTION_LOG_DATA matches 15755
  const prodLogMatch = freshHtml.match(/const PRODUCTION_LOG_DATA = (\{[\s\S]*?\});\s*\/\* PRODLOG_END \*\//);
  if (prodLogMatch) {
    const prodLogData = JSON.parse(prodLogMatch[1]);
    const pet16Rows = (prodLogData.rows || []).filter(r =>
      String(r.date).includes('16') && (String(r.machine).toUpperCase().startsWith('PF') || String(r.machine).toUpperCase().startsWith('PET'))
    );
    const pet16Total = pet16Rows.reduce((sum, r) => sum + (r.good || 0), 0);
    assert.strictEqual(pet16Total, 15755, `Expected 16-Sep PET production to sum to 15755, got ${pet16Total}`);
  }

  check2Pass = true;
} catch (err) {
  failDetails.push(`Check 2 (Task 7) Failed: ${err.message}`);
}

// ============================================================================
// CHECK 1: Load Tubex.html with runScripts: 'dangerously', stub navigator.serviceWorker.register,
// and confirm register was called with './sw.js' with no script errors (Task 1 check).
// ============================================================================
const htmlContent = fs.readFileSync(htmlPath, 'utf8');

const vc1 = new VirtualConsole();
const scriptErrors1 = [];
vc1.on('jsdomError', (err) => scriptErrors1.push(err));

let registeredPath1 = null;

try {
  const dom1 = new JSDOM(htmlContent, {
    runScripts: 'dangerously',
    url: 'http://localhost/',
    virtualConsole: vc1,
    beforeParse(window) {
      Object.defineProperty(window.navigator, 'serviceWorker', {
        value: {
          register: (script, opts) => {
            registeredPath1 = script;
            return Promise.resolve({});
          },
          addEventListener: () => {}
        },
        configurable: true
      });
    }
  });

  assert.strictEqual(scriptErrors1.length, 0, `Script errors occurred: ${scriptErrors1.map(e => e.message || e).join('; ')}`);
  assert.strictEqual(registeredPath1, './sw.js', `Expected serviceWorker.register('./sw.js'), got '${registeredPath1}'`);
  check1Pass = true;
} catch (err) {
  failDetails.push(`Check 1 (Task 1) Failed: ${err.message}`);
}

// ============================================================================
// CHECK 3: Append <b id=probeN></b> to the names of a product, a customer, a machine,
// and an FG item, reload the page, click through every tab, and confirm NO probe id appears
// in the DOM (Task 10 check).
// ============================================================================
try {
  // Inject probes into raw data blocks of the HTML string
  let probedHtml = htmlContent;

  // 1. Product name probe: append <b id=probeProduct></b>
  probedHtml = probedHtml.replace(
    /name:\s*['"]HELLO HAIR COLOR['"]/,
    'name: "HELLO HAIR COLOR<b id=\\"probeProduct\\"></b>"'
  );
  probedHtml = probedHtml.replace(
    /"product":\s*"HELLO HAIR COLOR"/,
    '"product": "HELLO HAIR COLOR<b id=\\"probeProduct\\"></b>"'
  );

  // 2. Customer name probe: append <b id=probeCustomer></b>
  probedHtml = probedHtml.replace(
    /customer:\s*['"]Golden Pearl Cosmetics \(PVT\) LTD['"]/,
    'customer: "Golden Pearl Cosmetics (PVT) LTD<b id=\\"probeCustomer\\"></b>"'
  );
  probedHtml = probedHtml.replace(
    /"customer":\s*"Golden Pearl Cosmetics \(PVT\) LTD"/,
    '"customer": "Golden Pearl Cosmetics (PVT) LTD<b id=\\"probeCustomer\\"></b>"'
  );

  // 3. Machine name probe: append <b id=probeMachine></b>
  probedHtml = probedHtml.replace(
    /"machine":\s*"Printing-04"/,
    '"machine": "Printing-04<b id=\\"probeMachine\\"></b>"'
  );
  probedHtml = probedHtml.replace(
    /"machine":\s*"PF Machine"/,
    '"machine": "PF Machine<b id=\\"probeMachine\\"></b>"'
  );

  // 4. FG item probe: append <b id=probeFG></b>
  probedHtml = probedHtml.replace(
    /"product":\s*"PET BOTTLE SMALL \(120ML\) YELLOW"/,
    '"product": "PET BOTTLE SMALL (120ML) YELLOW<b id=\\"probeFG\\"></b>"'
  );

  const vc3 = new VirtualConsole();
  const scriptErrors3 = [];
  vc3.on('jsdomError', (err) => scriptErrors3.push(err));

  const probedDom = new JSDOM(probedHtml, {
    runScripts: 'dangerously',
    url: 'http://localhost/',
    virtualConsole: vc3,
    beforeParse(window) {
      Object.defineProperty(window.navigator, 'serviceWorker', {
        value: {
          register: () => Promise.resolve({}),
          addEventListener: () => {}
        },
        configurable: true
      });
    }
  });

  const win = probedDom.window;
  const doc = win.document;

  // Click through every tab
  const tabs = ['dashboard', 'customer_report', 'calc', 'prodlog', 'fgstock', 'inventory'];
  tabs.forEach(tabName => {
    const btn = doc.querySelector(`.tab[onclick*="${tabName}"]`) || doc.createElement('button');
    if (typeof win.switchTab === 'function') {
      win.switchTab(tabName, btn);
    }
  });

  // Also trigger filter / render functions across components
  if (typeof win.updateGrandTotal === 'function') win.updateGrandTotal();
  if (typeof win.filterProducts === 'function') win.filterProducts();
  if (typeof win.filterProdLog === 'function') win.filterProdLog();
  if (typeof win.filterFGStock === 'function') win.filterFGStock();
  if (typeof win.filterInventory === 'function') win.filterInventory();
  if (typeof win.renderNativeCustMain === 'function') win.renderNativeCustMain();

  // Confirm NO probe id appears in the DOM
  const foundProbes = doc.querySelectorAll('#probeProduct, #probeCustomer, #probeMachine, #probeFG, [id^="probe"]');
  if (foundProbes.length > 0) {
    const ids = Array.from(foundProbes).map(el => el.id || el.tagName);
    throw new Error(`Probe element(s) found in DOM: ${ids.join(', ')}`);
  }

  check3Pass = true;
} catch (err) {
  failDetails.push(`Check 3 (Task 10) Failed: ${err.message}`);
}

// ============================================================================
// OUTPUT RESULTS
// "Print only PASS/FAIL for each of the 3 checks. Full output only if something fails."
// ============================================================================
console.log(`Check 1 (Task 1 - Service Worker registration & zero script errors): ${check1Pass ? 'PASS' : 'FAIL'}`);
console.log(`Check 2 (Task 7 - DASH_DATA petYest & tubeYest): ${check2Pass ? 'PASS' : 'FAIL'}`);
console.log(`Check 3 (Task 10 - Escape dynamic text & probe check across all tabs): ${check3Pass ? 'PASS' : 'FAIL'}`);

if (!check1Pass || !check2Pass || !check3Pass) {
  console.error('\n--- Failure Details ---');
  failDetails.forEach(d => console.error(d));
  process.exit(1);
} else {
  process.exit(0);
}
