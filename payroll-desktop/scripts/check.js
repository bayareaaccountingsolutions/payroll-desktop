// Pre-release checks: every JS / GS file parses, versions are consistent, publish target is set.
const fs = require('fs'), path = require('path'), vm = require('vm');
const root = path.join(__dirname, '..');
let bad = 0;
const ok = (c, m) => { console.log((c ? '✔ ' : '✖ ') + m); if (!c) bad++; };
const walk = d => fs.readdirSync(d, { withFileTypes: true }).flatMap(e => e.isDirectory() ? walk(path.join(d, e.name)) : [path.join(d, e.name)]);
walk(path.join(root, 'src')).concat(walk(path.join(root, 'apps-script'))).filter(f => /\.(js|gs)$/.test(f)).forEach(f => {
  try { new vm.Script(fs.readFileSync(f, 'utf8'), { filename: f }); ok(true, 'syntax ' + path.relative(root, f)); }
  catch (e) { ok(false, 'syntax ' + path.relative(root, f) + ': ' + e.message); }
});
const pkg = require(path.join(root, 'package.json'));
ok(/^\d+\.\d+\.\d+$/.test(pkg.version), 'app version ' + pkg.version);
const pub = pkg.build.publish[0];
ok(pub.owner && pub.owner !== 'YOUR-GITHUB-USERNAME', 'GitHub owner set in package.json ▸ build.publish (needed for auto-update)');
const api = fs.readFileSync(path.join(root, 'apps-script', 'DesktopApi.gs'), 'utf8');
ok(/DESKTOP_API_VERSION\s*=\s*'[\d.]+'/.test(api), 'DesktopApi.gs has an API version');
if (process.env.GITHUB_REF_NAME && /^v/.test(process.env.GITHUB_REF_NAME))
  ok(process.env.GITHUB_REF_NAME === 'v' + pkg.version, `tag ${process.env.GITHUB_REF_NAME} matches package.json version ${pkg.version}`);
if (bad) { console.error(`\n${bad} check(s) failed.`); process.exit(1); }
console.log('\nAll checks passed.');
