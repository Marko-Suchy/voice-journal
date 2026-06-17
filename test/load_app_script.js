const fs = require('fs');
const path = require('path');
function loadAppScript() {
  const srcDir = path.join(__dirname, '..', 'src');
  const source = fs.readdirSync(srcDir)
    .filter(function(file) {
      return /\.js$/.test(file) && file !== 'Index.html';
    })
    .sort()
    .map(function(file) {
      return fs.readFileSync(path.join(srcDir, file), 'utf8');
    })
    .join('\n');
  const module = { exports: {} };
  const run = new Function('module', 'exports', 'require', source);
  run(module, module.exports, require);
  return module.exports;
}

module.exports = loadAppScript();
