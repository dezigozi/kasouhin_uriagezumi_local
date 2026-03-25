const fs = require('fs');
const path = require('path');

const srcRoot = path.resolve(__dirname, '..');
const dstRoot = path.join(srcRoot, 'バックアップ');

function copyDir(src, dst, skipDirs) {
  if (!fs.existsSync(dst)) fs.mkdirSync(dst, { recursive: true });
  for (const entry of fs.readdirSync(src, { withFileTypes: true })) {
    const srcPath = path.join(src, entry.name);
    const dstPath = path.join(dst, entry.name);
    if (entry.isDirectory()) {
      if (skipDirs.includes(entry.name)) continue;
      copyDir(srcPath, dstPath, skipDirs);
    } else {
      fs.copyFileSync(srcPath, dstPath);
    }
  }
}

const skipDirs = ['node_modules', 'バックアップ', 'バックアップ_20260220', '20260220', '.git'];

console.log('Source:', srcRoot);
console.log('Destination:', dstRoot);

try {
  copyDir(srcRoot, dstRoot, skipDirs);
  console.log('Backup complete!');
  const count = (function countFiles(dir) {
    let n = 0;
    for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
      if (e.isDirectory()) n += countFiles(path.join(dir, e.name));
      else n++;
    }
    return n;
  })(dstRoot);
  console.log('Total files copied:', count);
} catch (err) {
  console.error('Error:', err.message);
}
