const fs = require('fs');
const path = require('path');
const cacheDir = path.join(process.env.APPDATA, 'lease-report-app', 'excel-cache');
console.log('Cache dir:', cacheDir);
if (fs.existsSync(cacheDir)) {
  const files = fs.readdirSync(cacheDir);
  files.forEach(f => {
    const fp = path.join(cacheDir, f);
    fs.unlinkSync(fp);
    console.log('Deleted:', f);
  });
  console.log(`Cleared ${files.length} cache files`);
} else {
  console.log('Cache dir not found');
}
