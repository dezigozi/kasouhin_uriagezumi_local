const { execSync } = require('child_process');
const path = require('path');
process.chdir(__dirname);
execSync('npx electron .', { stdio: 'inherit' });
