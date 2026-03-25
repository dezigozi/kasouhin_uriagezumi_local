const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const appDir = __dirname;
const parentDir = path.dirname(appDir);
const icoPath = path.join(appDir, 'app-icon.ico');
const shortcutPath = path.join(parentDir, 'リース実績レポート.lnk');
const batPath = path.join(parentDir, 'リース実績レポート.bat');

function createShortcut() {
  const psScript = `
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ws = New-Object -ComObject WScript.Shell
$sc = $ws.CreateShortcut('${shortcutPath}')
$sc.TargetPath = '${batPath}'
$sc.WorkingDirectory = '${parentDir}'
$sc.IconLocation = '${icoPath},0'
$sc.Description = 'Special Sales Report'
$sc.Save()
Write-Host 'OK'
`;

  const tempPs = path.join(appDir, '_mklink.ps1');
  fs.writeFileSync(tempPs, '\ufeff' + psScript, 'utf16le');

  try {
    const out = execSync(
      `powershell -NoProfile -ExecutionPolicy Bypass -File "${tempPs}"`,
      { encoding: 'utf8' }
    );
    console.log('PowerShell:', out.trim());
    console.log(`Shortcut created: ${shortcutPath}`);
  } finally {
    try { fs.unlinkSync(tempPs); } catch {}
  }
}

try {
  if (!fs.existsSync(batPath)) {
    console.error(`BAT not found: ${batPath}`);
    process.exit(1);
  }
  if (!fs.existsSync(icoPath)) {
    console.error(`ICO not found: ${icoPath}`);
    process.exit(1);
  }
  createShortcut();
  console.log('Done!');
} catch (err) {
  console.error('Error:', err.message);
}
