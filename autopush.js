const { execSync } = require('child_process');
const path = require('path');

// Menggunakan git dari path MinGit atau environment
const gitExe = 'C:\\Users\\user\\AppData\\Local\\Programs\\MinGit\\cmd\\git.exe';

function run(cmd) {
  return execSync(cmd, {
    cwd: __dirname,
    stdio: 'inherit',
    env: {
      ...process.env,
      PATH: 'C:\\Users\\user\\AppData\\Local\\Programs\\MinGit\\cmd;' + (process.env.PATH || '')
    }
  });
}

try {
  console.log('\n======================================================');
  console.log('🚀 AUTO-PUSH GITHUB: https://github.com/bimcuk/bayi-app');
  console.log('======================================================\n');

  console.log('🔍 Memeriksa status perubahan...');
  const status = execSync(`"${gitExe}" status --porcelain`, {
    cwd: __dirname,
    env: { ...process.env, PATH: 'C:\\Users\\user\\AppData\\Local\\Programs\\MinGit\\cmd;' + (process.env.PATH || '') }
  }).toString().trim();

  let unpushed = '';
  try {
    unpushed = execSync(`"${gitExe}" cherry -v`, {
      cwd: __dirname,
      env: { ...process.env, PATH: 'C:\\Users\\user\\AppData\\Local\\Programs\\MinGit\\cmd;' + (process.env.PATH || '') }
    }).toString().trim();
  } catch (e) {}

  if (!status && !unpushed) {
    console.log('✅ Tidak ada file yang berubah & semua commit sudah ter-push ke GitHub!');
    console.log('------------------------------------------------------\n');
    process.exit(0);
  }

  if (status) {
    console.log('📁 Staging semua file...');
    run(`"${gitExe}" add -A`);

    const now = new Date();
    const dateStr = now.toISOString().replace('T', ' ').substring(0, 19);
    const userMsg = process.argv.slice(2).join(' ').trim();
    const commitMsg = userMsg || `Update aplikasi - ${dateStr}`;

    console.log(`\n💬 Membuat commit: "${commitMsg}"...`);
    run(`"${gitExe}" commit -m "${commitMsg}"`);
  } else {
    console.log('ℹ️ Ada commit lokal yang belum di-push ke GitHub.');
  }

  console.log('\n☁️ Mengirim perubahan ke branch main...');
  run(`"${gitExe}" push origin main`);

  console.log('\n======================================================');
  console.log('🎉 SUKSES! Perubahan telah ter-push ke GitHub:');
  console.log('👉 https://github.com/bimcuk/bayi-app');
  console.log('======================================================\n');
} catch (err) {
  console.error('\n❌ Terjadi kendala saat push ke GitHub:');
  console.error(err.message || err);
  console.log('\n💡 CATATAN OTENTIKASI GITHUB:');
  console.log('Jika GitHub meminta autentikasi, kamu bisa:');
  console.log('1. Gunakan GitHub Personal Access Token (PAT) sebagai password.');
  console.log('2. Atau jalankan "git push origin main" di terminal untuk login interaktif.');
  process.exit(1);
}
