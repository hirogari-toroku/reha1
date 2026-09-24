#!/usr/bin/env node
// 公開リポジトリへ個人情報・秘密情報が混入していないか調べる。
//   node tools/scan-sensitive.cjs            変更のあるファイル（未コミット＋originとの差分）
//   node tools/scan-sensitive.cjs --all      追跡中の全ファイル
// 見つかると終了コード1で止まる。誤検知は ALLOW_PATTERNS に理由付きで追加する。
const { execFileSync } = require('node:child_process');
const fs = require('node:fs');
const path = require('node:path');

const SKIP_FILES = [
  'tools/scan-sensitive.cjs',
  // 検査そのものを試すため、見本の秘密情報や氏名をあえて書いてあるファイル
  'tests/scan-sensitive.test.cjs',
  'package-lock.json'
];

// 拡張子で除外（画像・フォントなど）
const SKIP_EXTENSIONS = ['.png', '.jpg', '.jpeg', '.gif', '.webp', '.ico', '.pdf', '.woff', '.woff2', '.zip'];

const RULES = [
  {
    name: 'LINEチャネルアクセストークン',
    // LINEのチャネルアクセストークンは長いBase64風の文字列
    pattern: /\b[A-Za-z0-9+/]{100,}={0,2}\b/g,
    hint: 'トークンはスクリプトプロパティに保存し、コードには書かない'
  },
  {
    name: 'LINEチャネルシークレット/シークレット風の代入',
    pattern: /(channel_?secret|client_?secret|api[_-]?key|access_?token|refresh_?token|password|passwd)\s*[:=]\s*["'][^"'\s]{12,}["']/gi,
    hint: '秘密情報はスクリプトプロパティかGitHub Secretへ'
  },
  {
    name: 'Google OAuthクライアントシークレット',
    pattern: /GOCSPX-[A-Za-z0-9_-]{10,}/g,
    hint: '.clasprc.json の中身は絶対にコミットしない'
  },
  {
    name: 'GitHubトークン',
    pattern: /\bgh[pousr]_[A-Za-z0-9]{20,}\b/g,
    hint: 'tokenはKeychainとGitHub Secretだけに置く'
  },
  {
    name: '秘密鍵',
    pattern: /-----BEGIN [A-Z ]*PRIVATE KEY-----/g,
    hint: '鍵ファイルはリポジトリに置かない'
  },
  {
    name: '銀行口座らしい数値の並び（口座番号＋支店番号）',
    pattern: /(口座番号|支店番号|支店コード|銀行コード)\s*[:：=]?\s*["']?\d{3,8}/g,
    hint: '口座情報はスプレッドシートのみ。コードや手順書には書かない'
  },
  {
    name: '氏名らしい記述（さん・様つきの実名）',
    // 「〇〇さん」「〇〇様」。利用者様・スタッフ様のような一般語は除外する。
    pattern: /[一-龥]{2,5}(さん|様)(?![々])/g,
    hint: '実名は非公開の台帳へ。コード・コミット・docsには書かない',
    allow: [
      '利用者さん', '利用者様', 'スタッフさん', 'スタッフ様', '管理者さん', '管理者様',
      'ご家族さん', 'ご家族様', '家族さん', '家族様', '皆さん', '皆様', 'お客様', '担当者様', '関係者様',
      '本人様', 'ご本人様', '運営者さん', '代表者様', '規利用者様',
      // テストで使う架空の名前と、その部分一致（実在の利用者・スタッフではない）
      '山田太郎様', '山田太郎さん', '山田太郎様様', '太郎様', '太郎さん', '山田様', '山田さん',
      '鈴木花子様', '鈴木花子さん', '花子様', '花子さん',
      '佐藤一郎様', '佐藤一郎さん', '佐藤花子さん', '鈴木次郎さん', '田中一郎さん'
    ]
  },
  {
    name: 'LINEユーザーID',
    pattern: /\bU[0-9a-f]{32}\b/g,
    hint: '管理者IDは既存コードの1件のみ許容。新たなIDは書かない',
    allow: ['Uc21fa34144f5bc50c6e5324d5e4de344']
  }
];

function listFiles(all) {
  const git = args => execFileSync('git', args, { encoding: 'utf8' }).split('\n').filter(Boolean);
  if (all) return git(['ls-files']);
  const changed = new Set();
  git(['status', '--porcelain']).forEach(line => {
    const file = line.slice(3).trim();
    if (file) changed.add(file.replace(/^.* -> /, ''));
  });
  try {
    git(['diff', '--name-only', 'origin/main...HEAD']).forEach(file => changed.add(file));
  } catch (error) {
    // originが無い環境（初回clone前など）では未コミット分だけを見る
  }
  return Array.from(changed);
}

function scanFile(file) {
  const findings = [];
  const lines = fs.readFileSync(file, 'utf8').split('\n');
  lines.forEach((line, index) => {
    RULES.forEach(rule => {
      rule.pattern.lastIndex = 0;
      let match;
      while ((match = rule.pattern.exec(line)) !== null) {
        const hit = match[0];
        // 完全一致だけ許可する（部分一致にすると、似た実名を見逃す）
        if ((rule.allow || []).indexOf(hit) !== -1) continue;
        findings.push({ file: file, line: index + 1, rule: rule.name, hint: rule.hint, sample: hit.slice(0, 24) });
      }
    });
  });
  return findings;
}

function main() {
  const all = process.argv.includes('--all');
  const files = listFiles(all).filter(file => {
    if (SKIP_FILES.includes(file)) return false;
    if (SKIP_EXTENSIONS.includes(path.extname(file).toLowerCase())) return false;
    return fs.existsSync(file) && fs.statSync(file).isFile();
  });

  const findings = files.flatMap(scanFile);
  if (!findings.length) {
    console.log('個人情報・秘密情報は見つかりませんでした（' + files.length + 'ファイル確認）');
    return 0;
  }

  console.error('公開してはいけない内容が見つかりました：');
  findings.forEach(f => {
    console.error('  ' + f.file + ':' + f.line + '  [' + f.rule + '] ' + f.sample + ' … ' + f.hint);
  });
  console.error('\n誤検知の場合は tools/scan-sensitive.cjs の allow に理由を添えて追加してください。');
  return 1;
}

process.exit(main());
