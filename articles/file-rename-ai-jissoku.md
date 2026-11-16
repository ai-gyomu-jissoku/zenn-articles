---
title: "名前がバラバラなファイル100件を、GASで「日付_種別_連番」に一括リネームして台帳まで作った（コード全文）"
emoji: "🗂️"
type: "tech"
topics: ["gas", "googleappsscript", "業務効率化", "googledrive"]
published: true
---

## やったこと

Googleドライブの1つのフォルダに、名前の付け方がバラバラなファイル100件を入れて、Google Apps Script（GAS）で「日付_種別_連番」の統一名に一括リネームし、元ファイル名と対応する台帳スプレッドシートまで自動で作りました。

ファイル名の例は、こんな状態です。

- 請求書（清水設備）2025年10月3日.pdf
- seikyu_2025/9/26_山本運送.pdf
- IMG_3290.jpg
- gijiroku_20250915.pdf
- メモ書き.txt

これが、実行後は「20251003_請求書_004.pdf」のような形にそろいます。

結果は、対象100件すべてが事前に決めたルールどおりにリネームされ、台帳の全100行を正解表（元ファイル名と期待する新ファイル名の対応表）と照合して全件一致、不一致0でした。プログラム自体の実行時間は約51秒です（人の作業時間と比べた数字ではありません）。

種別ごとの件数は、請求書25件・見積書20件・契約書15件・議事録15件・領収書10件・画像10件・その他5件です。

動画版（実行画面つき）はこちらです。

https://youtu.be/S2sLI4VCenk

## 使い方

1. Google Apps Script（script.google.com）で新しいプロジェクトを作り、下のコードを全文貼ります
2. コード冒頭の FOLDER_NAME に、整理したいフォルダの名前を入れます（フォルダ名で探す作りです）
3. 関数 renameFilesAndBuildLedger を実行します。初回はドライブとスプレッドシートへのアクセス承認が出ます
4. フォルダ内のファイルがリネームされ、マイドライブに「ファイル台帳」というスプレッドシートが作られます

注意：リネームは元に戻すのが手作業になります。最初は、複製したテスト用フォルダで試してください。台帳に元ファイル名とURLが残るので、万一のときも「どれがどれになったか」はたどれます。

## コード全文

実測で使ったコードそのままです。公開用に、対象フォルダ名だけを設定欄に切り出しています（処理の中身は変えていません）。

```javascript
// フォルダ内のファイルを「日付_種別_連番」に一括リネームし、台帳スプレッドシートを自動作成する
// 【設定】ここに、整理したいフォルダの名前を入れてください（フォルダ名で探します）
const FOLDER_NAME = '整理したいフォルダの名前';
const LEDGER_NAME = 'ファイル台帳';

function parseDateFromName(name, fallbackDate) {
  let m = name.match(/(20\d{2})(\d{2})(\d{2})/);
  if (m) return m[1] + m[2] + m[3];
  m = name.match(/(20\d{2})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if (m) return m[1] + ('0' + m[2]).slice(-2) + ('0' + m[3]).slice(-2);
  m = name.match(/(20\d{2})年(\d{1,2})月(\d{1,2})日/);
  if (m) return m[1] + ('0' + m[2]).slice(-2) + ('0' + m[3]).slice(-2);
  return Utilities.formatDate(fallbackDate, 'Asia/Tokyo', 'yyyyMMdd');
}

function detectCategory(name) {
  if (name.indexOf('契約') >= 0 || name.indexOf('keiyaku') >= 0) return '契約書';
  if (name.indexOf('請求') >= 0 || name.indexOf('seikyu') >= 0) return '請求書';
  if (name.indexOf('見積') >= 0 || name.indexOf('mitsumori') >= 0) return '見積書';
  if (name.indexOf('議事録') >= 0 || name.indexOf('会議メモ') >= 0 || name.indexOf('gijiroku') >= 0) return '議事録';
  if (name.indexOf('領収') >= 0 || name.indexOf('receipt') >= 0) return '領収書';
  if (name.indexOf('IMG_') >= 0 || name.indexOf('写真') >= 0 || name.indexOf('スクリーンショット') >= 0) return '画像';
  if (/\.(jpg|png)$/i.test(name)) return '画像';
  return 'その他';
}

function renameFilesAndBuildLedger() {
  const folders = DriveApp.getFoldersByName(FOLDER_NAME);
  if (!folders.hasNext()) { Logger.log('フォルダが見つかりません'); return; }
  const folder = folders.next();
  const it = folder.getFiles();
  const items = [];
  while (it.hasNext()) {
    const file = it.next();
    const oldName = file.getName();
    const ext = oldName.indexOf('.') >= 0 ? oldName.split('.').pop().toLowerCase() : '';
    items.push({
      file: file, oldName: oldName, ext: ext,
      category: detectCategory(oldName),
      date: parseDateFromName(oldName, file.getDateCreated()),
      url: file.getUrl()
    });
  }
  items.sort(function(a, b) {
    if (a.category !== b.category) return a.category < b.category ? -1 : 1;
    if (a.date !== b.date) return a.date < b.date ? -1 : 1;
    return a.oldName < b.oldName ? -1 : 1;
  });
  const counters = {};
  const rows = [];
  items.forEach(function(item, idx) {
    counters[item.category] = (counters[item.category] || 0) + 1;
    const num = ('00' + counters[item.category]).slice(-3);
    const newName = item.date + '_' + item.category + '_' + num + (item.ext ? '.' + item.ext : '');
    item.file.setName(newName);
    rows.push([idx + 1, item.category, item.date, newName, item.oldName, item.url]);
  });
  const ss = SpreadsheetApp.create(LEDGER_NAME);
  const sheet = ss.getActiveSheet();
  sheet.appendRow(['No.', '種別', '日付', '新ファイル名', '元ファイル名', 'URL']);
  rows.forEach(function(r) { sheet.appendRow(r); });
  Logger.log('処理件数: ' + items.length);
  Object.keys(counters).forEach(function(k) { Logger.log(k + ': ' + counters[k] + '件'); });
  Logger.log('台帳URL: ' + ss.getUrl());
}
```

## 仕組み

- 日付：ファイル名から取り出します。対応する書き方は4形式（20251109／2025-11-05／2025/9/26／2025年10月7日）で、YYYYMMDDの8桁に統一します
- 種別：ファイル名に含まれるキーワードで判定します。契約書／請求書／見積書／議事録（会議メモを含む）／領収書／画像（IMG_・写真・スクリーンショット・jpg/png）の順に判定し、どれにも当たらないものは「その他」です
- 連番：種別ごとに、日付の昇順（同じ日は元ファイル名の順）で001から振ります
- 台帳：No.／種別／日付／新ファイル名／元ファイル名／URL の6列で作ります。元ファイル名とURLを残すのがポイントで、変換の検証と探し直しに使えます

## 正直な制限が2つある

1. 名前に日付が無いファイルは、ファイルの作成日で処理されます。今回の100件では32件が該当しました（例：IMG_3290.jpg → 20261116_画像_008.jpg）。実務では「取り込んだ日」と「実際に作られた日」がずれるので、日付なしの行は台帳で人が確認する前提です
2. キーワードで種別を判定できないファイルは「その他」に入ります。今回は5件（メモ書き.txt、無題のドキュメント.docx など）。名前に手がかりが無いファイルは自動では仕分けきれないので、「その他」をゼロにするより、台帳で最終確認する前提で作るのが現実的です

## 実測の条件と注意

- ファイル名・会社名はすべて実験用のダミーです
- 試行は1回だけ（n=1）なので、条件が変われば結果も変わります
- 想定していない日付の書き方（例：25.11.09のような形式）は読み取れず、作成日に回ります。使う前に、自分のフォルダに混ざっている書き方を確認してください

命名ルールと台帳管理のチェックリスト（無料）はnoteに置いてあります。

https://note.com/gas_recipe/n/n688ea10c57b9
