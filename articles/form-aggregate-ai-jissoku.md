---
title: "申請フォームの回答60件を、GASで集計してグラフ化し、担当者へメールまで送った（コード全文）"
emoji: "📋"
type: "tech"
topics: ["gas", "googleappsscript", "業務効率化", "googleフォーム"]
published: true
---

## やったこと

Googleフォームの回答が溜まったスプレッドシート（申請60件）を、Google Apps Script（GAS）で種別ごとに集計し、集計表と棒グラフを作り、集計結果を担当者へメールで送るまでを自動化しました。

回答データはこんな状態です。金額欄の表記がそろっていません。

- 数値のまま：27600
- カンマ入り文字列：「49,600」
- 「円」付き：「3500円」
- 全角数字：「１２８００」
- 金額が未定で「未定」と入っているもの、空欄のもの

結果は、全60件の集計が事前に固定した正解表と全項目一致しました。合計金額は1,503,800円（金額が読み取れた57件分）、未対応は24件です。検証は、集計シートの全セルとグラフの有無を読み取り専用の別スクリプトで正解表と照合し、一致24件・不一致0件でした。送信メールも実際に届き、件名と本文の数字が正解表と一致しています。プログラム自体の実行時間は約7秒です（人の作業時間と比べた数字ではありません）。

種別ごとの内訳は、備品購入16件・755,900円、交通費14件・97,200円、会議費10件・165,600円、研修費9件・316,000円、印刷費11件・169,100円です。

動画版（実行画面つき）はこちらです。

https://youtu.be/G8ovAzHkOz8

## 使い方

1. Googleフォームの回答が入ったスプレッドシートを用意します（この記事の検証では、回答シートの列が「タイムスタンプ／申請者名／種別／金額／申請日／状態／備考」の並びです）
2. Google Apps Script（script.google.com）で新しいプロジェクトを作り、下のコードを全文貼ります
3. コード冒頭の設定を自分の環境に合わせます。BOOK_NAME は回答スプレッドシートの名前、ANSWER_SHEET は回答シートの名前、RECIPIENT は集計メールを送る担当者のメールアドレスです（名前で探す作りで、IDやURLは使いません）
4. 関数 aggregateApplications を実行します。初回はスプレッドシートとメール送信へのアクセス承認が出ます
5. 同じスプレッドシートに「集計」シートが作られ、集計表と棒グラフが入り、担当者へ集計メールが届きます

注意：実行するたびにメールが1通送られます。試すときは、複製したテスト用のスプレッドシートと自分のメールアドレスで試してください。

## コード全文

実測に使ったコードと同一のものです。冒頭の設定3点（対象スプレッドシート名の例示、回答シート名、宛先メールアドレスの例示）だけを、読み手の設定欄として差し替えてあります。

```javascript
// V35 本番コード：申請フォームの回答を集計し、集計表とグラフを作って担当者へメールする
const BOOK_NAME = '申請データ'; // 集計したい回答スプレッドシートの名前
const ANSWER_SHEET = '回答';
const RECIPIENT = 'tantosha@example.com'; // 集計メールを送る担当者のメールアドレス
const CATEGORIES = ['備品購入', '交通費', '会議費', '研修費', '印刷費'];

function parseAmount(cell) {
  if (typeof cell === 'number') return Math.round(cell);
  let s = String(cell);
  s = s.replace(/[０-９]/g, function (c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); });
  s = s.replace(/[,，\s　]/g, '').replace(/円/g, '');
  if (/^\d+$/.test(s)) return parseInt(s, 10);
  return null;
}

function yen(n) {
  return String(n).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function aggregateApplications() {
  const files = DriveApp.getFilesByName(BOOK_NAME);
  if (!files.hasNext()) { Logger.log('スプレッドシートが見つかりません'); return; }
  const ss = SpreadsheetApp.open(files.next());
  const sheet = ss.getSheetByName(ANSWER_SHEET);
  const values = sheet.getDataRange().getValues();

  const agg = {};
  CATEGORIES.forEach(function (c) { agg[c] = { count: 0, sum: 0, unknown: 0 }; });
  let pending = 0;
  for (let i = 1; i < values.length; i++) {
    const cat = String(values[i][2]);
    const amount = parseAmount(values[i][3]);
    const status = String(values[i][5]);
    if (!agg[cat]) agg[cat] = { count: 0, sum: 0, unknown: 0 };
    agg[cat].count += 1;
    if (amount === null) agg[cat].unknown += 1; else agg[cat].sum += amount;
    if (status === '未対応') pending += 1;
  }

  const old = ss.getSheetByName('集計');
  if (old) ss.deleteSheet(old);
  const out = ss.insertSheet('集計');
  out.appendRow(['種別', '金額合計', '件数', '金額不明件数']);
  let totalCount = 0, totalSum = 0, totalUnknown = 0;
  CATEGORIES.forEach(function (c) {
    const a = agg[c];
    out.appendRow([c, a.sum, a.count, a.unknown]);
    totalCount += a.count; totalSum += a.sum; totalUnknown += a.unknown;
    Logger.log(c + '：' + a.count + '件・' + yen(a.sum) + '円・金額不明' + a.unknown + '件');
  });
  out.appendRow(['合計', totalSum, totalCount, totalUnknown]);
  Logger.log('合計：' + totalCount + '件・' + yen(totalSum) + '円・金額不明' + totalUnknown + '件・未対応' + pending + '件');

  const chart = out.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(out.getRange('A1:B6'))
    .setOption('title', '種別別の金額合計')
    .setPosition(8, 1, 0, 0)
    .build();
  out.insertChart(chart);
  Logger.log('グラフを作成しました');

  const lines = CATEGORIES.map(function (c) {
    return c + '：' + agg[c].count + '件・' + yen(agg[c].sum) + '円';
  });
  const subject = '申請集計レポート（11月分）全' + totalCount + '件・合計' + yen(totalSum) + '円・未対応' + pending + '件';
  const body = '申請フォームの集計が完了しました。\n\n'
    + '全件数：' + totalCount + '件\n'
    + '合計金額：' + yen(totalSum) + '円（金額が読み取れた' + (totalCount - totalUnknown) + '件分）\n'
    + '金額不明：' + totalUnknown + '件\n'
    + '未対応：' + pending + '件\n\n'
    + '種別ごとの内訳\n' + lines.join('\n') + '\n\n'
    + '集計表とグラフは「' + BOOK_NAME + '」の集計シートに作成されています。';
  MailApp.sendEmail(RECIPIENT, subject, body);
  Logger.log('メールを送信しました：' + subject);
}
```

## 検証に使ったコード（読み取り専用）

集計シートが正解表と一致しているかを確かめるための、読み取り専用のスクリプトです。メール送信や書き込みはしません。正解表の数字は、自分のデータに合わせて書き換えて使ってください。

```javascript
// V35 検証コード（読み取り専用）：集計シートを正解表と照合する。メール送信や書き込みはしない。
const EXPECTED = {
  '備品購入': [16, 755900, 1],
  '交通費': [14, 97200, 1],
  '会議費': [10, 165600, 0],
  '研修費': [9, 316000, 0],
  '印刷費': [11, 169100, 1]
};
const EXP_TOTAL = [60, 1503800, 3];
const EXP_CATS = ['備品購入', '交通費', '会議費', '研修費', '印刷費'];

function verifyAggregation() {
  const files = DriveApp.getFilesByName('申請データ');
  const ss = SpreadsheetApp.open(files.next());
  const out = ss.getSheetByName('集計');
  const values = out.getDataRange().getValues();
  let ok = 0, ng = 0;
  function check(label, actual, expected) {
    if (actual === expected) { ok++; Logger.log('一致：' + label + ' = ' + actual); }
    else { ng++; Logger.log('不一致：' + label + ' 実際=' + actual + ' 正解=' + expected); }
  }
  EXP_CATS.forEach(function (c, i) {
    const row = values[i + 1];
    check(c + 'の種別名', String(row[0]), c);
    check(c + 'の金額合計', row[1], EXPECTED[c][1]);
    check(c + 'の件数', row[2], EXPECTED[c][0]);
    check(c + 'の金額不明件数', row[3], EXPECTED[c][2]);
  });
  const t = values[6];
  check('合計の金額', t[1], EXP_TOTAL[1]);
  check('合計の件数', t[2], EXP_TOTAL[0]);
  check('合計の金額不明件数', t[3], EXP_TOTAL[2]);
  const charts = out.getCharts().length;
  if (charts >= 1) { ok++; Logger.log('一致：グラフの個数 = ' + charts); }
  else { ng++; Logger.log('不一致：グラフがありません'); }
  Logger.log('検証結果：一致' + ok + '件・不一致' + ng + '件');
}
```

## 正直な制限

- 金額が読み取れない行（この検証では「未定」2件・空欄1件）は、金額合計には入りません。件数には数えられ、「金額不明」として集計表とメールに別掲されます
- 金額の表記ゆれ（カンマ入り・「円」付き・全角数字）は正規化して集計できますが、「約3万円」のような書き方までは拾えません
- 種別はフォームの選択式（ラジオボタン等）が前提です。自由記述だと表記ゆれで集計が割れます

## 関連

- 動画版（実行画面）：https://youtu.be/G8ovAzHkOz8
- チェックリスト（note）：https://note.com/gas_recipe/n/nd26e2687d421
