// 回歸測試：驗證 index.html 的解析邏輯沒有壞掉。
//   跑法： npm test   （等同 node test.js）
//
// 原理：直接從 index.html 抽出 //__PARSER_CORE_START__ 到 //__PARSER_CORE_END__
// 之間「同一份位元組」來執行——所以這裡測的就是瀏覽器真正跑的邏輯，
// 不存在「另一份副本悄悄走樣」的問題。改解析請只改 index.html 的 PARSER CORE 區塊。
//
// fixture 是 test/fixture.xlsx（匿名資料，見 test/make-fixture.js）。
// 下面 golden 值是刻意設計的預期結果；若你「有意」改動解析行為，
// 請一併更新這些 golden 值，並在 commit 訊息說明為什麼。
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

// ---- 從 index.html 抽出 PARSER CORE 並實體化 ----
const html = fs.readFileSync(path.join(__dirname, 'index.html'), 'utf8');
const START = '//__PARSER_CORE_START__', END = '//__PARSER_CORE_END__';
if (!html.includes(START) || !html.includes(END)) {
  console.error('✗ 在 index.html 找不到 PARSER CORE 標記；有人刪掉或改名了標記。');
  process.exit(2);
}
const block = html.split(START)[1].split(END)[0];
let P;
try {
  const factory = new Function('XLSX',
    block + '\n;return {parseTemplate,parseEstimate,fmtSec,FN,N,num,trunc2,effDelta,leafIssues};');
  P = factory(XLSX);
} catch (e) {
  console.error('✗ PARSER CORE 抽出後無法執行（可能語法錯誤或引用了 DOM）：', e.message);
  process.exit(2);
}

// ---- 讀 fixture ----
const fixture = path.join(__dirname, 'test', 'fixture.xlsx');
if (!fs.existsSync(fixture)) {
  console.error('✗ 找不到 test/fixture.xlsx，請先跑： node test/make-fixture.js');
  process.exit(2);
}
const wb = XLSX.readFile(fixture);

// ---- 斷言工具 ----
let fails = 0;
function eq(label, got, want) {
  const ok = got === want;
  if (!ok) fails++;
  console.log((ok ? '✓' : '✗') + ' ' + label + '：得 ' + JSON.stringify(got) +
    (ok ? '' : '，應為 ' + JSON.stringify(want)));
}
const sum = arr => arr.reduce((s, it) => s + it.amt, 0);
const warns = arr => arr.filter(it => it.warn).length;

// ---- 範本格式 ----
console.log('\n[範本格式 parseTemplate]');
const t = P.parseTemplate(wb);
eq('追加筆數', t.add.length, 3);
eq('追加總額', sum(t.add), 36000);
eq('追加警告數', warns(t.add), 1);
eq('B002 原數量空白', t.add[1].origQty, '');
eq('追減筆數', t.reduce.length, 2);
eq('追減總額', sum(t.reduce), -11000);
eq('新增筆數', t.newItems.length, 1);
eq('新增總額', sum(t.newItems), 30000);
eq('新增不議筆數', t.noNeg.length, 1);
eq('新增不議總額', sum(t.noNeg), 50000);

// ---- 概算書格式（明細表 v2 版型，計算整套對齊 skill）----
console.log('\n[概算書格式 parseEstimate（skill 移植版）]');
const e = P.parseEstimate(wb, true);   // fixture 是 .xlsx → 手動填值檢查開啟
const S = e.S;
eq('版型偵測', e.budget.layout, 'v2');
eq('(一)追加 leaf 數（大項0111不算）', S.add.leaves.filter(l => l.cat === '011').length, 2);
eq('(一)表列小計', S.add_011, 39000);
eq('(一)逐項實算', S.add_011_calc, 39000);
eq('(二)表列小計', S.ded_011, -15000);
eq('(三)新增合計（議價50000+不議價20000，退回逐項加總）', S.new_011, 70000);
eq('(四)變更項目總合', S.sec4, 94000);
eq('(五)包工費以外 012+013+014', S.sec5, 5850);
eq('(六)變更總金額(未稅)', S.sec6, 99850);
eq('(七)0B營業稅', S.sec7, 4993);
eq('(八)最終變更金額(含稅)', S.sec8, 104843);

const g1 = e.report.groups[0].lines;
eq('追加句（v5 句型：單位＋元）', g1[0],
  '1.項目0111A，工項名稱：排水溝，原契約數量100M，單價為1,500元，追加數量20M，追加金額30,000元。');
eq('紅字警告（表填49,000 vs 實算48,000，附列號）',
  g1[1].includes('【警告(Excel第9列)：變更金額 表填 49,000，實算 48,000'), true);
eq('追加小計行（表列優先）', g1[2], '小計：追加39,000元。');
const g2 = e.report.groups[1].lines;
eq('全數追減句型', g2[1],
  '2.項目0112B，工項名稱：舊柵欄，原契約數量30M，單價為200元，全數追減，追減金額6,000元。');
eq('追減小計行（abs 呈現）', g2[2], '小計：追減15,000元。');
const g3 = e.report.groups[2].lines;
eq('新增分「需議價」標頭', g3[0], '◎需議價部分（依政府採購法規辦理）：');
eq('新增分「不議價」標頭', g3[2], '◎不議價部分：');
eq('不議價項目接續編號', g3[3].startsWith('2.項目N401'), true);
const gw = e.report.groups.find(g => g.title.startsWith('檢查提醒'));
eq('手動填值提醒恰 1 處', gw.lines.filter(l => l.includes('手打常數')).length, 1);
eq('手動填值指向第9列變更金額', gw.lines[0].includes('第9列') && gw.lines[0].includes('變更金額'), true);

// ---- 計算原則單點測試（DECISIONS #14/#25：取 2 位無條件捨去）----
console.log('\n[計算原則 trunc2 / effDelta]');
eq('無條件捨去', P.trunc2(123.456), 123.45);
eq('浮點雜訊先吸收再截斷', P.trunc2(3 * 2575.4), 7726.2);
eq('effDelta 增減欄優先', P.effDelta({ da: -5, na: 100, oa: 90 }, 'add'), -5);
eq('effDelta 新增合約側視為0', P.effDelta({ da: null, na: 500, oa: null }, 'new'), 500);

// ---- 句型 fmtSec（對齊 v5 範本：單價/新增單價帶「元」、數量帶單位、小計行） ----
console.log('\n[句型 fmtSec]');
const fa = P.fmtSec('add', [{code:'A101', name:'排水溝', origQty:'10', unit:'M', price:1000, qty:2, amt:2000}]);
eq('追加句', fa.lines[0].t,
  '1、項目：A101，工項名稱：排水溝，原契約數量10M，單價為1,000元，追加數量2M，追加金額2,000元。');
eq('追加小計行', fa.lines[1].t, '小計：追加2,000元。');
const fr = P.fmtSec('reduce', [{code:'B201', name:'圍籬', origQty:'5', unit:'M', price:100, qty:-2, amt:-200}]);
eq('追減句（abs 呈現）', fr.lines[0].t,
  '1、項目：B201，工項名稱：圍籬，原契約數量5M，單價為100元，追減數量2M，追減金額200元。');
eq('追減小計行（abs 呈現）', fr.lines[1].t, '小計：追減200元。');
const fx = P.fmtSec('new', [{code:'N301', name:'新工項', origQty:'', unit:'座', price:500, qty:3, amt:1500}]);
eq('新增句', fx.lines[0].t,
  '1、項目：N301，工項名稱：新工項，新增數量3座，新增單價500元，追加金額1,500元。');

// ---- 數字格式 FN ----
console.log('\n[數字格式 FN]');
eq('千分位', P.FN(1234567), '1,234,567');
eq('負數', P.FN(-9000), '-9,000');
eq('小數去尾0', P.FN(18.4), '18.4');

console.log('\n' + (fails === 0
  ? '結論：全部通過（0 異常）。解析邏輯與 golden 一致。'
  : '結論：' + fails + ' 處異常——解析行為與 golden 不符，請檢查是否為預期改動。'));
process.exit(fails === 0 ? 0 : 1);
