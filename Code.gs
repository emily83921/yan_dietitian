/****************************************************************
 * 營養師彥彥｜代謝卡關檢測 後端（Google Apps Script）
 * ---------------------------------------------------------------
 * 第一次使用：
 *   1. 把下面的 ADMIN_KEY 改成你自己的後台密碼
 *   2. 上方函式選單選「setup」→ 按「執行」→ 授權
 *      會自動建立 4 個分頁：填答資料、題目、型態、網站文案
 *   3. 「部署」→「新增部署作業」→ 類型選「網頁應用程式」
 *      執行身分：我　／　存取權：所有人
 *   4. 第一次寄信前，上方函式選單選「testMail」→ 執行 → 授權寄信
 *      （信會從「部署這份程式的 Gmail」寄出，免費帳號一天約 100 封）
 * 之後在「題目」「型態」「網站文案」分頁改內容，網站和信件就會自動更新。
 ****************************************************************/

const ADMIN_KEY = '請改成你的後台密碼';   // ← 改這裡，至少 8 個字，不要用生日

const SHEET_DATA = '填答資料', SHEET_Q = '題目', SHEET_TYPE = '型態', SHEET_TEXT = '網站文案';
const Q_HEADERS = ['啟用', '維度', '題目', '情境小語', '選項A', 'A加分', '選項B', 'B加分', '選項C', 'C加分', '選項D', 'D加分'];
const T_HEADERS = ['代號', '名稱', '英文名稱', '一句話', '卡關原因', '常踩地雷', '飲食調整', '生活調整', '核心句', '菜單提醒', '早餐原因'];
const BASE_HEADERS = ['編號', '填寫時間', '狀態', '報告寄送', '備註'];
const SHEET_ID = '1LWcy7e5Oe6kcSZPDLhZhbvUQyp8_ghPetr155U3HY4g';   // 你的 Google 試算表 ID（網址 /d/ 後面那串）
function SS() { return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet(); }
const SITE_URL = 'https://emily83921.github.io/yan_dietitian/';   // 信件裡「回到我的完整報告」只允許連回這個網址

const DEFAULT_TEXT = {
  "品牌名稱": "營養師彥彥",
  "品牌英文": "Metabolic Block Assessment",
  "首頁徽章": "60 秒 · 6 型代謝卡關檢測",
  "首頁前言": "少吃、多動，體重還是不動？",
  "首頁標題": "你的身體\n**卡在哪一關**？",
  "首頁英文": "Find your block. Then fix it.",
  "首頁說明": "同樣少吃多動，有人瘦、有人卡關，差別在於**你了不瞭解自己**。60 秒找出你的代謝卡關類型，拿到專屬於你的分析報告——每日熱量建議、早餐怎麼吃，完整報告直接寄到你的信箱。",
  "開始按鈕": "開始檢測，找出我的卡關點",
  "資歷": "臺北醫學大學｜肥胖研究學會 體重管理營養師｜CTSSN 運動營養專業認證｜體適能健身 C 級指導員",
  "IG網址": "https://www.instagram.com/yan_dietitian/",
  "IG帳號": "@yan_dietitian",
  "LINE網址": "",
  "健康狀況選項": "血壓偏高,血糖偏高或胰島素阻抗,多囊性卵巢（PCOS）,甲狀腺功能低下,血脂偏高,便秘或腸胃不適,經期不規則,長期失眠,懷孕或哺乳中,腎臟或其他慢性病",
  "解鎖標題": "你的完整代謝報告已經生成",
  "解鎖說明": "輸入 email，我把**完整報告**寄到你的信箱。\n這是第一步「看到問題」，下一頁開始第二步：你這型該怎麼吃。",
  "同意條款": "按下「解鎖」，代表你同意營養師彥彥蒐集你的作答（包含健康狀況），只用於提供營養建議與寄送報告，不會提供給第三方，你可以隨時來信要求刪除。",
  "結果CTA標題": "這一次，你真的下定決心要瘦了嗎？",
  "結果CTA說明": "知道自己是哪一型，只是第一步。如果你已經準備好，這次想用對方法、好好瘦一次，私訊我「代謝」兩個字，我會依照你的型態和生活習慣，陪你找出最適合你的瘦法。",
  "分布顯示門檻": "30",
  "免責聲明": "本檢測依你的作答整理出「傾向」，僅供生活參考，不是醫療診斷。懷孕、哺乳、正在用藥或有慢性病者，請先諮詢醫師或營養師。",
  "信件寄件人": "營養師彥彥",
  "信件主旨": "你的代謝報告好了——你是「{型態}」",
  "信件開頭": "Hello！我是營養師彥彥。\n你剛做完的代謝卡關檢測，我看過你的結果了。先說一件我想讓你放心的事：你會一直卡住，很多時候真的不是你不夠自律。",
  "信件結尾": "知道自己是哪一型，只是第一步。如果你已經下定決心，這次想用對方法好好瘦一次，到 IG 私訊我「代謝」兩個字，我會依照你的型態和生活習慣，陪你找出最適合你的瘦法。",
  "信件按鈕": "我決定了，私訊彥彥"
};
const DEFAULT_TYPES = [
 {
  "code": "P",
  "name": "高壓緊繃型",
  "en": "High-Tension",
  "tag": "身體一直處在備戰狀態，脂肪也跟著守在腹部。",
  "why": [
   "長期壓力讓壓力荷爾蒙（皮質醇）偏高，身體傾向把脂肪囤在腹部",
   "壓力大時，食慾和選擇容易偏向高油、高糖",
   "緊繃會影響睡眠，睡不好又讓壓力更難消化，形成循環"
  ],
  "mines": [
   "靠咖啡硬撐",
   "忙到忘記吃，晚上補償性大吃",
   "睡不夠還做高強度運動",
   "把甜食當紓壓"
  ],
  "food": [
   "三餐定時，不讓血糖大起大落",
   "每餐都有蛋白質和深綠色蔬菜",
   "咖啡一天 1–2 杯，下午 2 點後不喝"
  ],
  "life": [
   "每天留 10 分鐘放鬆：散步、伸展或深呼吸",
   "運動以中低強度和肌力訓練為主",
   "睡前 1 小時離開工作訊息"
  ],
  "core": "先讓身體鬆下來，脂肪才願意離開。",
  "bkWhy": "早餐有蛋白質和高纖澱粉，血糖穩、飽足感久，身體不用靠壓力荷爾蒙硬撐，下午比較不會累到想亂吃。",
  "menuTip": "咖啡放在早餐後再喝，不要空腹靠咖啡撐到中午。"
 },
 {
  "code": "E",
  "name": "情緒飢餓型",
  "en": "Emotional Hunger",
  "tag": "你餓的不是胃，而是心情。",
  "why": [
   "疲累或低落時，大腦會想用食物快速得到安慰",
   "邊滑手機邊吃，飽足訊號被忽略，不知不覺吃多",
   "吃完自責，壓力更大，又想再吃"
  ],
  "mines": [
   "追劇配零食",
   "心情差就叫外送",
   "用甜點獎勵自己",
   "整包吃完才回神"
  ],
  "food": [
   "想吃時先喝一杯水，等 3 分鐘再決定",
   "零食換成無糖優格、毛豆或堅果，先分裝好",
   "每吃一口放下筷子，一餐吃 20 分鐘"
  ],
  "life": [
   "列 5 件不靠食物的紓壓方法",
   "吃東西時不看手機",
   "記下想吃東西的時間和心情，找出觸發點"
  ],
  "core": "照顧好情緒，食物就不用負責安慰你。",
  "bkWhy": "蛋白質＋原型澱粉的組合消化慢、飽足感久，上午情緒比較穩，不容易想用零食找安慰。",
  "menuTip": "坐下來吃，放下手機，吃 15 分鐘以上。"
 },
 {
  "code": "N",
  "name": "作息失衡型",
  "en": "Off-Rhythm",
  "tag": "身體的時鐘，跟你的生活對不上。",
  "why": [
   "睡不夠時飢餓感上升、飽足感下降，隔天特別想吃",
   "晚睡讓進食時間一路往後延，宵夜變成常態",
   "白天沒精神，只好靠糖和咖啡撐住"
  ],
  "mines": [
   "熬夜追劇",
   "深夜泡麵、鹹酥雞",
   "睡前一杯手搖飲",
   "睡過頭跳過早餐"
  ],
  "food": [
   "晚餐吃夠蛋白質和蔬菜，半夜比較不會餓",
   "睡前 3 小時不再進食",
   "真的餓了，選無糖豆漿或一顆蛋"
  ],
  "life": [
   "固定起床時間，假日差距不超過 1 小時",
   "12 點前上床，先從提早 30 分鐘開始",
   "早上曬 10 分鐘太陽，幫身體校正時鐘"
  ],
  "core": "把時鐘調回來，代謝就會跟上。",
  "bkWhy": "固定時間吃一份有蛋白質的早餐，是幫身體校正時鐘的訊號，也能減少晚上的飢餓感和宵夜。",
  "menuTip": "就算前一晚晚睡，也在差不多的時間吃。"
 },
 {
  "code": "M",
  "name": "代謝低速型",
  "en": "Slow Burner",
  "tag": "不是吃太多，是身體燃燒得太少。",
  "why": [
   "久坐少動，每天實際消耗的熱量比想像中少",
   "肌肉量不足，基礎代謝跟著往下掉",
   "循環變慢，容易怕冷、下半身水腫"
  ],
  "mines": [
   "吃完就坐著或躺著",
   "通勤全程都坐",
   "只節食不運動",
   "週末宅一整天"
  ],
  "food": [
   "每餐一掌心蛋白質，幫助保住肌肉",
   "運動後 1 小時內補一份蛋白質",
   "不要用少吃取代運動"
  ],
  "life": [
   "每坐 1 小時，起來動 3 分鐘",
   "每天多走 2,000 步",
   "每週 2–3 次肌力訓練：深蹲、臀橋都可以"
  ],
  "core": "肌肉，是你最好的燃脂引擎。",
  "bkWhy": "早餐的蛋白質幫你保住肌肉，肌肉越多，你平常沒在動的時候燃燒的熱量也越多。",
  "menuTip": "吃完可以提早一站下車，走 10 分鐘再進辦公室。"
 },
 {
  "code": "G",
  "name": "血糖起伏型",
  "en": "Sugar Swing",
  "tag": "能量忽高忽低，跟著血糖坐雲霄飛車。",
  "why": [
   "精緻澱粉和含糖飲料讓血糖快速上升，又快速掉下來",
   "血糖一掉，就頭暈、想睡、特別想吃甜",
   "胰島素長期偏高時，脂肪比較容易囤積"
  ],
  "mines": [
   "全糖手搖飲",
   "麵包當早餐",
   "飯後一定要甜點",
   "下午三點的餅乾"
  ],
  "food": [
   "吃飯順序：先菜、再肉、最後飯",
   "精緻澱粉換成全穀、地瓜、南瓜",
   "飲料每週降一級甜度，最後到無糖"
  ],
  "life": [
   "飯後散步 10 分鐘",
   "下午固定安排一份點心，避免餓過頭",
   "睡飽 7 小時，隔天比較不會想吃甜"
  ],
  "core": "穩住血糖，嘴饞就會安靜下來。",
  "bkWhy": "原型澱粉消化慢，搭配蛋白質和蔬菜一起吃，血糖上升比較平緩，不會一早就坐雲霄飛車，下午也比較不會想吃甜。",
  "menuTip": "吃的順序：蔬菜和蛋白質先吃，澱粉留到最後。"
 },
 {
  "code": "D",
  "name": "節食耗損型",
  "en": "Under-Fueled",
  "tag": "吃得很少，身體卻開始省電。",
  "why": [
   "長期吃太少，身體會降低消耗來保護自己",
   "蛋白質不夠，肌肉流失、代謝跟著下降",
   "一放鬆就復胖，越減越難瘦"
  ],
  "mines": [
   "一杯咖啡當早餐",
   "只吃沙拉",
   "極端斷食",
   "天天量體重"
  ],
  "food": [
   "每餐都吃到一掌心蛋白質",
   "熱量不要低於基礎代謝",
   "三餐定時，不跳餐"
  ],
  "life": [
   "一週量一次體重就好",
   "搭配肌力訓練，把肌肉練回來",
   "把目標放在體脂和腰圍，不只看體重"
  ],
  "core": "先吃夠，身體才願意燃燒。",
  "bkWhy": "份量足、醣類和蛋白質都有，讓身體知道食物是夠的，它才願意把代謝開回來。",
  "menuTip": "請照份量吃完，不要自己再減。"
 },
 {
  "code": "B",
  "name": "穩定平衡型",
  "en": "In Balance",
  "tag": "你的代謝節奏很穩，接下來只差細節。",
  "why": [
   "飲食、作息、壓力都沒有明顯的卡關點",
   "接下來的重點是份量控制和持續",
   "照著建議熱量吃，維持或調整體態都會很順"
  ],
  "mines": [
   "偶爾的聚餐大餐",
   "忙起來就忘了喝水"
  ],
  "food": [
   "照建議熱量吃",
   "每週嘗試一種新的蔬菜或蛋白質"
  ],
  "life": [
   "維持現在的運動習慣",
   "每週量一次體重和腰圍"
  ],
  "core": "穩定，就是最強的優勢。",
  "bkWhy": "份量剛好、營養均衡，延續你目前穩定的節奏就好。",
  "menuTip": "好好享受食物，偶爾換換不同的蛋白質來源。"
 }
];
const DEFAULT_QS = [{"dim": "Body Signals · 身體訊號", "t": "脂肪最容易堆在哪裡？", "s": "身體囤脂的位置，會透露一些線索。", "o": [["沒有特別明顯的地方", ""], ["肚子和腰圍，四肢相對細", "P2 G2"], ["全身軟軟的，沒什麼肌肉線條", "M3 D1"], ["早晚體重差很多，容易水腫", "N2 M1"]]}, {"dim": "Body Signals · 身體訊號", "t": "下午三、四點的你，通常是？", "s": "下午的精神狀態，和血糖、睡眠都有關係。", "o": [["精神還不錯", ""], ["很睏，特別想吃甜的", "G3"], ["靠咖啡或提神飲料撐著", "N2 P1"], ["頭暈手抖，一定要馬上吃東西", "G2 D1"]]}, {"dim": "Eating Patterns · 飲食行為", "t": "肚子餓的時候，你的反應是？", "s": "飢餓的樣子，每個人都不一樣。", "o": [["慢慢有感覺，可以等到正餐", ""], ["心慌手抖、脾氣變差，必須馬上吃", "G3"], ["常常不太餓，會忘記吃", "D3"], ["肚子不餓，但嘴巴就是想吃", "E3"]]}, {"dim": "Eating Patterns · 飲食行為", "t": "壓力很大的那幾天，你的飲食會？", "s": "壓力，會悄悄改變你想吃的東西。", "o": [["跟平常差不多", ""], ["吃不下，胃緊緊的", "P3"], ["特別想吃炸物、甜點", "E2 P1"], ["邊忙邊吃，回神整包都吃完了", "E2 G1"]]}, {"dim": "Eating Patterns · 飲食行為", "t": "你對甜食和澱粉的渴望？", "s": "嘴饞，常常不只是因為肚子餓。", "o": [["偶爾想吃，吃一點就滿足", ""], ["正餐吃飽了，還是要來點甜的", "G3"], ["心情不好時特別想吃", "E3"], ["越晚越想吃重口味或泡麵", "N2 E1"]]}, {"dim": "Eating Patterns · 飲食行為", "t": "每一餐都有一掌心的豆魚蛋肉嗎？", "s": "蛋白質不夠，肌肉會悄悄流失。", "o": [["每餐都有", ""], ["大部分有", "D1"], ["偶爾才有", "D2 M1"], ["很少，常常只吃澱粉", "D2 G1"]]}, {"dim": "Energy & Metabolism · 能量與代謝", "t": "誠實說，你過去的減重經驗？", "s": "過去用過的方法，會在身體留下痕跡。", "o": [["沒有刻意減過，或減得很順", ""], ["吃得很少，體重卻不動", "D2 M2"], ["壓力一大就變胖", "P3"], ["減了又復胖，來回好幾次", "D2 E1"]]}, {"dim": "Energy & Metabolism · 能量與代謝", "t": "你平常的活動狀態？", "s": "身體的燃燒，大多發生在你沒注意的時候。", "o": [["每週運動 3 次以上", ""], ["偶爾運動，但一天坐很久", "M2"], ["幾乎不運動，通勤也都坐著", "M3"], ["只靠少吃來控制，不太運動", "D1 M2"]]}, {"dim": "Sleep & Stress · 作息與壓力", "t": "最近的睡眠狀況？", "s": "睡得好不好，隔天的食慾都知道。", "o": [["11 點前睡，一覺到天亮", ""], ["身體很累，但腦袋停不下來", "P3"], ["常常超過 12 點睡，睡不夠", "N3"], ["半夜容易醒，起床還是很累", "N2 P1"]]}, {"dim": "Sleep & Stress · 作息與壓力", "t": "你目前最大的壓力來源？", "s": "壓力從哪裡來，決定了它怎麼影響你。", "o": [["還好，沒有特別的壓力", ""], ["工作或課業", "P3"], ["感情、家人或人際關係", "E3"], ["對自己的身材不滿意，一直在減肥", "D2 E1"]]}];

/* ---------- 第一次設定 ---------- */
function setup() {
  const ss = SS();

  if (!ss.getSheetByName(SHEET_TEXT)) {
    const t = ss.insertSheet(SHEET_TEXT);
    const rows = [['項目', '內容', '說明']].concat(Object.keys(DEFAULT_TEXT).map(k => [k, DEFAULT_TEXT[k], hintFor(k)]));
    t.getRange(1, 1, rows.length, 3).setValues(rows);
    styleHeader(t, 3); t.setColumnWidth(1, 130); t.setColumnWidth(2, 520); t.setColumnWidth(3, 320);
    t.getRange(2, 2, rows.length, 2).setWrap(true);
  }

  else {   // 已經有分頁：只補上新的項目，不動你改過的內容
    const t = ss.getSheetByName(SHEET_TEXT);
    const have = t.getLastRow() > 1 ? t.getRange(2, 1, t.getLastRow() - 1, 1).getValues().map(r => String(r[0]).trim()) : [];
    const add = Object.keys(DEFAULT_TEXT).filter(k => have.indexOf(k) === -1).map(k => [k, DEFAULT_TEXT[k], hintFor(k)]);
    if (add.length) t.getRange(t.getLastRow() + 1, 1, add.length, 3).setValues(add).setWrap(true);
  }

  if (!ss.getSheetByName(SHEET_TYPE)) {
    const t = ss.insertSheet(SHEET_TYPE);
    const rows = [T_HEADERS].concat(DEFAULT_TYPES.map(x => [x.code, x.name, x.en, x.tag, x.why.join('｜'), x.mines.join('｜'), x.food.join('｜'), x.life.join('｜'), x.core, x.menuTip, x.bkWhy]));
    t.getRange(1, 1, rows.length, T_HEADERS.length).setValues(rows);
    styleHeader(t, T_HEADERS.length); t.setFrozenColumns(2);
    for (let c = 4; c <= 11; c++) t.setColumnWidth(c, 280);
    t.getRange(2, 1, rows.length, T_HEADERS.length).setWrap(true).setVerticalAlignment('top');
    t.getRange(1, 1).setNote('代號不要改：P 高壓、E 情緒、N 作息、M 代謝、G 血糖、D 節食、B 平衡（分數都不明顯時）。網站上 No.01–06 依這裡的排列順序編號。多個項目用「｜」隔開。');
  }

  if (!ss.getSheetByName(SHEET_Q)) {
    const q = ss.insertSheet(SHEET_Q);
    const rows = [Q_HEADERS].concat(DEFAULT_QS.map(x => {
      const r = ['是', x.dim, x.t, x.s];
      for (let i = 0; i < 4; i++) { r.push(x.o[i] ? x.o[i][0] : ''); r.push(x.o[i] ? x.o[i][1] : ''); }
      return r;
    }));
    q.getRange(1, 1, rows.length, Q_HEADERS.length).setValues(rows);
    styleHeader(q, Q_HEADERS.length); q.setFrozenColumns(3);
    q.setColumnWidth(2, 200); q.setColumnWidth(3, 300); q.setColumnWidth(4, 260);
    q.getRange(2, 1, 300, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['是', '否']).build());
    q.getRange(1, 6).setNote('加分寫法：型態代號＋分數，例如「G2 D1」= 血糖起伏型加 2 分、節食耗損型加 1 分。留空代表這個選項不加分（最健康的選項）。');
  }

  if (!ss.getSheetByName(SHEET_DATA)) {
    const d = ss.insertSheet(SHEET_DATA, 0);
    d.getRange(1, 1, 1, BASE_HEADERS.length).setValues([BASE_HEADERS]);
    styleHeader(d, BASE_HEADERS.length);
  }
  const blank = ss.getSheetByName('工作表1') || ss.getSheetByName('Sheet1');
  if (blank && ss.getSheets().length > 1 && blank.getLastRow() === 0) ss.deleteSheet(blank);
}

function styleHeader(sh, n) {
  sh.getRange(1, 1, 1, n).setFontWeight('bold').setBackground('#0D0B0C').setFontColor('#F2A7BD');
  sh.setFrozenRows(1);
}
function hintFor(k) {
  return ({
    '首頁標題': '用 **兩個星號** 包起來的字會變粉色；換行按 Alt+Enter（Mac：⌘+Enter）',
    '資歷': '用「｜」隔開每一項',
    '健康狀況選項': '用逗號隔開',
    'LINE網址': '還沒有就留空，網站會顯示「即將開放」',
    '信件主旨': '{型態} 會自動換成填答者的型態名稱',
    '信件開頭': '換行按 Alt+Enter（Mac：⌘+Enter）',
    '信件寄件人': '收件人看到的寄件者名稱',
    '分布顯示門檻': '填答人數到這個數字，首頁才會顯示「型態分布」和完成人數',
  })[k] || '';
}

/* ---------- 網站讀取內容 ---------- */
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  if (action === 'config') return json(getConfig());
  return json({ ok: true, message: '連線成功' });
}

function getConfig() {
  const ss = SS();
  const read = (name, n) => { const s = ss.getSheetByName(name); return s && s.getLastRow() > 1 ? s.getRange(2, 1, s.getLastRow() - 1, n).getDisplayValues() : []; };

  const text = {};
  read(SHEET_TEXT, 2).forEach(r => { if (r[0]) text[r[0].trim()] = r[1]; });

  const types = read(SHEET_TYPE, T_HEADERS.length).filter(r => r[0] && r[1]).map(r => ({
    code: r[0].trim().toUpperCase(), name: r[1], en: r[2], tag: r[3],
    why: r[4], mines: r[5], food: r[6], life: r[7], core: r[8], menuTip: r[9], bkWhy: r[10] || ''
  }));

  const questions = [];
  read(SHEET_Q, Q_HEADERS.length).forEach(r => {
    if (r[0] === '否' || !r[2]) return;
    const o = [];
    for (let i = 0; i < 4; i++) { const txt = r[4 + i * 2].trim(); if (txt) o.push([txt, r[5 + i * 2]]); }
    if (o.length >= 2) questions.push({ dim: r[1], t: r[2], s: r[3], o: o });
  });
  // 型態分布：只統計真實填答（網站人數達「分布顯示門檻」才會顯示）
  const stats = { total: 0, types: {} };
  const ds = ss.getSheetByName(SHEET_DATA);
  if (ds && ds.getLastRow() > 1) {
    const col = ds.getRange(1, 1, 1, ds.getLastColumn()).getValues()[0].indexOf('主要型態') + 1;
    if (col) ds.getRange(2, col, ds.getLastRow() - 1, 1).getValues().forEach(r => { const k = String(r[0]); if (k) { stats.total++; stats.types[k] = (stats.types[k] || 0) + 1; } });
  }
  return { ok: true, text: text, types: types, questions: questions, stats: stats };
}

/* ---------- 收資料 / 後台 ---------- */
function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents || '{}'); } catch (err) { return json({ ok: false, error: '資料格式錯誤' }); }
  if (body.action === 'submit') return json(submit(body.data || {}, body.mail));

  if (!body.key || body.key !== ADMIN_KEY || ADMIN_KEY === '請改成你的後台密碼') {
    Utilities.sleep(800);
    return json({ ok: false, error: '密碼錯誤' });
  }
  if (body.action === 'list') return json(listRows());
  if (body.action === 'update') return json(updateRow(body.id, body.status, body.note));
  return json({ ok: false, error: '未知的動作' });
}

function submit(data, mail) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SS();
    if (!ss.getSheetByName(SHEET_DATA)) setup();
    const sh = ss.getSheetByName(SHEET_DATA);
    const row = Object.assign({
      '編號': Utilities.getUuid().slice(0, 8),
      '填寫時間': Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm'),
      '狀態': '未聯繫', '報告寄送': '', '備註': ''
    }, data);
    if (mail) row['報告寄送'] = sendReport(String(data['Email'] || ''), mail);
    const headers = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].filter(String);
    Object.keys(row).forEach(k => { if (headers.indexOf(k) === -1) headers.push(k); });
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    styleHeader(sh, headers.length);
    sh.appendRow(headers.map(h => clean(row[h])));
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

// 避免有人在表單輸入試算表公式
function clean(v) {
  if (v === undefined || v === null) return '';
  if (typeof v === 'string' && /^[=+\-@]/.test(v)) return "'" + v;
  return v;
}

function listRows() {
  const sh = SS().getSheetByName(SHEET_DATA);
  if (!sh || sh.getLastRow() < 2) return { ok: true, rows: [] };
  const values = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getDisplayValues();
  const headers = values.shift();
  const rows = values.map(r => { const o = {}; headers.forEach((h, i) => { if (h) o[h] = r[i]; }); return o; });
  return { ok: true, rows: rows.reverse() };
}

function updateRow(id, status, note) {
  const sh = SS().getSheetByName(SHEET_DATA);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idCol = headers.indexOf('編號') + 1, stCol = headers.indexOf('狀態') + 1, noteCol = headers.indexOf('備註') + 1;
  const ids = sh.getRange(2, idCol, Math.max(sh.getLastRow() - 1, 1), 1).getValues().map(r => String(r[0]));
  const i = ids.indexOf(String(id));
  if (i === -1) return { ok: false, error: '找不到這筆資料' };
  if (status !== undefined) sh.getRange(i + 2, stCol).setValue(clean(String(status)));
  if (note !== undefined) sh.getRange(i + 2, noteCol).setValue(clean(String(note)));
  return { ok: true };
}


/* ---------- 寄送報告 ---------- */
// 信件內容由這裡用試算表的「型態」「網站文案」組出來；網站只傳型態代號和數字，避免被拿去亂寄信
function sendReport(to, m) {
  try {
    to = to.trim();
    if (!/^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/.test(to) || to.length > 120) return '未寄：email 格式錯誤';
    const cache = CacheService.getScriptCache(), ck = 'sent_' + to.toLowerCase();
    if (cache.get(ck)) return '未寄：10 分鐘內已寄過';
    if (MailApp.getRemainingDailyQuota() < 1) return '未寄：今天寄信額度用完';
    const cfg = getConfig();
    const text = Object.assign({}, DEFAULT_TEXT, cfg.text);
    const types = (cfg.types.length >= 2 ? cfg.types : DEFAULT_TYPES.map(x => Object.assign({}, x, {
      why: x.why.join('｜'), mines: x.mines.join('｜'), food: x.food.join('｜'), life: x.life.join('｜') })));
    const main = types.filter(t => t.code !== 'B');
    const T = c => types.filter(t => t.code === String(c || '').toUpperCase())[0];
    const NO = c => { const i = main.map(t => t.code).indexOf(c); return i < 0 ? '0' : String(i + 1); };
    const t = T(m.main) || types[0];
    const s = m.second ? T(m.second) : null;
    const pct = {}; main.forEach(x => { pct[x.code] = Math.max(0, Math.min(100, Math.round(+((m.pct || {})[x.code]) || 0))); });
    const n = m.numbers ? {
      kcal: Math.round(+m.numbers.kcal) || 0, bmi: (+m.numbers.bmi || 0).toFixed(1), sex: m.numbers.sex === '男' ? '男' : '女',
      bf: +m.numbers.bf || 0, waist: +m.numbers.waist || 0,
      ideal: [Math.round(+(m.numbers.ideal || [])[0]) || 0, Math.round(+(m.numbers.ideal || [])[1]) || 0],
      note: String(m.numbers.note || '').slice(0, 60)
    } : null;
    if (n && (n.kcal < 1000 || n.kcal > 5000)) return '未寄：數字異常';
    const link = String(m.link || '').indexOf(SITE_URL) === 0 && String(m.link).length < 2000 ? String(m.link) : SITE_URL;
    const bk = makeBreakfast(t.code, n ? n.kcal : 1500, ['葷食', '蛋奶素', '全素'].indexOf(m.diet) >= 0 ? m.diet : '葷食', m.meal === '外食' ? '外食' : '自煮');
    const data = { t: t, s: s, sNo: s ? NO(s.code) : '', no: NO(t.code), pct: pct, main: main, n: n, bk: bk, link: link, reportNo: String(m.reportNo || '').replace(/\D/g, '').slice(0, 8) };

    const html = buildMail(data, text, false);
    const subject = String(text['信件主旨'] || '你的代謝報告').replace(/\{型態\}/g, t.name);
    const opts = { to: to, subject: subject, htmlBody: html, name: text['信件寄件人'] || '營養師彥彥' };
    try {
      const pdf = Utilities.newBlob(buildMail(data, text, true), 'text/html', 'report.html').getAs('application/pdf').setName('代謝報告_' + t.name + '.pdf');
      opts.attachments = [pdf];
    } catch (e) { /* PDF 失敗就只寄信 */ }
    MailApp.sendEmail(opts);
    cache.put(ck, '1', 600);
    return '已寄出';
  } catch (e) {
    return '寄送失敗：' + String(e.message || e).slice(0, 80);
  }
}

function makeBreakfast(code,kcal,diet,meal){
  const vegan=diet==="全素", ovo=diet==="蛋奶素", carbs=kcal<1800?2:3;
  // 每一型的早餐組合：s 原型澱粉、milk 要不要鮮奶、pro 蛋白質、drink 不喝鮮奶時的飲品、out 外食選擇
  const PLAN={
    G:{s:"地瓜",milk:true,pro:["egg2"],out:"store"},
    P:{s:"燕麥",milk:false,pro:["egg1"],drink:"soy",out:"toast"},
    E:{s:"馬鈴薯",milk:false,pro:["tuna","egg1"],drink:"coffee",out:"crepe"},
    N:{s:"地瓜",milk:false,pro:["egg2"],drink:"soy",out:"store"},
    M:{s:"燕麥",milk:true,pro:["egg1","tuna"],out:"toast"},
    D:{s:"地瓜",milk:true,pro:["egg1","tuna"],out:"toast"},
    B:{s:"燕麥",milk:true,pro:["egg1"],out:"crepe"}}[code]||{s:"地瓜",milk:true,pro:["egg2"],out:"store"};
  const STARCH={地瓜:[55,"蒸地瓜",k=>k===1?"約半條小的":k===2?"約 1 條小的":"約 1 條中的"],馬鈴薯:[90,"水煮馬鈴薯",k=>k===1?"約半顆":k===2?"約 1 顆中的":"約 1 顆大的"],燕麥:[20,"大燕麥片（熱水或豆漿泡開）",k=>`約 ${k*3} 湯匙`]};
  const PRO={egg2:vegan?["板豆腐（煎或涼拌）","160 g",150]:["水煮蛋或荷包蛋","2 顆",150],
             egg1:vegan?["板豆腐（煎或涼拌）","80 g",75]:["水煮蛋或荷包蛋","1 顆",75],
             tuna:vegan||ovo?["鹽水毛豆","50 g",75]:["水煮鮪魚罐頭（瀝乾）","半罐",55]};
  const items=[],add=(g,n,a,k)=>{ items.push({g,n,a}); kc+=k; }; let kc=0;
  if(meal==="外食"){
    const o=vegan?"veg":ovo&&PLAN.out==="toast"?"crepe":PLAN.out;
    if(o==="toast"){ add("grain","里肌生菜蛋吐司（不加醬）","1 份",330); add("protein","無糖豆漿","1 杯",90); }
    else if(o==="crepe"){ add("grain","蔬菜蛋餅","1 份",280); add("protein","無糖豆漿","1 杯",90); }
    else { add("grain","超商 25 元地瓜","1 條",140); add("protein",vegan?"毛豆":"茶葉蛋",vegan?"1 包":"2 顆",150); add("protein","無糖豆漿","1 瓶",90); }
    if(carbs>2&&!vegan) add("dairy","低脂鮮奶","1 瓶（240 ml）",120);
    add("veg","店家的生菜或超商沙拉","1 份",25);
    const home=STARCH[PLAN.s][1].replace(/（.*）/,"");
    return {items,kcal:kc,alt:`自己準備的話：${home}＋${PRO[PLAN.pro[0]][0].replace(/（.*）/,"")}＋大番茄、小黃瓜`,tip:o==="toast"?"跟老闆說「不要醬」，美乃滋或甜辣醬一份就將近 100 大卡。":o==="crepe"?"蛋餅請店家少油、醬油膏另外放，沾一點就好。":"飲料選無糖豆漿，不要拿含糖奶茶。"};
  }
  const milk=PLAN.milk&&!vegan?1:0, k=carbs-milk, st=STARCH[PLAN.s];
  add("grain",st[1],`${st[0]*k} g（${st[2](k)}）`,70*k);
  PLAN.pro.forEach(x=>add("protein",PRO[x][0],PRO[x][1],PRO[x][2]));
  if(milk) add("dairy","低脂鮮奶","240 ml",120);
  else if(PLAN.drink==="coffee") add("drink","美式咖啡或無糖茶","1 杯",5);
  else add("protein","無糖豆漿","260 ml",90);
  add("veg","大番茄＋小黃瓜","1 碗",25);
  return {items,kcal:kc,alt:`外食的話：${vegan?"超商 25 元地瓜＋毛豆＋無糖豆漿":PLAN.out==="toast"&&!ovo?"不加醬的里肌生菜蛋吐司＋無糖豆漿":PLAN.out==="store"?"超商 25 元地瓜＋茶葉蛋 2 顆＋無糖豆漿":"蔬菜蛋餅＋無糖豆漿"}`,tip:""};
}

function buildMail(d, text, forPdf) {
  const e = v => String(v == null ? '' : v).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
  const br = v => e(String(v).replace(/\\n/g, '\n')).replace(/\*\*(.+?)\*\*/g, '<b style="color:#FF5C8D">$1</b>').replace(/\n/g, '<br>');
  const L = v => String(v || '').split(/[｜|\n]/).map(x => x.trim()).filter(Boolean);
  const C = { bg: '#FFF7F9', panel: '#FFFFFF', panel2: '#FFF0F4', ink: '#2E1A24', muted: '#7D6470', faint: '#B49AA6', line: '#F5DCE5', pink: '#FF5C8D', rose: '#FF3D7A' };
  const font = "'Noto Sans TC','PingFang TC','Microsoft JhengHei',sans-serif";
  const mono = "'DM Mono',Menlo,monospace";
  const h2 = (n, zh, en) => '<tr><td style="padding:28px 28px 10px"><div style="font-family:' + mono + ';font-size:11px;letter-spacing:2px;color:' + C.pink + '">' + n + ' · ' + en.toUpperCase() + '</div><div style="font-size:20px;font-weight:700;color:' + C.ink + ';margin-top:4px">' + zh + '</div></td></tr>';
  const p = (html, st) => '<tr><td style="padding:0 28px 12px;font-size:15px;line-height:1.8;color:' + C.muted + ';' + (st || '') + '">' + html + '</td></tr>';
  const t = d.t, sorted = d.main.slice().sort((a, b) => d.pct[b.code] - d.pct[a.code]);
  const bar = (x) => { const on = x.code === t.code, v = d.pct[x.code];
    return '<tr><td style="padding:3px 0;font-size:13.5px;color:' + (on ? C.ink : C.muted) + ';width:44%">No.' + ('0' + (d.main.indexOf(x) + 1)).slice(-2) + ' ' + e(x.name) + '</td>' +
      '<td style="padding:3px 8px"><div style="background:' + C.line + ';height:6px;border-radius:3px"><div style="width:' + v + '%;background:' + (on ? C.pink : C.faint) + ';height:6px;border-radius:3px"></div></div></td>' +
      '<td style="padding:3px 0;width:44px;text-align:right;font-family:' + mono + ';font-size:13px;color:' + (on ? C.pink : C.muted) + '">' + v + '%</td></tr>'; };
  const cell = (k, v, s) => '<td width="50%" style="padding:6px"><div style="background:' + C.panel2 + ';border-radius:14px;padding:14px 16px"><div style="font-size:12px;color:' + C.faint + '">' + k + '</div><div style="font-family:' + mono + ';font-size:22px;color:' + C.ink + ';margin:2px 0">' + v + '</div><div style="font-size:12px;color:' + C.muted + '">' + s + '</div></div></td>';
  const n = d.n;
  const bmiTag = b => b < 18.5 ? '可以再多吃一點，把身體養好' : b < 24 ? '在健康範圍內' : b < 27 ? '比健康範圍多一點點' : '還有一些調整空間';
  const sig = L(text['資歷']);
  const btn = (href, label) => '<a href="' + e(href) + '" style="display:inline-block;background:' + C.pink + ';color:#FFFFFF;text-decoration:none;font-weight:700;font-size:15px;padding:14px 28px;border-radius:99px">' + label + '</a>';

  let rows = '';
  // hero
  rows += '<tr><td style="padding:28px"><div style="background:' + C.bg + ';border:1px solid ' + C.line + ';border-radius:22px;padding:28px 24px;background-image:radial-gradient(120% 100% at 100% 0%,rgba(255,92,141,.16),transparent 60%)">' +
    '<div style="font-family:' + mono + ';font-size:11px;letter-spacing:2px;color:' + C.faint + '">YOUR METABOLIC TYPE · REPORT NO.' + e(d.reportNo) + '</div>' +
    '<div style="font-family:' + mono + ';font-size:14px;color:' + C.pink + ';margin-top:16px">No.' + d.no + '</div>' +
    '<div style="font-family:\'Noto Sans TC\',sans-serif;font-size:32px;font-weight:900;color:' + C.ink + ';line-height:1.3">' + e(t.name) + '</div>' +
    '<div style="font-family:\'Playfair Display\',serif;font-style:italic;font-size:16px;color:' + C.pink + '">' + e(t.en) + '</div>' +
    '<div style="font-size:15px;color:' + C.muted + ';margin-top:12px;line-height:1.8">' + e(t.tag) + '</div>' +
    (d.s ? '<div style="display:inline-block;margin-top:14px;border:1px solid ' + C.line + ';border-radius:99px;padding:5px 14px;font-size:13px;color:' + C.muted + '">混合 <b style="color:' + C.ink + '">No.' + d.sNo + ' ' + e(d.s.name) + '</b> 傾向</div>' : '') +
    '</div></td></tr>';
  rows += p(br(text['信件開頭']), 'color:' + C.ink);
  // body now
  rows += h2('01', '你的身體，現在大概是這樣', 'Where you are');
  rows += p(L(t.why).map((w, i) => '<div style="margin-bottom:8px"><span style="color:' + C.pink + ';font-family:' + mono + '">' + (i + 1) + '　</span>' + e(w) + '</div>').join(''));
  rows += '<tr><td style="padding:4px 28px 8px"><table width="100%" cellpadding="0" cellspacing="0">' + sorted.map(bar).join('') + '</table></td></tr>';
  if (d.s) rows += '<tr><td style="padding:10px 28px"><div style="background:' + C.panel2 + ';border-radius:14px;padding:14px 16px;font-size:14px;color:' + C.muted + ';line-height:1.8"><b style="color:' + C.ink + '">你是混合型。</b>主要是' + e(t.name) + '，也帶有 ' + d.pct[d.s.code] + '% 的' + e(d.s.name) + '傾向。調整時先從主要型態下手，另一型的地雷也順便避開。</div></td></tr>';
  const mines = L(t.mines);
  if (mines.length) rows += '<tr><td style="padding:12px 28px"><div style="border-left:2px solid ' + C.pink + ';padding:4px 0 4px 14px;font-size:14.5px;color:' + C.muted + ';line-height:1.8"><b style="color:' + C.pink + '">你最該先避開的地雷：</b><span style="color:' + C.ink + '">' + e(mines[0]) + '</span><br>' + (mines.length > 1 ? '其他常見的還有：' + mines.slice(1).map(e).join('、') : '') + '</div></td></tr>';
  // protocol
  rows += h2('02', e(t.name) + '的調整方向', 'Protocol');
  rows += p('<b style="color:' + C.ink + '">飲食</b><br>' + L(t.food).map(x => '・' + e(x)).join('<br>') + '<br><br><b style="color:' + C.ink + '">生活</b><br>' + L(t.life).map(x => '・' + e(x)).join('<br>'));
  if (t.core) rows += p('<i style="font-family:\'Noto Sans TC\',sans-serif;font-size:17px;color:' + C.pink + '">「' + e(t.core) + '」</i>');
  // numbers
  if (n) {
    rows += h2('03', '你每天該吃多少', 'Your numbers');
    rows += '<tr><td style="padding:0 22px"><table width="100%" cellpadding="0" cellspacing="0"><tr><td colspan="2" style="padding:6px"><div style="background:' + C.bg + ';border:1px solid ' + C.pink + ';border-radius:14px;padding:16px 18px"><div style="font-size:12px;color:' + C.faint + '">建議熱量</div><div style="font-family:' + mono + ';font-size:30px;color:' + C.pink + '">' + n.kcal.toLocaleString() + ' <span style="font-size:13px;color:' + C.muted + '">大卡／天</span></div><div style="font-size:12.5px;color:' + C.muted + '">' + e(n.note) + '</div></div></td></tr>' +
      '<tr>' + cell('BMI', e(n.bmi), bmiTag(+n.bmi)) + cell('體脂率', n.bf ? n.bf + '%' : '—', '建議 ' + (n.sex === '男' ? 25 : 30) + '% 以內') + '</tr>' +
      '<tr>' + cell('理想體重', n.ideal[0] + '–' + n.ideal[1] + ' kg', '健康 BMI 的範圍') + cell('腰圍', n.waist ? n.waist + ' cm' : '—', '建議 ' + (n.sex === '男' ? 90 : 80) + ' cm 以內') + '</tr></table></td></tr>';
  }
  // breakfast
  rows += h2(n ? '04' : '03', '明天早餐這樣吃', "Tomorrow's breakfast");
  rows += '<tr><td style="padding:0 28px">' + d.bk.items.map(x => '<div style="background:' + C.panel2 + ';border-radius:14px;padding:12px 16px;margin-bottom:8px"><table width="100%" cellpadding="0" cellspacing="0"><tr><td style="font-size:15px;color:' + C.ink + ';font-weight:600">' + e(x.n) + '</td><td style="text-align:right;font-family:' + mono + ';color:' + C.pink + ';font-size:13.5px">' + e(x.a) + '</td></tr></table></div>').join('') + '</td></tr>';
  rows += p('<span style="font-size:13.5px">' + e(d.bk.alt) + (d.bk.tip ? '<br><b style="color:' + C.ink + '">點餐小提醒　</b>' + e(d.bk.tip) : '') + '</span>');
  if (t.bkWhy) rows += '<tr><td style="padding:0 28px 12px"><div style="background:' + C.panel2 + ';border-radius:14px;padding:14px 16px"><div style="font-family:' + mono + ';font-size:11px;letter-spacing:2px;color:' + C.pink + '">WHY · 為什麼這樣吃</div><div style="font-size:14.5px;color:' + C.ink + ';line-height:1.8;margin-top:4px">' + e(t.bkWhy) + '</div></div></td></tr>';
  if (t.menuTip) rows += p('<div style="border-left:2px solid ' + C.pink + ';padding-left:14px"><b style="color:' + C.pink + '">提醒　</b>' + e(t.menuTip) + '</div>');
  // CTA
  rows += '<tr><td style="padding:24px 28px"><div style="border:1px solid ' + C.pink + ';border-radius:22px;padding:24px;text-align:center">' +
    '<div style="font-size:19px;font-weight:700;color:' + C.ink + '">' + e(text['結果CTA標題']) + '</div>' +
    '<div style="font-size:14.5px;color:' + C.muted + ';line-height:1.8;margin:10px 0 18px">' + br(text['信件結尾']) + '</div>' +
    btn(text['IG網址'] || SITE_URL, e(text['信件按鈕'] || 'IG 私訊「代謝」')) + '</div></td></tr>';
  if (!forPdf) rows += '<tr><td style="padding:0 28px 8px;text-align:center"><a href="' + e(d.link) + '" style="color:' + C.pink + ';font-size:14px">回到我的完整報告 →</a></td></tr>';
  // signature
  rows += '<tr><td style="padding:22px 28px;border-top:1px solid ' + C.line + '"><div style="font-family:\'Noto Sans TC\',sans-serif;font-size:17px;color:' + C.ink + ';font-weight:700">' + e(text['品牌名稱']) + '</div>' +
    '<div style="font-size:12.5px;color:' + C.faint + ';line-height:1.8;margin-top:4px">' + sig.map(e).join('<br>') + '</div>' +
    '<div style="font-size:12.5px;margin-top:6px"><a href="' + e(text['IG網址']) + '" style="color:' + C.pink + '">' + e(text['IG帳號']) + '</a></div></td></tr>';
  rows += '<tr><td style="padding:14px 28px 24px;font-size:11.5px;color:' + C.faint + ';line-height:1.7">' + e(text['免責聲明']) + (forPdf ? '' : '<br>不想再收到信？直接回覆這封信告訴我，我會把你的資料刪除。') + '</td></tr>';

  return '<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width"></head>' +
    '<body style="margin:0;padding:0;background:' + C.bg + '"><table width="100%" cellpadding="0" cellspacing="0" style="background:' + C.bg + '"><tr><td align="center" style="padding:' + (forPdf ? '0' : '20px 10px') + '">' +
    '<table width="100%" cellpadding="0" cellspacing="0" style="max-width:600px;background:' + C.panel + ';border:1px solid ' + C.line + ';border-radius:24px;font-family:' + font + '">' + rows + '</table></td></tr></table></body></html>';
}

// 第一次使用：在編輯器選這個函式按「執行」，授權寄信，並寄一封範例信給你自己
function testMail() {
  const me = Session.getEffectiveUser().getEmail();
  CacheService.getScriptCache().remove('sent_' + me.toLowerCase());
  const r = sendReport(me, { main: 'G', second: 'M', pct: { G: 72, M: 54, P: 46, N: 41, E: 30, D: 18 },
    numbers: { kcal: 1400, bmi: 24.2, sex: '女', bf: 31, waist: 82, ideal: [47, 61], note: '依國健署的體位與活動量建議計算' }, diet: '葷食', meal: '自煮', link: SITE_URL, reportNo: '000000' });
  Logger.log(me + '：' + r);
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
