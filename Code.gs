/****************************************************************
 * 營養師彥彥｜代謝卡關檢測 後端（Google Apps Script）
 * ---------------------------------------------------------------
 * 第一次使用：
 *   1. 把下面的 ADMIN_KEY 改成你自己的後台密碼
 *   2. 上方函式選單選「setup」→ 按「執行」→ 授權
 *      會自動建立 4 個分頁：填答資料、題目、型態、網站文案
 *   3. 「部署」→「新增部署作業」→ 類型選「網頁應用程式」
 *      執行身分：我　／　存取權：所有人
 * 之後在「題目」「型態」「網站文案」分頁改內容，網站就會自動更新。
 ****************************************************************/

const ADMIN_KEY = '請改成你的後台密碼';   // ← 改這裡，至少 8 個字，不要用生日

const SHEET_DATA = '填答資料', SHEET_Q = '題目', SHEET_TYPE = '型態', SHEET_TEXT = '網站文案';
const Q_HEADERS = ['啟用', '維度', '題目', '情境小語', '選項A', 'A加分', '選項B', 'B加分', '選項C', 'C加分', '選項D', 'D加分'];
const T_HEADERS = ['代號', '名稱', '英文名稱', '一句話', '卡關原因', '常踩地雷', '飲食調整', '生活調整', '核心句', '菜單提醒'];
const BASE_HEADERS = ['編號', '填寫時間', '狀態', '備註'];

const DEFAULT_TEXT = {
  "品牌名稱": "營養師彥彥",
  "品牌英文": "Metabolic Block Assessment",
  "首頁徽章": "3 分鐘 · 6 型代謝卡關檢測",
  "首頁前言": "瘦不下來，卡住你的不是意志力——",
  "首頁標題": "是這 **6 種代謝卡關**\n的其中一種",
  "首頁英文": "It's not willpower. It's your metabolic pattern.",
  "首頁說明": "同樣少吃多動，有人瘦、有人卡住，差別在**你是哪一型**。回答 16 題，找出你的代謝卡關型態，拿到每日熱量建議和一份第一天菜單。",
  "開始按鈕": "免費檢測，看看我是哪一型",
  "資歷": "臺北醫學大學｜肥胖研究學會 體重管理營養師｜CTSSN 運動營養專業認證｜體適能健身 C 級指導員",
  "IG網址": "https://www.instagram.com/yan_dietitian/",
  "IG帳號": "@yan_dietitian",
  "LINE網址": "",
  "健康狀況選項": "血壓偏高,血糖偏高或胰島素阻抗,多囊性卵巢（PCOS）,甲狀腺功能低下,血脂偏高,便秘或腸胃不適,經期不規則,長期失眠,懷孕或哺乳中,腎臟或其他慢性病",
  "解鎖標題": "你的代謝報告已經生成",
  "解鎖說明": "留下 IG 就能解鎖。我會親自看過你的結果，有需要會私訊你。",
  "同意條款": "我同意營養師彥彥蒐集我填寫的資料（包含健康狀況），只用於提供營養建議與後續聯繫，不會提供給第三方，我可以隨時要求刪除。",
  "結果CTA標題": "想知道你這一型，怎麼瘦才最有效？",
  "結果CTA說明": "報告只是起點。私訊我「代謝」兩個字，跟我分享你的作息和飲食習慣，我會依照你的型態，給你更深入、更適合你的減脂建議。",
  "免責聲明": "本檢測依你的作答整理出「傾向」，僅供生活參考，不是醫療診斷。懷孕、哺乳、正在用藥或有慢性病者，請先諮詢醫師或營養師。"
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
  "menuTip": "今天每一餐都準時吃，下午那份點心不要省，避免晚上補償性大吃。"
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
  "menuTip": "今天每一餐都放下手機，專心吃完再做別的事。"
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
  "menuTip": "晚餐已經排了足夠的蛋白質和菜，吃完之後就讓廚房打烊。"
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
  "menuTip": "午餐和晚餐後各散步 10 分鐘，讓這份菜單發揮最大效果。"
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
  "menuTip": "每一餐都照「菜、肉、飯」的順序吃，下午點心選水果配堅果。"
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
  "menuTip": "今天的份量照表吃完，不要再自己減量，讓身體知道食物是夠的。"
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
   "照菜單的份量吃",
   "每週嘗試一種新的蔬菜或蛋白質"
  ],
  "life": [
   "維持現在的運動習慣",
   "每週量一次體重和腰圍"
  ],
  "core": "穩定，就是最強的優勢。",
  "menuTip": "你的節奏很好，今天照份量吃，好好享受食物。"
 }
];
const DEFAULT_QS = [{"dim": "Body Signals · 身體訊號", "t": "脂肪最容易堆在哪裡？", "s": "身體囤脂的位置，會透露一些線索。", "o": [["沒有特別明顯的地方", ""], ["肚子和腰圍，四肢相對細", "P2 G2"], ["全身軟軟的，沒什麼肌肉線條", "M3 D1"], ["早晚體重差很多，容易水腫", "N2 M1"]]}, {"dim": "Body Signals · 身體訊號", "t": "下午三、四點的你，通常是？", "s": "下午的精神狀態，和血糖、睡眠都有關係。", "o": [["精神還不錯", ""], ["很睏，特別想吃甜的", "G3"], ["靠咖啡或提神飲料撐著", "N2 P1"], ["頭暈手抖，一定要馬上吃東西", "G2 D1"]]}, {"dim": "Body Signals · 身體訊號", "t": "關於體溫和手腳，哪個最像你？", "s": "燃燒的效率，常常反映在體溫上。", "o": [["正常，不特別怕冷怕熱", ""], ["手腳常冰冷，比別人怕冷", "M2 D2"], ["一緊張就容易冒汗、發熱", "P3"], ["晚上反而比較有精神、身體發熱", "N3"]]}, {"dim": "Eating Patterns · 飲食行為", "t": "肚子餓的時候，你的反應是？", "s": "飢餓的樣子，每個人都不一樣。", "o": [["慢慢有感覺，可以等到正餐", ""], ["心慌手抖、脾氣變差，必須馬上吃", "G3"], ["常常不太餓，會忘記吃", "D3"], ["肚子不餓，但嘴巴就是想吃", "E3"]]}, {"dim": "Eating Patterns · 飲食行為", "t": "壓力很大的那幾天，你的飲食會？", "s": "壓力，會悄悄改變你想吃的東西。", "o": [["跟平常差不多", ""], ["吃不下，胃緊緊的", "P3"], ["特別想吃炸物、甜點", "E2 P1"], ["邊忙邊吃，回神整包都吃完了", "E2 G1"]]}, {"dim": "Eating Patterns · 飲食行為", "t": "你對甜食和澱粉的渴望？", "s": "嘴饞，常常不只是因為肚子餓。", "o": [["偶爾想吃，吃一點就滿足", ""], ["正餐吃飽了，還是要來點甜的", "G3"], ["心情不好時特別想吃", "E3"], ["越晚越想吃重口味或泡麵", "N2 E1"]]}, {"dim": "Eating Patterns · 飲食行為", "t": "手搖飲、含糖飲料，一週喝幾杯？", "s": "喝下去的糖，最容易被忽略。", "o": [["幾乎不喝", ""], ["1 到 2 杯", "G1"], ["3 到 5 杯", "G2"], ["幾乎天天", "G3"]]}, {"dim": "Eating Patterns · 飲食行為", "t": "每一餐都有一掌心的豆魚蛋肉嗎？", "s": "蛋白質不夠，肌肉會悄悄流失。", "o": [["每餐都有", ""], ["大部分有", "D1"], ["偶爾才有", "D2 M1"], ["很少，常常只吃澱粉", "D2 G1"]]}, {"dim": "Energy & Metabolism · 能量與代謝", "t": "吃完一頓正餐後，你通常？", "s": "飯後的感覺，是身體給你的回饋。", "o": [["精神飽滿", ""], ["很想睡，像被關機", "G3"], ["脹氣，消化很慢", "P2 E1"], ["很快又餓了", "G2 D1"]]}, {"dim": "Energy & Metabolism · 能量與代謝", "t": "誠實說，你過去的減重經驗？", "s": "過去用過的方法，會在身體留下痕跡。", "o": [["沒有刻意減過，或減得很順", ""], ["吃得很少，體重卻不動", "D2 M2"], ["壓力一大就變胖", "P3"], ["減了又復胖，來回好幾次", "D2 E1"]]}, {"dim": "Energy & Metabolism · 能量與代謝", "t": "一天坐著的時間？（工作、通勤、滑手機）", "s": "身體的燃燒，大多發生在你沒注意的時候。", "o": [["不到 6 小時", ""], ["6 到 8 小時", "M1"], ["8 到 10 小時", "M2"], ["10 小時以上", "M3"]]}, {"dim": "Energy & Metabolism · 能量與代謝", "t": "一週運動幾次？（30 分鐘以上、會喘）", "s": "肌肉，是身上最會燃燒熱量的地方。", "o": [["4 次以上", ""], ["2 到 3 次", "M1"], ["1 次", "M2"], ["幾乎沒有", "M3"]]}, {"dim": "Energy & Metabolism · 能量與代謝", "t": "你對咖啡因的依賴程度？", "s": "提神，有時候是在透支。", "o": [["不喝也沒差", ""], ["每天 1 杯，習慣而已", "N1"], ["沒喝就沒辦法工作", "N2 P1"], ["喝了會心悸、睡不好，但還是要喝", "P2 N1"]]}, {"dim": "Sleep & Stress · 作息與壓力", "t": "你通常幾點睡？", "s": "睡眠，是最被低估的代謝工具。", "o": [["11 點前", ""], ["11 點到 12 點", "N1"], ["12 點到 1 點", "N2"], ["1 點以後", "N3"]]}, {"dim": "Sleep & Stress · 作息與壓力", "t": "最近的睡眠品質？", "s": "睡得好不好，隔天的食慾都知道。", "o": [["一覺到天亮", ""], ["身體很累，但腦袋停不下來", "P3"], ["睡不夠，白天很難清醒", "N3"], ["半夜容易醒來", "P1 N2"]]}, {"dim": "Sleep & Stress · 作息與壓力", "t": "你目前最大的壓力來源？", "s": "壓力從哪裡來，決定了它怎麼影響你。", "o": [["還好，沒有特別的壓力", ""], ["工作或課業", "P3"], ["感情、家人或人際關係", "E3"], ["對自己的身材不滿意，一直在減肥", "D2 E1"]]}];

/* ---------- 第一次設定 ---------- */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss.getSheetByName(SHEET_TEXT)) {
    const t = ss.insertSheet(SHEET_TEXT);
    const rows = [['項目', '內容', '說明']].concat(Object.keys(DEFAULT_TEXT).map(k => [k, DEFAULT_TEXT[k], hintFor(k)]));
    t.getRange(1, 1, rows.length, 3).setValues(rows);
    styleHeader(t, 3); t.setColumnWidth(1, 130); t.setColumnWidth(2, 520); t.setColumnWidth(3, 320);
    t.getRange(2, 2, rows.length, 2).setWrap(true);
  }

  if (!ss.getSheetByName(SHEET_TYPE)) {
    const t = ss.insertSheet(SHEET_TYPE);
    const rows = [T_HEADERS].concat(DEFAULT_TYPES.map(x => [x.code, x.name, x.en, x.tag, x.why.join('｜'), x.mines.join('｜'), x.food.join('｜'), x.life.join('｜'), x.core, x.menuTip]));
    t.getRange(1, 1, rows.length, T_HEADERS.length).setValues(rows);
    styleHeader(t, T_HEADERS.length); t.setFrozenColumns(2);
    for (let c = 4; c <= 10; c++) t.setColumnWidth(c, 280);
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
  })[k] || '';
}

/* ---------- 網站讀取內容 ---------- */
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  if (action === 'config') return json(getConfig());
  return json({ ok: true, message: '連線成功' });
}

function getConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const read = (name, n) => { const s = ss.getSheetByName(name); return s && s.getLastRow() > 1 ? s.getRange(2, 1, s.getLastRow() - 1, n).getDisplayValues() : []; };

  const text = {};
  read(SHEET_TEXT, 2).forEach(r => { if (r[0]) text[r[0].trim()] = r[1]; });

  const types = read(SHEET_TYPE, T_HEADERS.length).filter(r => r[0] && r[1]).map(r => ({
    code: r[0].trim().toUpperCase(), name: r[1], en: r[2], tag: r[3],
    why: r[4], mines: r[5], food: r[6], life: r[7], core: r[8], menuTip: r[9]
  }));

  const questions = [];
  read(SHEET_Q, Q_HEADERS.length).forEach(r => {
    if (r[0] === '否' || !r[2]) return;
    const o = [];
    for (let i = 0; i < 4; i++) { const txt = r[4 + i * 2].trim(); if (txt) o.push([txt, r[5 + i * 2]]); }
    if (o.length >= 2) questions.push({ dim: r[1], t: r[2], s: r[3], o: o });
  });
  return { ok: true, text: text, types: types, questions: questions };
}

/* ---------- 收資料 / 後台 ---------- */
function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents || '{}'); } catch (err) { return json({ ok: false, error: '資料格式錯誤' }); }
  if (body.action === 'submit') return json(submit(body.data || {}));

  if (!body.key || body.key !== ADMIN_KEY || ADMIN_KEY === '請改成你的後台密碼') {
    Utilities.sleep(800);
    return json({ ok: false, error: '密碼錯誤' });
  }
  if (body.action === 'list') return json(listRows());
  if (body.action === 'update') return json(updateRow(body.id, body.status, body.note));
  return json({ ok: false, error: '未知的動作' });
}

function submit(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss.getSheetByName(SHEET_DATA)) setup();
    const sh = ss.getSheetByName(SHEET_DATA);
    const row = Object.assign({
      '編號': Utilities.getUuid().slice(0, 8),
      '填寫時間': Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm'),
      '狀態': '未聯繫', '備註': ''
    }, data);
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
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_DATA);
  if (!sh || sh.getLastRow() < 2) return { ok: true, rows: [] };
  const values = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getDisplayValues();
  const headers = values.shift();
  const rows = values.map(r => { const o = {}; headers.forEach((h, i) => { if (h) o[h] = r[i]; }); return o; });
  return { ok: true, rows: rows.reverse() };
}

function updateRow(id, status, note) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_DATA);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idCol = headers.indexOf('編號') + 1, stCol = headers.indexOf('狀態') + 1, noteCol = headers.indexOf('備註') + 1;
  const ids = sh.getRange(2, idCol, Math.max(sh.getLastRow() - 1, 1), 1).getValues().map(r => String(r[0]));
  const i = ids.indexOf(String(id));
  if (i === -1) return { ok: false, error: '找不到這筆資料' };
  if (status !== undefined) sh.getRange(i + 2, stCol).setValue(clean(String(status)));
  if (note !== undefined) sh.getRange(i + 2, noteCol).setValue(clean(String(note)));
  return { ok: true };
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
