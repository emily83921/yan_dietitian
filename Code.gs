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
const T_HEADERS = ['代號', '名稱', '英文名稱', '一句話', '卡關原因', '常踩地雷', '飲食調整', '生活調整', '核心句', '菜單提醒', '早餐原因', '卡關位置', '卡關圖示', '卡關一句', '換個吃法', '原因圖示'];
const BASE_HEADERS = ['編號', '填寫時間', '狀態', '報告寄送', '備註'];
const SHEET_ID = '1LWcy7e5Oe6kcSZPDLhZhbvUQyp8_ghPetr155U3HY4g';   // 你的 Google 試算表 ID（網址 /d/ 後面那串）
function SS() { return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet(); }
const SITE_URL = 'https://emily83921.github.io/yan_dietitian/';   // 信件裡「回到我的完整報告」只允許連回這個網址

const DEFAULT_TEXT = {
  "品牌名稱": "營養師彥彥",
  "品牌英文": "Brighter Me 輕盈閃耀",
  "Slogan英文": "Lighter every day, brighter every way",
  "Slogan中文": "每天輕盈一點，遇見更閃耀的自己",
  "品牌故事": "不論你是正在執行減脂計畫、培養運動習慣，還是純粹想在日常生活中建立起一個精緻、有儀式感的小習慣，每一步微小的累積，都會讓你在不知不覺中散發出迷人的光芒 🌟",
  "首頁徽章": "60 秒 · 6 型代謝卡關檢測",
  "首頁前言": "少吃、多動，體重還是不動？",
  "首頁標題": "也許，是你的代謝\n**卡關了**",
  "首頁英文": "Find your block. Then fix it.",
  "首頁說明": "同樣少吃多動，有人慢慢瘦下來，有人卻一直卡住——這往往不是你不夠努力，只是**還沒找到身體真正卡住的地方**。花 60 秒認識自己的代謝類型，拿到專屬於你的分析報告：每日熱量建議、早餐怎麼吃，完整報告會寄到你的信箱。",
  "開始按鈕": "免費檢測，找出我的卡關點",
  "資歷": "臺北醫學大學 Taipei Medical University｜肥胖研究學會 體重管理營養師｜CTSSN 運動營養專業認證｜體適能健身 C 級指導員",
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
  "信件寄件人": "營養師彥彥｜Brighter Me",
  "信件主旨": "你的代謝報告好了——你是「{型態}」",
  "信件開頭": "先跟你說一件事：**你會卡住，不是因為你不夠努力。**\n每個人的身體都有自己的節奏，只要找到卡住的那個地方，瘦下來會比你想像的輕鬆。",
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
  "menuTip": "咖啡放在早餐後再喝，不要空腹靠咖啡撐到中午。",
  "area": "腹部 × 壓力",
  "areaIcon": "⚡",
  "areaLine": "身體一直在備戰，脂肪守在肚子",
  "swaps": "靠咖啡硬撐→早餐先吃，再喝咖啡｜忙到忘記吃、晚上大吃→下午先備一份點心｜睡不夠還硬做高強度→改成散步或伸展｜壓力大就吃甜食→想吃甜就配無糖茶，慢慢吃一小份",
  "whyIcons": "⚡｜🍟｜😵‍💫"
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
  "menuTip": "坐下來吃，放下手機，吃 15 分鐘以上。",
  "area": "大腦 × 情緒",
  "areaIcon": "🧠",
  "areaLine": "餓的不是胃，是心情",
  "swaps": "追劇配洋芋片→毛豆、無糖優格，先分裝一小碗｜心情差就叫外送→先喝一杯水，等 3 分鐘再決定｜用甜點獎勵自己→甜點照吃，放在飯後、分一半｜整包吃完才回神→放下手機，一口一口慢慢吃",
  "whyIcons": "😮‍💨｜📱｜🔁"
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
  "menuTip": "就算前一晚晚睡，也在差不多的時間吃。",
  "area": "生理時鐘",
  "areaIcon": "🌙",
  "areaLine": "身體的時間，跟生活對不上",
  "swaps": "熬夜追劇→提早 30 分鐘關螢幕｜深夜泡麵、鹹酥雞→晚餐蛋白質吃夠，餓了喝溫豆漿｜睡前一杯手搖飲→換成溫熱的無糖豆漿或鮮奶｜睡過頭跳過早餐→超商茶葉蛋＋地瓜帶著走",
  "whyIcons": "😴｜🌙｜☕"
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
  "menuTip": "吃完可以提早一站下車，走 10 分鐘再進辦公室。",
  "area": "肌肉 × 循環",
  "areaIcon": "🔋",
  "areaLine": "不是吃太多，是燃燒得太少",
  "swaps": "吃完就坐著或躺著→飯後走 10 分鐘｜通勤全程都坐著→提早一站下車走路｜只節食不運動→吃夠，每週加兩次肌力訓練｜週末宅一整天→約朋友散步或逛街",
  "whyIcons": "🪑｜💪｜🧊"
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
  "menuTip": "吃的順序：蔬菜和蛋白質先吃，澱粉留到最後。",
  "area": "血糖",
  "areaIcon": "🎢",
  "areaLine": "能量忽高忽低，跟著血糖坐雲霄飛車",
  "swaps": "全糖手搖飲→每週降一級甜度｜麵包當早餐→地瓜或燕麥＋蛋｜飯後一定要甜點→甜點照吃，先吃完正餐、分一半｜下午三點的餅乾→一小把堅果或無糖優格",
  "whyIcons": "🧋｜😵｜📈"
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
  "menuTip": "請照份量吃完，不要自己再減。",
  "area": "整體代謝",
  "areaIcon": "🪫",
  "areaLine": "吃得太少，身體開始省電",
  "swaps": "一杯咖啡當早餐→咖啡配蛋和地瓜｜只吃沙拉→沙拉加一掌心蛋白質和一份澱粉｜極端斷食→三餐定時，份量吃夠｜天天量體重→一週量一次，多看腰圍",
  "whyIcons": "🪫｜🥚｜🔁"
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
  "menuTip": "好好享受食物，偶爾換換不同的蛋白質來源。",
  "area": "目前很穩定",
  "areaIcon": "✨",
  "areaLine": "節奏很好，只差一點細節",
  "swaps": "聚餐大吃→先吃菜和肉，飯減半｜忙起來忘了喝水→水瓶放在看得到的地方",
  "whyIcons": "✅｜🍽️｜📏"
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
    const rows = [T_HEADERS].concat(DEFAULT_TYPES.map(x => [x.code, x.name, x.en, x.tag, x.why.join('｜'), x.mines.join('｜'), x.food.join('｜'), x.life.join('｜'), x.core, x.menuTip, x.bkWhy, x.area, x.areaIcon, x.areaLine, x.swaps, x.whyIcons]));
    t.getRange(1, 1, rows.length, T_HEADERS.length).setValues(rows);
    styleHeader(t, T_HEADERS.length); t.setFrozenColumns(2);
    for (let c = 4; c <= 16; c++) t.setColumnWidth(c, 280);
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
    '品牌故事': '會出現在報告最下方和信件簽名旁',
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
    why: r[4], mines: r[5], food: r[6], life: r[7], core: r[8], menuTip: r[9], bkWhy: r[10] || '', area: r[11] || '', areaIcon: r[12] || '', areaLine: r[13] || '', swaps: r[14] || '', whyIcons: r[15] || ''
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
  if (body.action === 'book') return json(book(body.data || {}));

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

/* ---------- 預約一對一諮詢 ---------- */
const SHEET_BOOK = '預約諮詢';
function book(data) {
  const keep = ['稱呼', '聯絡方式', '聯絡帳號', '型態', '來源'];
  const d = {}; keep.forEach(k => { d[k] = String(data[k] || '').slice(0, 80); });
  if (!d['稱呼'] || !d['聯絡帳號']) return { ok: false, error: '資料不完整' };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SS();
    let sh = ss.getSheetByName(SHEET_BOOK);
    const headers = ['編號', '預約時間', '狀態', '備註'].concat(keep);
    if (!sh) { sh = ss.insertSheet(SHEET_BOOK, 1); sh.getRange(1, 1, 1, headers.length).setValues([headers]); styleHeader(sh, headers.length); sh.setColumnWidth(7, 200); }
    const when = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm');
    const row = { '編號': Utilities.getUuid().slice(0, 8), '預約時間': when, '狀態': '未聯繫', '備註': '' };
    keep.forEach(k => row[k] = d[k]);
    sh.appendRow(headers.map(h => clean(row[h])));
  } finally { lock.releaseLock(); }
  // 有新預約，寄一封通知給自己（寄件額度不夠就略過）
  try {
    if (MailApp.getRemainingDailyQuota() > 5) {
      MailApp.sendEmail(Session.getEffectiveUser().getEmail(), '【新預約】' + d['稱呼'] + '（' + (d['型態'] || '未填型態') + '）',
        '有人預約一對一營養諮詢：\n\n稱呼：' + d['稱呼'] + '\n聯絡方式：' + d['聯絡方式'] + '　' + d['聯絡帳號'] + '\n型態：' + (d['型態'] || '—') + '\n\n完整名單在試算表「' + SHEET_BOOK + '」分頁。');
    }
  } catch (e) { /* 通知失敗不影響預約 */ }
  return { ok: true };
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
  const br = v => e(String(v).replace(/\\n/g, '\n')).replace(/\*\*(.+?)\*\*/g, '<b style="color:#FF3D7A">$1</b>').replace(/\n/g, '<br>');
  const L = v => String(v || '').split(/[｜|\n]/).map(x => x.trim()).filter(Boolean);
  const C = { ink: '#2E1A24', muted: '#7D6470', faint: '#B49AA6', line: '#F5DCE5', pink: '#FF5C8D', rose: '#FF3D7A', tint: '#FFE6EE', p2: '#FFF0F4', peach: '#FFF3EA' };
  const font = "'Noto Sans TC','PingFang TC','Microsoft JhengHei',sans-serif", mono = "Menlo,Consolas,monospace";
  const IMG = SITE_URL.replace(/\/?$/, '/');
  const W = 680, PAD = 32, INNER = W - PAD * 2;
  const t = d.t, sorted = d.main.slice().sort((a, b) => d.pct[b.code] - d.pct[a.code]);
  const noOf = x => d.main.indexOf(x) + 1;
  const grad = 'background:#FF5C8D;background-image:linear-gradient(135deg,#FF5C8D,#FF3D7A);';
  const soft = 'background:#FFE6EE;background-image:linear-gradient(135deg,#FFE6EE,#FFF3EA);';
  const sec = (n, en, zh) => '<tr><td style="padding:30px ' + PAD + 'px 12px"><div style="font-family:' + mono + ';font-size:11px;letter-spacing:2px;color:' + C.pink + '">' + n + ' · ' + en.toUpperCase() + '</div><div style="font-size:21px;font-weight:900;color:' + C.ink + ';margin-top:4px">' + zh + '</div></td></tr>';
  const row = html => '<tr><td style="padding:0 ' + PAD + 'px">' + html + '</td></tr>';
  // 等寬格子：cells = 每格 HTML
  const grid = (cells, cols, gap, style) => {
    let h = '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:separate;border-spacing:' + gap + 'px">';
    for (let i = 0; i < cells.length; i += cols) {
      h += '<tr>';
      for (let j = 0; j < cols; j++) { const c = cells[i + j]; h += '<td valign="top" width="' + Math.floor(100 / cols) + '%" style="' + (c == null ? '' : style) + '">' + (c == null ? '' : c) + '</td>'; }
      h += '</tr>';
    }
    return h + '</table>';
  };
  const n = d.n;
  const bmiTag = b => b < 18.5 ? '可以再多吃一點' : b < 24 ? '在健康範圍內' : b < 27 ? '比健康範圍多一點點' : '還有一些調整空間';
  const ICON = { '地瓜': '🍠', '馬鈴薯': '🥔', '燕麥': '🥣', '吐司': '🥪', '蛋餅': '🌯', '豆漿': '🥛', '鮮奶': '🥛', '咖啡': '☕', '鮪魚': '🐟', '蛋': '🥚', '豆腐': '🍲', '毛豆': '🫛', '番茄': '🍅', '沙拉': '🥗', '生菜': '🥗' };
  const iconOf = name => { for (const k in ICON) if (name.indexOf(k) >= 0) return ICON[k]; return '🍽️'; };
  const swaps = L(t.swaps).map(x => x.split(/→|->/).map(y => y.trim())).filter(x => x.length === 2 && x[0] && x[1]);
  const chartCfg = { type: 'radar', data: { labels: d.main.map(x => x.name.slice(0, 2) + ' ' + d.pct[x.code] + '%'), datasets: [{ data: d.main.map(x => d.pct[x.code]), backgroundColor: 'rgba(255,92,141,0.22)', borderColor: '#FF5C8D', borderWidth: 2, pointRadius: 4, pointBackgroundColor: '#FF3D7A' }] },
    options: { legend: { display: false }, scale: { ticks: { display: false, min: 0, max: 100, stepSize: 25 }, gridLines: { color: '#F5DCE5' }, angleLines: { color: '#F5DCE5' }, pointLabels: { fontSize: 14, fontColor: '#7D6470' } } } };
  const chartUrl = 'https://quickchart.io/chart?w=300&h=270&bkg=white&c=' + encodeURIComponent(JSON.stringify(chartCfg));

  let r = '';
  // 頂部
  r += '<tr><td style="padding:18px ' + PAD + 'px;border-bottom:1px solid ' + C.line + '"><table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr><td style="font-size:14px;color:' + C.ink + '"><b style="font-weight:900">' + e(text['品牌名稱']) + '</b>&nbsp; <i style="color:' + C.rose + ';font-family:Georgia,serif">' + e(text['品牌英文']) + '</i></td><td align="right" style="font-family:Georgia,serif;font-style:italic;font-size:12.5px;color:' + C.muted + '">' + e(text['Slogan英文']) + '</td></tr></table></td></tr>';
  // 主卡
  r += '<tr><td style="padding:24px ' + PAD + 'px 0"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="' + soft + 'border-radius:24px"><tr>' +
    '<td valign="middle" style="padding:28px 0 28px 28px">' +
      '<table role="presentation" cellpadding="0" cellspacing="0"><tr><td><img src="' + IMG + 'yan-face.jpg" width="40" height="40" style="display:block;border-radius:50%;border:2px solid #fff" alt=""></td><td style="padding-left:10px;font-size:14px;color:' + C.muted + '">Hi，我是彥彥，你的報告我看過了</td></tr></table>' +
      '<div style="font-size:16px;color:' + C.ink + ';margin-top:14px;font-weight:500">你現在卡在這一關 ——</div>' +
      '<div style="font-family:' + mono + ';color:' + C.rose + ';font-size:14px;margin-top:6px">No.' + d.no + '</div>' +
      '<div style="font-size:38px;font-weight:900;color:' + C.ink + ';line-height:1.2">' + e(t.name) + '</div>' +
      '<div style="font-family:Georgia,serif;font-style:italic;color:' + C.rose + ';font-size:16px">' + e(t.en) + '</div>' +
      (d.s ? '<div style="display:inline-block;margin-top:12px;background:#fff;border-radius:99px;padding:5px 14px;font-size:13px;color:' + C.muted + '">同時帶有 <b style="color:' + C.ink + '">No.' + d.sNo + ' ' + e(d.s.name) + '</b> 傾向</div>' : '') +
    '</td>' +
    (t.area ? '<td valign="middle" width="170" style="padding:20px 22px 20px 12px"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:20px"><tr><td align="center" style="padding:18px 10px">' +
      '<div style="width:54px;height:54px;line-height:54px;border-radius:50%;background:' + C.tint + ';font-size:26px;text-align:center;margin:0 auto 8px">' + e(t.areaIcon) + '</div>' +
      '<div style="font-size:11px;color:' + C.faint + '">卡關的位置</div><div style="font-size:15px;font-weight:700;color:' + C.rose + '">' + e(t.area) + '</div><div style="font-size:11.5px;color:' + C.faint + ';margin-top:6px;line-height:1.5">' + e(t.areaLine) + '</div></td></tr></table></td>' : '') +
    '</tr></table></td></tr>';
  // 彥彥的話
  r += '<tr><td style="padding:22px ' + PAD + 'px 0"><div style="border-left:3px solid ' + C.pink + ';background:' + C.p2 + ';border-radius:0 16px 16px 0;padding:18px 20px;font-size:15px;line-height:1.85;color:' + C.ink + '">' + br(text['信件開頭']) + '<br><span style="color:' + C.muted + '">' + e(t.tag) + '</span></div></td></tr>';
  // 01 六角圖
  r += sec('01', 'Where you are', '你的身體，現在大概是這樣');
  const tile = (lab, xx, hl, muted) => '<div style="' + (hl ? grad + 'color:#fff;' : 'background:' + C.p2 + ';') + 'border-radius:16px;padding:12px 16px;margin-bottom:10px"><div style="font-size:12px;' + (hl ? 'color:#FFE6EE' : 'color:' + C.faint) + '">' + lab + '</div><div style="font-size:19px;font-weight:900;' + (hl ? 'color:#fff' : 'color:' + C.ink) + '">' + e(xx.name.replace(/型$/, '')) + '<span style="font-family:' + mono + ';font-size:17px;margin-left:8px;' + (hl ? 'color:#fff' : 'color:' + (muted ? C.muted : C.rose)) + '">' + d.pct[xx.code] + '%</span></div></div>';
  r += row('<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr><td valign="middle" width="300"><img src="' + chartUrl + '" width="300" height="270" style="display:block;max-width:100%" alt="代謝指紋六角圖"></td><td valign="middle" style="padding-left:16px">' + tile('主要卡關', sorted[0], true) + tile('第二個要注意的', sorted[1]) + tile('相對穩定的', sorted[sorted.length - 1], false, true) + '</td></tr></table><div style="font-size:12px;color:' + C.faint + ';margin-top:4px">百分比越高，代表這個卡關點越明顯。</div>');
  // 02 原因
  r += sec('02', "Why you're stuck", '為什麼你會卡在這一關');
  const wi = L(t.whyIcons);
  r += row(grid(L(t.why).slice(0, 3).map((w, i) => '<div style="font-size:26px;line-height:1.2;margin-bottom:8px">' + e(wi[i] || '•') + '</div><div style="font-size:14px;line-height:1.7;color:' + C.ink + '">' + e(w) + '</div>'), 3, 6, 'border:1px solid ' + C.line + ';border-radius:18px;padding:16px'));
  // 03 換個吃法
  if (swaps.length) {
    r += sec('03', "Swap, don't skip", '不用忌口，換個吃法就好');
    r += row(grid(swaps.map(x => '<div style="font-size:13px;color:' + C.faint + ';text-decoration:line-through">' + e(x[0]) + '</div><div style="font-size:14px;color:' + C.ink + ';font-weight:700;margin-top:4px"><span style="color:' + C.rose + '">→</span> ' + e(x[1]) + '</div>'), 2, 6, 'background:' + C.p2 + ';border-radius:16px;padding:14px 16px'));
  }
  let k = swaps.length ? 4 : 3;
  // 04 數字
  if (n) {
    r += sec('0' + k++, 'Your numbers', '你每天該吃多少');
    r += row('<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="' + grad + 'border-radius:20px"><tr><td style="padding:20px 22px;color:#fff"><div style="font-size:13px;opacity:.85">每日建議熱量</div><div style="font-family:' + mono + ';font-size:40px;font-weight:700;line-height:1.1">' + n.kcal.toLocaleString() + '</div></td><td align="right" valign="bottom" style="padding:20px 22px;color:#fff;font-size:13px;line-height:1.6">大卡／天<br>' + e(n.note) + '</td></tr></table><div style="height:6px"></div>');
    const nt = (k2, v, s) => '<div style="font-size:11.5px;color:' + C.faint + '">' + k2 + '</div><div style="font-family:' + mono + ';font-size:20px;color:' + C.ink + ';margin:2px 0">' + v + '</div><div style="font-size:11.5px;color:' + C.muted + '">' + s + '</div>';
    r += row(grid([nt('BMI', e(n.bmi), bmiTag(+n.bmi)), nt('體脂率', n.bf ? n.bf + '%' : '—', '建議 ' + (n.sex === '男' ? 25 : 30) + '% 以內'), nt('理想體重', n.ideal[0] + '–' + n.ideal[1], 'kg'), nt('腰圍', n.waist ? n.waist : '—', '建議 ' + (n.sex === '男' ? 90 : 80) + ' cm 以內')], 4, 5, 'background:' + C.p2 + ';border-radius:16px;padding:12px 14px'));
  }
  // 05 早餐
  r += sec('0' + k++, "Tomorrow's breakfast", '明天早餐這樣吃');
  const items = d.bk.items;
  r += row(grid(items.map(x => '<div style="font-size:30px;line-height:1.2">' + iconOf(x.n) + '</div><div style="font-size:13.5px;font-weight:700;color:' + C.ink + ';margin-top:6px;line-height:1.4">' + e(x.n.replace(/（.*）/, '').replace('或荷包蛋', '')) + '</div><div style="font-family:' + mono + ';color:' + C.rose + ';font-size:12.5px;margin-top:2px">' + e(x.a.replace(/（.*）/, '')) + '</div>'), Math.min(items.length, 5), 4, 'border:1px solid ' + C.line + ';border-radius:18px;padding:16px 8px;text-align:center'));
  r += row('<div style="font-size:13px;color:' + C.muted + ';margin:4px 0 6px;line-height:1.7">' + e(d.bk.alt) + '</div>');
  const tip = [t.menuTip, d.bk.tip].filter(Boolean).join(' ');
  r += row(grid([
    '<div style="font-family:' + mono + ';font-size:10.5px;letter-spacing:2px;color:' + C.rose + ';margin-bottom:4px">WHY · 為什麼這樣吃</div><div style="font-size:14px;line-height:1.75;color:' + C.ink + '">' + e(t.bkWhy) + '</div>',
    '<div style="font-family:' + mono + ';font-size:10.5px;letter-spacing:2px;color:' + C.rose + ';margin-bottom:4px">TIP · 彥彥提醒</div><div style="font-size:14px;line-height:1.75;color:' + C.ink + '">' + e(tip) + '</div>'
  ], 2, 6, 'background:' + C.p2 + ';border-radius:18px;padding:14px 16px'));
  // CTA
  r += '<tr><td style="padding:28px ' + PAD + 'px 0"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="' + soft + 'border-radius:24px"><tr>' +
    '<td valign="middle" width="96" style="padding:24px 0 24px 24px"><img src="' + IMG + 'yan-face.jpg" width="96" height="96" style="display:block;border-radius:50%;border:3px solid #fff" alt=""></td>' +
    '<td valign="middle" style="padding:24px"><div style="font-size:19px;font-weight:900;color:' + C.ink + '">' + e(text['結果CTA標題']) + '</div><div style="font-size:14px;color:' + C.muted + ';line-height:1.75;margin:6px 0 14px">' + br(text['信件結尾']) + '</div>' +
    '<a href="' + e(text['IG網址'] || SITE_URL) + '" style="display:inline-block;' + grad + 'color:#fff;text-decoration:none;font-weight:700;font-size:14.5px;padding:12px 24px;border-radius:99px">' + e(text['信件按鈕'] || 'IG 私訊彥彥') + ' →</a></td></tr></table></td></tr>';
  if (!forPdf) r += '<tr><td align="center" style="padding:18px"><a href="' + e(d.link) + '" style="color:' + C.rose + ';font-size:14px">回到我的完整報告 →</a></td></tr>';
  // 彥彥的話＋簽名（簽名小小放在框框右下角）
  r += '<tr><td style="padding:10px ' + PAD + 'px 0"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border:1px solid ' + C.line + ';border-radius:22px"><tr>' +
    '<td valign="top" width="64" style="padding:20px 0 20px 20px"><img src="' + IMG + 'yan-face.jpg" width="52" height="52" style="display:block;border-radius:50%" alt=""></td>' +
    '<td valign="top" style="padding:20px 20px 4px 14px"><div style="font-family:' + mono + ';font-size:11px;letter-spacing:2px;color:' + C.pink + '">FROM 彥彥</div>' +
      '<div style="font-size:13.5px;color:' + C.rose + ';margin-top:6px;line-height:1.7">' + e(text['Slogan英文']) + '<br>' + e(text['Slogan中文']) + '</div>' +
      '<div style="font-size:14px;color:' + C.ink + ';line-height:1.85;margin-top:8px">' + e(text['品牌故事']) + '</div>' +
      '<div style="text-align:right;padding:4px 0 10px"><img src="' + IMG + 'sig.png" width="44" style="display:inline-block;opacity:.85" alt="彥彥"></div></td></tr></table></td></tr>';
  r += '<tr><td style="padding:16px ' + PAD + 'px 18px;font-size:12px;color:' + C.faint + ';line-height:1.8"><b style="color:' + C.ink + ';font-size:13.5px">' + e(text['品牌名稱']) + '</b>　<i style="color:' + C.rose + ';font-family:Georgia,serif">' + e(text['品牌英文']) + '</i><br>' + L(text['資歷']).map(e).join('<br>') + '<br><a href="' + e(text['IG網址']) + '" style="color:' + C.rose + '">' + e(text['IG帳號']) + '</a></td></tr>';
  r += '<tr><td style="padding:0 ' + PAD + 'px 24px;font-size:11px;color:' + C.faint + ';line-height:1.7">' + e(text['免責聲明']) + (forPdf ? '' : '<br>不想再收到信？直接回覆這封信告訴我，我會把你的資料刪除。') + '</td></tr>';

  return '<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width"></head>' +
    '<body style="margin:0;padding:0;background:#F4EEF0"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#F4EEF0"><tr><td align="center" style="padding:' + (forPdf ? '0' : '24px 8px') + '">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:' + W + 'px;background:#ffffff;border-radius:28px;font-family:' + font + ';color:' + C.ink + '">' + r + '</table></td></tr></table></body></html>';
}

// 改版用：把「網站文案」的指定項目和整個「型態」分頁，換成程式裡的最新預設內容（題目不動）
function refreshContent() {
  const ss = SS();
  const keys = ['品牌英文', 'Slogan英文', 'Slogan中文', '品牌故事', '首頁標題', '首頁說明', '開始按鈕', '資歷', '首頁徽章', '首頁前言', '首頁英文', '結果CTA標題', '結果CTA說明', '信件開頭', '信件結尾', '信件按鈕', '信件寄件人', '信件主旨'];
  const t = ss.getSheetByName(SHEET_TEXT);
  const vals = t.getRange(2, 1, Math.max(t.getLastRow() - 1, 1), 1).getValues().map(r => String(r[0]).trim());
  keys.forEach(k => { const i = vals.indexOf(k); if (i >= 0) t.getRange(i + 2, 2).setValue(DEFAULT_TEXT[k]); });
  setup();   // 補上新的項目
  const ty = ss.getSheetByName(SHEET_TYPE);
  ty.clear();
  const rows = [T_HEADERS].concat(DEFAULT_TYPES.map(x => [x.code, x.name, x.en, x.tag, x.why.join('｜'), x.mines.join('｜'), x.food.join('｜'), x.life.join('｜'), x.core, x.menuTip, x.bkWhy, x.area, x.areaIcon, x.areaLine, x.swaps, x.whyIcons]));
  ty.getRange(1, 1, rows.length, T_HEADERS.length).setValues(rows);
  styleHeader(ty, T_HEADERS.length);
  ty.getRange(2, 1, rows.length, T_HEADERS.length).setWrap(true).setVerticalAlignment('top');
  for (let c = 4; c <= 16; c++) ty.setColumnWidth(c, 260);
  Logger.log('已更新');
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
