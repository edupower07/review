/**
 * Manabase デモ用データ投入スクリプト（説明・デモ専用）
 * ──────────────────────────────────────────────────────────────
 * ※ アプリ本体（Code.gs / Index.html）には含めません。デモのときだけ使います。
 *
 * 【使い方】
 *   1. このファイルの内容を Apps Script プロジェクトに新しいスクリプトファイル
 *      （例：SeedDemo.gs）として追加する。
 *   2. 上部の関数選択で「seedDemoData」を選び、「実行」する。
 *      → 生徒名簿・ボード・セクション・投稿・コメント・リアクション・リンク例が入ります。
 *   3. 実行ログ（表示 → ログ）に、ログイン用のパスワードが表示されます。
 *
 * 【ログイン情報（デモ）】
 *   ・生徒：クラス（5年1組/5年2組）→ 名前を選んで、パスワードは全員「1234」
 *   ・先生：パスワード「demo1234」／表示名「福田先生」（2クラスを管理）
 *     （※先生パスワードが未設定のときだけ設定します。設定済みなら変更しません）
 *
 * 【作り直したいとき】
 *   ・「clearAllData」を実行すると全データを消去できます（ヘッダーは残します）。
 *     消去後にもう一度「seedDemoData」を実行すればきれいな状態で入れ直せます。
 *     ※ clearAllData は全データを消すので、本番データが入っている環境では使わないでください。
 *
 * このスクリプトは Code.gs の共通関数（ensureInit_ / getSheet_ / genId_ /
 * sha256_ / newSalt_ / readSheet_ と各 SHEET_* 定数）を利用します。
 */

function seedDemoData() {
  ensureInit_();
  var props = PropertiesService.getScriptProperties();

  // ---- 先生（未設定のときだけパスワードを設定）----
  if (!props.getProperty(PROP_TEACHER_HASH)) {
    var ts = newSalt_();
    props.setProperty(PROP_TEACHER_SALT, ts);
    props.setProperty(PROP_TEACHER_HASH, sha256_(ts + 'demo1234'));
  }
  props.setProperty(PROP_TEACHER_NAME, '福田先生');

  // ---- クラス（2クラス）----
  var cls1 = demoClass_('5年1組', 1);
  var cls2 = demoClass_('5年2組', 2);

  // ---- 生徒名簿（パスワードは全員 1234）----
  var roster1 = [
    [1, '佐藤 あおい'], [2, '鈴木 はると'], [3, '高橋 ゆい'], [4, '田中 そうた'],
    [5, '伊藤 めい'], [6, '渡辺 りく'], [7, '山本 ひなた'], [8, '中村 かなと'],
    [9, '小林 さくら'], [10, '加藤 ゆうと'], [11, '吉田 みお'], [12, '山田 はる']
  ];
  roster1.forEach(function (p) { addDemoStudent_(p[0], p[1], '1234', cls1); });

  var roster2 = [
    [1, '松本 れん'], [2, '井上 ひな'], [3, '木村 そら'], [4, '林 つむぎ'],
    [5, '清水 いつき'], [6, '森 あかり']
  ];
  roster2.forEach(function (p) { addDemoStudent_(p[0], p[1], '1234', cls2); });

  var T = '福田先生';

  // ====================== ボード1：理科（5年1組）======================
  var b1 = demoBoard_('理科', 'ふりこのきまり', 'ふりこのきまり', 1, false, cls1);
  var b1s1 = demoSection_(b1, '気づいたこと', 1, '#2faf6b');
  var b1s2 = demoSection_(b1, 'ぎもん・もっと知りたい', 2, '#e08a3c');
  var b1s3 = demoSection_(b1, 'まとめ', 3, '#3f7fd6');
  var b1s4 = demoSection_(b1, '自分のふりかえり', 4, '#8c5bd0');

  var r;
  r = demoPost_(b1, b1s1, '佐藤 あおい', 'ふりこの長さ',
    'ふりこの長さを長くすると、1往復する時間が長くなった。みじかくすると速くなった。', '#fff7c0',
    { order: 1, pinned: true, minsAgo: 180 });
  demoLike_(r, '鈴木 はると', '❤'); demoLike_(r, '高橋 ゆい', '❤'); demoLike_(r, '田中 そうた', '👍');
  demoComment_(r, '鈴木 はると', 'ぼくのはんも同じ結果になったよ！');
  demoComment_(r, T, 'よく気づきましたね。長さがポイントですね。');

  r = demoPost_(b1, b1s1, '鈴木 はると', 'おもりの重さ',
    'おもりを1個から3個にふやしても、1往復の時間はかわらなかった。意外だった。', '#cfe3ff',
    { order: 2, minsAgo: 160 });
  demoLike_(r, '佐藤 あおい', '😲'); demoLike_(r, '山本 ひなた', '😲');

  r = demoPost_(b1, b1s1, '高橋 ゆい', 'はかり方のくふう',
    '10往復の時間をはかって10でわると、1往復をより正かくにもとめられた。', '#c8f0d0',
    { order: 3, minsAgo: 120 });
  demoLike_(r, '田中 そうた', '👍'); demoComment_(r, '加藤 ゆうと', 'なるほど、それいいね。');

  r = demoPost_(b1, b1s2, '田中 そうた', '',
    'もっと長いふりこだったら何秒になるのか、ためしてみたい。', '#ffe2bf',
    { order: 1, minsAgo: 100 });
  demoLike_(r, '小林 さくら', '🤔');

  r = demoPost_(b1, b1s2, '伊藤 めい', 'ふれはばは？',
    'ふれはばを大きくしたら時間はかわるのかな？次の実験でたしかめたい。', '#ffd1dc',
    { order: 2, minsAgo: 80 });
  demoComment_(r, T, 'いい問いですね。予想も書いてみましょう。');

  r = demoPost_(b1, b1s3, '山本 ひなた', '今日のまとめ',
    'ふりこが1往復する時間は「ふりこの長さ」でかわる。おもりの重さやふれはばはかんけいなかった。', '#fff7c0',
    { order: 1, minsAgo: 30 });
  demoLike_(r, '佐藤 あおい', '❤'); demoLike_(r, '鈴木 はると', '❤'); demoLike_(r, '高橋 ゆい', '👍');
  demoLike_(r, '田中 そうた', '❤');

  r = demoPost_(b1, b1s3, '渡辺 りく', 'グループのけつろん',
    '長さが同じなら、だれがはかっても同じくらいの時間になった。だから長さで決まると言える。', '#c8f0d0',
    { order: 2, minsAgo: 25 });
  demoLike_(r, '中村 かなと', '👍');

  r = demoPost_(b1, b1s4, '中村 かなと', '',
    '予想とちがって重さがかんけいないとわかって、実験してたしかめる大切さがわかった。', '#e8d6ff',
    { order: 1, minsAgo: 20 });
  demoLike_(r, '山本 ひなた', '❤'); demoComment_(r, T, 'たしかめる姿勢、すばらしいです。');

  r = demoPost_(b1, b1s4, '小林 さくら', '',
    'はかり方をくふうすると正かくになることがわかった。次もていねいにはかりたい。', '#ffe2bf',
    { order: 2, minsAgo: 15 });
  demoLike_(r, '高橋 ゆい', '👍'); demoLike_(r, '伊藤 めい', '❤');

  // ====================== ボード2：国語（5年1組）======================
  var b2 = demoBoard_('国語', 'ごんぎつね', 'ごんぎつね', 3, false, cls1);
  var b2s1 = demoSection_(b2, '心にのこった場面', 1, '#d24b8c');
  var b2s2 = demoSection_(b2, '登場人物の気もち', 2, '#8c5bd0');
  var b2s3 = demoSection_(b2, 'すきな一文', 3, '#2faf6b');
  var b2s4 = demoSection_(b2, '感想・伝えたいこと', 4, '#e08a3c');

  r = demoPost_(b2, b2s1, '小林 さくら', '',
    'ごんがつぐないをするところが心にのこった。いたずらをこうかいしていたんだと思う。', '#e8d6ff',
    { order: 1, pinned: true, minsAgo: 1400 });
  demoLike_(r, '吉田 みお', '❤'); demoLike_(r, '山田 はる', '❤');
  demoComment_(r, '吉田 みお', 'わたしもそこが心にのこりました。');

  r = demoPost_(b2, b2s1, '加藤 ゆうと', '',
    '兵十がごんをうってしまう場面が悲しかった。気づくのがおそかった。', '#cfe3ff',
    { order: 2, minsAgo: 1380 });
  demoLike_(r, '小林 さくら', '😢'); demoLike_(r, '中村 かなと', '😢');

  r = demoPost_(b2, b2s2, '吉田 みお', 'ごんの気もち',
    'ひとりぼっちのごんは、兵十とつながりたかったんじゃないかな。', '#ffd1dc',
    { order: 1, minsAgo: 1300 });
  demoComment_(r, T, '「ひとりぼっち」に注目したのがいいですね。');

  r = demoPost_(b2, b2s2, '山田 はる', '兵十の気もち',
    '兵十は最後にごんの気もちに気づいて、後かいしたと思う。', '#cfe3ff',
    { order: 2, minsAgo: 1280 });
  demoLike_(r, '加藤 ゆうと', '😢');

  r = demoPost_(b2, b2s3, '佐藤 あおい', '',
    '「ごん、おまいだったのか。」という一文がいちばん心にのこった。', '#fff7c0',
    { order: 1, pinned: true, minsAgo: 1200 });
  demoLike_(r, '小林 さくら', '❤'); demoLike_(r, '吉田 みお', '❤'); demoLike_(r, '山田 はる', '❤');

  r = demoPost_(b2, b2s4, '高橋 ゆい', '',
    '気もちはきちんと言葉で伝えないと、すれちがってしまうと感じた。', '#c8f0d0',
    { order: 1, minsAgo: 1100 });
  demoLike_(r, '中村 かなと', '👍'); demoComment_(r, T, '物語から大切なことを受け取りましたね。');

  // ====================== ボード3：社会（5年1組・リンク例あり）======================
  var b3 = demoBoard_('社会', 'だれもがくらしやすいまち', 'だれもがくらしやすいまち', 5, false, cls1);
  var b3s1 = demoSection_(b3, '見つけたくふう', 1, '#2796a8');
  var b3s2 = demoSection_(b3, '調べてわかったこと', 2, '#c8862a');
  var b3s3 = demoSection_(b3, 'みんなに伝えたいこと', 3, '#3f7fd6');
  var b3s4 = demoSection_(b3, 'もっと調べたいこと', 4, '#e8554e');

  r = demoPost_(b3, b3s1, '中村 かなと', 'スロープを見つけた',
    '駅にスロープと点字ブロックがあった。だれでも使えるようにくふうされていた。', '#c8f0d0',
    { order: 1, minsAgo: 2000 });
  demoLike_(r, '佐藤 あおい', '👍'); demoLike_(r, '山田 はる', '👍');

  // リンク（ユニバーサルデザインの記事）。表示名つき・権限不要のリンクカード。
  r = demoPost_(b3, b3s2, '山田 はる', 'ユニバーサルデザインの記事',
    'ユニバーサルデザインについて調べたページです。みんなにも読んでほしい。', '#cfe3ff',
    {
      order: 1, minsAgo: 1900,
      link: {
        url: 'https://whill.inc/jp/column/14_universaldesign',
        title: '知っておきたい あなたの身近にユニバーサルデザイン',
        site: 'whill.inc',
        image: '',
        favicon: 'https://www.google.com/s2/favicons?sz=128&domain=whill.inc'
      }
    });
  demoLike_(r, '中村 かなと', '❤'); demoLike_(r, '小林 さくら', '👍');
  demoComment_(r, T, 'よく見つけましたね。どんなくふうがありましたか？');

  r = demoPost_(b3, b3s3, '伊藤 めい', '',
    'だれもが使えるくふうは、こまっている人だけでなく、みんなが助かるとわかった。', '#cfe3ff',
    { order: 1, minsAgo: 1700 });
  demoLike_(r, '田中 そうた', '❤'); demoLike_(r, '渡辺 りく', '👍');

  r = demoPost_(b3, b3s4, '渡辺 りく', '',
    '自分のまちには、ほかにどんなユニバーサルデザインがあるか調べてみたい。', '#ffe2bf',
    { order: 1, minsAgo: 1600 });
  demoLike_(r, '山田 はる', '🤔'); demoComment_(r, '加藤 ゆうと', 'いっしょに調べたい！');

  // ====================== ボード4：非表示（アーカイブ）の例（5年1組）======================
  // 児童のボード一覧には出ません。先生画面では「非表示」として確認・再表示できます。
  demoBoard_('算数・数学', 'いろいろな単位（おわった単元）', 'いろいろな単位', 20, true, cls1);

  // ====================== ボード5：5年2組 ======================
  var c2b1 = demoBoard_('理科', 'こん虫のかんさつ', 'こん虫のかんさつ', 2, false, cls2);
  var c2s1 = demoSection_(c2b1, '見つけたこと', 1, '#2faf6b');
  var c2s2 = demoSection_(c2b1, 'ぎもん', 2, '#e08a3c');
  r = demoPost_(c2b1, c2s1, '松本 れん', '', 'モンシロチョウのよう虫がキャベツの葉を食べていた。', '#fff7c0', { order: 1, minsAgo: 200 });
  demoLike_(r, '井上 ひな', '❤'); demoLike_(r, '森 あかり', '👍');
  r = demoPost_(c2b1, c2s1, '井上 ひな', '', 'アリは行列を作って同じ道を通っていた。', '#cfe3ff', { order: 2, minsAgo: 150 });
  demoComment_(r, T, 'よく観察できましたね。');
  r = demoPost_(c2b1, c2s2, '木村 そら', '', 'なぜチョウはひらひら飛ぶのかな？', '#ffe2bf', { order: 1, minsAgo: 100 });
  demoLike_(r, '林 つむぎ', '🤔');

  Logger.log('✅ デモデータを投入しました。\n'
    + '・クラス：5年1組 / 5年2組\n'
    + '・生徒ログイン：クラスを選び 名前を選んで パスワード「1234」\n'
    + '・先生ログイン：パスワード「demo1234」（表示名：福田先生）\n'
    + '・5年1組：理科／国語／社会（リンク例）＋非表示の例（算数）／5年2組：理科');
  return 'done';
}

function demoClass_(name, order) {
  var existing = readSheet_(SHEET_CLASSES).filter(function (c) { return c.name === name; })[0];
  if (existing) return existing.classId;
  var id = genId_('cls');
  getSheet_(SHEET_CLASSES).appendRow([id, name, order || 1, new Date()]);
  return id;
}

// ---- 低レベルの投入ヘルパー（列順は Code.gs の SHEET_DEFS と一致させる）----

function addDemoStudent_(num, name, pw, classId) {
  var exists = readSheet_(SHEET_STUDENTS).some(function (s) {
    return s.name === name && String(s.classId || '') === String(classId || '');
  });
  if (exists) return;
  var salt = newSalt_();
  // 列順：number, name, salt, passwordHash, createdAt, classId
  getSheet_(SHEET_STUDENTS).appendRow([num, name, salt, sha256_(salt + pw), new Date(), classId || '']);
}

function demoBoard_(subject, unit, title, daysAgo, archived, classId) {
  var id = genId_('b');
  var d = new Date(Date.now() - (daysAgo || 0) * 86400000);
  var ymd = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
  // 列順：boardId, subject, unit, date, title, createdAt, archived, classId
  getSheet_(SHEET_BOARDS).appendRow([id, subject, unit, "'" + ymd, title || unit, d, !!archived, classId || '']);
  return id;
}

function demoSection_(boardId, name, order, color) {
  var id = genId_('s');
  // 列順：sectionId, boardId, name, sortOrder, createdAt, color
  getSheet_(SHEET_SECTIONS).appendRow([id, boardId, name, order || 1, new Date(), color || '']);
  return id;
}

function demoPost_(boardId, sectionId, author, title, text, color, opts) {
  opts = opts || {};
  var id = genId_('r');
  var at = opts.at || new Date(Date.now() - (opts.minsAgo || 0) * 60000);
  // 列順：reflectionId, boardId, studentName, text, photoUrl, photoFileId, color,
  //       sortOrder, createdAt, sectionId, mediaType, updatedAt, pinned, title, link
  getSheet_(SHEET_REFLECTIONS).appendRow([
    id, boardId, author, text || '', opts.photoUrl || '', opts.photoFileId || '',
    color || '#fff7c0', opts.order || 1, at, sectionId, opts.mediaType || '', at,
    !!opts.pinned, title || '', opts.link ? JSON.stringify(opts.link) : ''
  ]);
  return id;
}

function demoComment_(reflectionId, author, text) {
  getSheet_(SHEET_COMMENTS).appendRow([genId_('c'), reflectionId, author, text, new Date()]);
}

function demoLike_(reflectionId, studentName, type) {
  // 列順：reflectionId, studentName, createdAt, type
  getSheet_(SHEET_LIKES).appendRow([reflectionId, studentName, new Date(), type || '❤']);
}

/**
 * 全データを消去します（各シートのヘッダー行は残します）。
 * ※ デモのやり直し用。本番データのある環境では使わないでください。
 *   先生・生徒のパスワード設定（スクリプトプロパティ）は消しません。
 */
function clearAllData() {
  ensureInit_();
  [SHEET_CLASSES, SHEET_STUDENTS, SHEET_BOARDS, SHEET_SECTIONS, SHEET_REFLECTIONS, SHEET_COMMENTS, SHEET_LIKES]
    .forEach(function (name) {
      var sh = getSpreadsheet_().getSheetByName(name);
      if (!sh) return;
      var last = sh.getLastRow();
      if (last > 1) sh.deleteRows(2, last - 1);
    });
  Logger.log('🧹 全データを消去しました（ヘッダーは保持）。seedDemoData で入れ直せます。');
  return 'cleared';
}
