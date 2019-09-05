// The MIT License (MIT)
//
// Copyright (c) SoftBank Corp.
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.

/**
 * スプレッドシート連携
 */

// ----------------------------------------------------------------------------
// グローバル
/* globals NaturalLanguageClassifierV1 */
/* globals GASLIB_Mail_send */
/* globals GASLIB_Dialog_open */
/* globals GASLIB_Text_normalize */
/* globals GASLIB_Trigger_set */
/* globals GASLIB_Trigger_del */
/* globals GASLIB_SheetLog_write */

/* globals RUNTIME_CONFIG */
/* globals RUNTIME_OPTION */
/* globals RUNTIME_ACTIVE */

/* globals CONF_INDEX */
/* globals CONFIG_SET */
/* globals NB_CLFS */

/* globals NLCLIB_MAX_TRAIN_RECORDS */
/* globals NLCLIB_MAX_TRAIN_STRINGS */

/**
 * 分類器名のプリフィクス
 * @type {String}
 */
var CLFNAME_PREFIX = "CLF";

/**
 * 分類器名のセパレータ
 * @type {String}
 */
var CLF_SEP = "#__#";

// ----------------------------------------------------------------------------

/**
 * 通知オプション
 * @type {Object}
 * @property {String} ON オン
 * @property {String} OFF オフ
 */
var NOTIF_OPT = {
  ON: "On",
  OFF: "Off"
};

/**
 * 通知ルールレコードのフィールドインデックス
 * @typedef {Object} NOTIF_INDEX
 * @property {Integer} result1 分類結果1
 * @property {Integer} result2 分類結果2
 * @property {Integer} result3 分類結果3
 * @property {Integer} from  送信元メールアドレス
 * @property {Integer} to 送信先メールアドレス
 * @property {Integer} cc CCメールアドレス
 * @property {Integer} bcc BCCメールアドレス
 * @property {Integer} subject 件名
 * @property {Integer} body 本文
 */
var NOTIF_INDEX = {
  result1: 0,
  result2: 1,
  result3: 2,
  from: 3,
  to: 4,
  cc: 5,
  bcc: 6,
  subject: 7,
  body: 8
};
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * 分類器削除ログ出力
 * @param       {Object} p_params params
 * @param       {Object} p_params.log_set    ログ設定
 * @param       {Object} p_params.del_set    削除設定
 * @param       {Object} p_params.del_result 削除結果
 */
function NLCAPP_log_delete(p_params) {
  NLCAPP_log_debug({ record: ["NLCAPP_log_delete", "START"] });

  var settings = {
    sheet: p_params.log_set.sheet,
    start_row: p_params.log_set.start_row,
    start_col: p_params.log_set.start_col
  };
  var record = [];
  var colors = [
    "black",
    "black",
    "black",
    "black",
    "black",
    "black",
    "black",
    "black",
    "black",
    "black"
  ];
  var params = {
    settings: settings,
    record: record,
    colors: colors
  };

  if (p_params.del_result.status === 200) {
    record = [
      "削除",
      Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss"),
      "成功",
      p_params.del_set.clf_no,
      "",
      "",
      "",
      "",
      p_params.del_set.clf_id,
      p_params.del_result.status
    ];
    params.record = record;
    GASLIB_SheetLog_write(params);
  } else {
    if (p_params.del_result["nlc"]) {
      record = [
        "削除",
        Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss"),
        "失敗",
        p_params.del_set.clf_no,
        "",
        "",
        "",
        "",
        p_params.del_set.clf_id,
        p_params.del_result.status,
        p_params.del_result.nlc.body.status_description
      ];
    } else {
      record = [
        "削除",
        Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss"),
        p_params.del_result.status,
        p_params.del_set.clf_no,
        "",
        "",
        "",
        "",
        p_params.del_set.clf_id,
        p_params.del_result.code
      ];
    }
    colors = [
      "red",
      "red",
      "red",
      "red",
      "red",
      "red",
      "red",
      "red",
      "red",
      "red"
    ];
    params.record = record;
    params.colors = colors;

    GASLIB_SheetLog_write(params);
  }
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * 学習結果ログ出力
 * @param       {Object} p_params   ログ出力設定
 * @param       {Object} p_params.log_set   ログ出力設定
 * @param       {Object} p_params.log_set.sheet   ログ出力設定
 * @param       {Object} p_params.log_set.start_row   ログ出力設定
 * @param       {Object} p_params.log_set.start_col   ログ出力設定
 * @param       {Object} p_params.train_set 学習設定
 * @param       {Object} p_params.train_result      学習結果
 * @throws      {Error} 分類器のエラーです。ログを確認してください
 */
function NLCAPP_log_train(p_params) {
  NLCAPP_log_debug({ record: ["NLCAPP_log_train", "START"] });
  /*
  処理,実行日時,ステータス,分類器,シート名,件数,テキスト列,インテント列
  分類結果列,Classifier_ID,status,code,description,created,version
  */

  var settings = {
    sheet: p_params.log_set.sheet,
    start_row: p_params.log_set.start_row,
    start_col: p_params.log_set.start_col
  };
  var record = [];
  var colors = [
    "black",
    "black",
    "black",
    "black",
    "black",
    "black",
    "black",
    "black",
    "black",
    "black",
    "black",
    "black",
    "black"
  ];
  var params = {
    settings: settings,
    record: record,
    colors: colors
  };

  if (p_params.train_result.status === 200) {
    record = [
      "学習",
      Utilities.formatDate(
        new Date(p_params.train_result.nlc.from),
        "JST",
        "yyyy/MM/dd HH:mm:ss"
      ),
      p_params.train_result.nlc.body.status,
      p_params.train_set.clf_no,
      p_params.train_set.ws_name,
      p_params.train_result.rows, //件数
      p_params.train_set.text_col,
      p_params.train_set.class_col,
      p_params.train_result.nlc.body.classifier_id,
      p_params.train_result.status,
      p_params.train_result.nlc.body.status_description,
      p_params.train_result.nlc.body.created,
      p_params.train_result.version
    ];

    if (p_params.train_result.rows > NLCLIB_MAX_TRAIN_RECORDS) {
      record[5] =
        NLCLIB_MAX_TRAIN_RECORDS +
        "(初めの" +
        (p_params.train_result.rows - NLCLIB_MAX_TRAIN_RECORDS) +
        "件は除外)";
      colors[5] = 1;
    }
    params.record = record;
    GASLIB_SheetLog_write(params);
  } else if (p_params.train_result.status === 2000) {
    record = [
      "学習",
      Utilities.formatDate(
        new Date(p_params.train_result.nlc.from),
        "JST",
        "yyyy/MM/dd HH:mm:ss"
      ),
      p_params.train_result.nlc.body.status,
      p_params.train_set.clf_no,
      "",
      "",
      "",
      "",
      p_params.train_result.nlc.body.classifier_id,
      p_params.train_result.nlc.status,
      p_params.train_result.nlc.body.status_description,
      p_params.train_result.nlc.body.created,
      p_params.train_result.version
    ];
    params.record = record;
    GASLIB_SheetLog_write(params);
  } else if (p_params.train_result.status === 0) {
    record = [
      "学習",
      Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss"),
      p_params.train_result.description,
      p_params.train_set.clf_no,
      p_params.train_set.ws_name,
      0,
      p_params.train_set.text_col,
      p_params.train_set.class_col,
      "",
      "",
      "",
      "",
      ""
    ];
    params.record = record;
    GASLIB_SheetLog_write(params);
  } else if (p_params.train_result.status === 999) {
    record = [
      "学習",
      Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss"),
      p_params.train_result.description,
      p_params.train_set.clf_no,
      p_params.train_set.ws_name,
      "N/A",
      p_params.train_set.text_col,
      p_params.train_set.class_col,
      "",
      p_params.train_result.status,
      p_params.train_result.error_desc ? p_params.train_result.error_desc : ""
    ];
    colors = [
      "red",
      "red",
      "red",
      "red",
      "red",
      "red",
      "red",
      "red",
      "red",
      "red",
      "red",
      "red",
      "red"
    ];
    params.record = record;
    params.colors = colors;
    GASLIB_SheetLog_write(params);
  } else {
    // エラー
    if (p_params.train_result["nlc"]) {
      record = [
        "学習",
        Utilities.formatDate(
          new Date(p_params.train_result.nlc.from),
          "JST",
          "yyyy/MM/dd HH:mm:ss"
        ),
        p_params.train_result.nlc.body.error,
        p_params.train_set.clf_no,
        p_params.train_set.ws_name,
        p_params.train_result.rows,
        p_params.train_set.text_col,
        p_params.train_set.class_col,
        "",
        p_params.train_result.status,
        p_params.train_result.nlc.body.description
      ];
    } else {
      record = [
        "学習",
        Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss"),
        p_params.train_result.description,
        p_params.train_set.clf_no,
        p_params.train_set.ws_name,
        "rows" in p_params.train_result ? p_params.train_result.rows : "N/A",
        p_params.train_set.text_col,
        p_params.train_set.class_col,
        "",
        p_params.train_result.status,
        p_params.train_result.error_desc ? p_params.train_result.error_desc : ""
      ];
    }
    params.colors = [
      "red",
      "red",
      "red",
      "red",
      "red",
      "red",
      "red",
      "red",
      "red",
      "red",
      "red"
    ];
    params.record = record;

    GASLIB_SheetLog_write(params);
    throw new Error("分類器のエラーです。ログを確認してください");
  }
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * 分類結果ログ出力
 * @param       {Object} p_params     ログ出力設定
 * @param       {Object} p_params.log_set     ログ出力設定
 * @param       {Object} p_params.test_set    テスト設定
 * @param       {Object} p_params.test_result テスト結果
 * @throws      {Error} 分類器のエラーです。ログを確認してください
 */
function NLCAPP_log_classify(p_params) {
  NLCAPP_log_debug({ record: ["NLCAPP_log_classify", "START"] });

  var settings = {
    sheet: p_params.log_set.sheet,
    start_row: p_params.log_set.start_row,
    start_col: p_params.log_set.start_col
  };
  var record = [];
  var colors = [
    "black",
    "black",
    "black",
    "black",
    "black",
    "black",
    "black",
    "black",
    "black",
    "black"
  ];
  var params = {
    settings: settings,
    record: record,
    colors: colors
  };

  // 成功
  if (p_params.test_result.status === 200) {
    record = [
      "分類",
      Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss"),
      "成功",
      p_params.test_set.clf_no,
      p_params.test_set.ws_name,
      p_params.test_result.rows,
      p_params.test_set.text_col,
      p_params.test_set.result_col,
      p_params.test_result.nlc.body.classifier_id,
      p_params.test_result.status
    ];
    params.record = record;
    GASLIB_SheetLog_write(params);
  } else if (p_params.test_result.status === 0) {
    record = [
      "分類",
      Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss"),
      "対象なし",
      p_params.test_set.clf_no,
      p_params.test_set.ws_name,
      p_params.test_result.rows,
      p_params.test_set.text_col,
      p_params.test_set.result_col,
      "",
      ""
    ];
    params.record = record;
    GASLIB_SheetLog_write(params);
  } else if (p_params.test_result.status === 800) {
    record = [
      "分類",
      Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss"),
      p_params.test_result.description,
      p_params.test_set.clf_no,
      p_params.test_set.ws_name,
      "N/A",
      p_params.test_set.text_col,
      p_params.test_set.result_col,
      p_params.test_result.clf_id,
      ""
    ];
    params.record = record;
    GASLIB_SheetLog_write(params);
  } else if (p_params.test_result.status === 900) {
    record = [
      "分類",
      Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss"),
      p_params.test_result.description,
      p_params.test_set.clf_no,
      p_params.test_set.ws_name,
      "N/A",
      p_params.test_set.text_col,
      p_params.test_set.result_col,
      p_params.test_result.clf_id,
      ""
    ];
    params.record = record;
    GASLIB_SheetLog_write(params);
  } else {
    if (p_params.test_result["nlc"]) {
      record = [
        "分類",
        Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss"),
        p_params.test_result.description,
        p_params.test_set.clf_no,
        p_params.test_set.ws_name,
        "N/A",
        p_params.test_set.text_col,
        p_params.test_set.result_col,
        "",
        ""
      ];
    } else {
      record = [
        "分類",
        Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss"),
        p_params.test_result.description,
        p_params.test_set.clf_no,
        p_params.test_set.ws_name,
        "N/A",
        p_params.test_set.text_col,
        p_params.test_set.result_col,
        p_params.test_result.clf_id,
        p_params.test_result.status
      ];
    }
    colors = [
      "red",
      "red",
      "red",
      "red",
      "red",
      "red",
      "red",
      "red",
      "red",
      "red"
    ];
    params.record = record;
    params.colors = colors;
    GASLIB_SheetLog_write(params);

    if (
      typeof p_params.test_result.throw_exception !== "undefined" &&
      p_params.test_result.throw_exception === false
    ) {
      return;
    }
    throw new Error("分類器のエラーです。ログを確認してください");
  }
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * デバッグログ出力
 * @param       {Object} p_params パラメータ
 */
function NLCAPP_log_debug(p_params) {
  if (!RUNTIME_OPTION.LOG_DEBUG) return;

  var conf = NLCAPP_load_config(CONFIG_SET);

  var log_sheet = conf.self_ss.getSheetByName(conf.sheet_conf.log_ws);
  if (log_sheet === null) {
    log_sheet = conf.self_ss.insertSheet(conf.sheet_conf.log_ws);
  }

  var settings = {
    sheet: log_sheet,
    start_row: CONFIG_SET.log_start_row,
    start_col: CONFIG_SET.log_start_col
  };
  var record = [
    "DEBUG",
    Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss"),
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    ""
  ];

  var pos = 2;
  var cnt = p_params.record.length < 8 ? p_params.record.length : 8;
  for (var i = 0; i < cnt; i += 1) {
    record[pos] = p_params.record[i];
    pos += 1;
  }
  var colors = [
    "blue",
    "blue",
    "blue",
    "blue",
    "blue",
    "blue",
    "blue",
    "blue",
    "blue",
    "blue"
  ];
  var params = {
    settings: settings,
    record: record,
    colors: colors
  };

  GASLIB_SheetLog_write(params);
  Logger.log(
    Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss.SSS")
  );
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * @typedef {Object} SheetSet シート基本情報
 * @property {String} ss_id スプレッドシートID
 * @property {String} ws_name シート名
 * @property {Integer} start_col 開始列
 * @property {Integer} start_row 開始行
 */

/**
 * @typedef {Object} Creds 資格情報
 * @property {String} username ユーザー名
 * @property {String} password パスワード
 * @property {String} url      エンドポイント
 */

/**
 * 資格情報の取得
 * <p>利用するNLCインスタンスの資格情報をスクリプトプロパティから取得する</p>
 * @return {Creds} 資格情報
 * @throws {Error}  資格情報が不明です
 */
function NLCAPP_load_creds() {
  // eslint-disable-line no-unused-vars

  NLCAPP_log_debug({ record: ["NLCAPP_load_creds", "START"] });

  var scriptProps = PropertiesService.getScriptProperties();

  var creds = {};
  creds["url"] = scriptProps.getProperty("CREDS_URL");
  creds["username"] = scriptProps.getProperty("CREDS_USERNAME");
  creds["password"] = scriptProps.getProperty("CREDS_PASSWORD");

  if (
    creds.url === null ||
    creds.username === null ||
    creds.password === null
  ) {
    throw new Error("資格情報が不明です");
  }

  return creds;
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * @typedef {Object} ConfigMeta 設定メタデータ
 * @property {String} ss_id スプレッドシートID
 * @property {String} ws_name 設定シート名
 * @property {Integer} st_start_row 定義開始行
 * @property {Integer} st_start_col 定義開始列
 */
/**
 * @typedef {Object} SheetConf データシート設定
 * @property {String} ws_name データシート名
 * @property {Integer} start_row 定義開始行
 * @property {Integer} start_col 定義開始列
 * @property {Integer[]} intent_col インテント列1to3
 * @property {Integer[]} result_col 分類結果列1to3
 * @property {Integer[]} resconf_col 確信度列1to3
 * @property {Integer[]} restime_col 分類日時列1to3
 * @property {String} log_ws ログシート名
 */
/**
 * @typedef {Object} NotifConf 通知設定
 * @property {String} notif_opt 通知オプション{On,Off}
 * @property {String} notif_ws 設定シート名
 */
/**
 * @typedef {Object} Config 設定情報
 * @property {SheetConf} sheet_conf データシート設定
 * @property {NotifConf} notif_conf 通知設定
 */
/**
 * 設定情報の取得
 * <p>メタデータを元にユーザーの設定情報を取得する</p>
 * @param  {ConfigMeta} config_set コンフィグメタデータ
 * @return {Config} コンフィグ
 * @throws {Error}  設定シートが不明です
 * @throws {Error}  設定シートに問題があります
 * @throws {Error}  学習・分類対象列が不正です'
 */
function NLCAPP_load_config(config_set) {
  //NLCAPP_log_debug({ record: ["NLCAPP_load_config", "START"] });

  var self_ss = SpreadsheetApp.getActiveSpreadsheet();
  var ss_id = self_ss.getId();

  var sheet = self_ss.getSheetByName(config_set.ws_name);
  if (sheet === null) {
    throw new Error("設定シートが不明です");
  }

  var nb_conf = Object.keys(CONF_INDEX).length;

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (
    lastRow < config_set.st_start_row + nb_conf - 1 ||
    lastCol < config_set.st_start_col
  ) {
    throw new Error("設定シートに問題があります");
  }

  // 設定シートの内容
  var records = sheet
    .getRange(config_set.st_start_row, config_set.st_start_col, nb_conf, 1)
    .getValues();

  /*
  4. 学習・分類対象列
  9. 「分類器2:手入力」列
  10. 「分類器2:Watson」列
  11. 「分類器2:確信度」列
  12. 「分類器2:処理日時」列
  13. 「分類器3:手入力」列
  14. 「分類器3:Watson」列
  15. 「分類器3:確信度」列
  16. 「分類器3:処理日時」列
  */

  var i = 0;
  var sheet_conf = {
    sheet: sheet,
    ws_name: records[CONF_INDEX.ws_name][i], //1. データシート名
    start_row: parseInt(records[CONF_INDEX.start_row][i], 10), //2. 開始列
    start_col: parseInt(records[CONF_INDEX.start_col][i], 10), //3. 開始行
    intent_col: [
      parseInt(records[CONF_INDEX.intent1_col][i], 10), //5. 「分類器1:手入力」列
      parseInt(records[CONF_INDEX.intent2_col][i], 10),
      parseInt(records[CONF_INDEX.intent3_col][i], 10)
    ],
    result_col: [
      parseInt(records[CONF_INDEX.result1_col][i], 10), //6. 「分類器1:Watson」列
      parseInt(records[CONF_INDEX.result2_col][i], 10),
      parseInt(records[CONF_INDEX.result3_col][i], 10)
    ],
    resconf_col: [
      parseInt(records[CONF_INDEX.resconf1_col][i], 10), //7. 「分類器1:確信度」列
      parseInt(records[CONF_INDEX.resconf2_col][i], 10),
      parseInt(records[CONF_INDEX.resconf3_col][i], 10)
    ],
    restime_col: [
      parseInt(records[CONF_INDEX.restime1_col][i], 10), //8. 「分類器1:処理日時」列
      parseInt(records[CONF_INDEX.restime2_col][i], 10),
      parseInt(records[CONF_INDEX.restime3_col][i], 10)
    ],
    log_ws: records[CONF_INDEX.log_ws][i] //17. ログシート名
  };

  if (CONF_INDEX.text_col) {
    sheet_conf.text_col = parseInt(records[CONF_INDEX.text_col][i], 10);
    if (parseInt(sheet_conf.text_col, 10) <= 0) {
      throw new Error("学習・分類対象列が不正です");
    }
  }

  var notif_conf = {};
  if (CONF_INDEX.notif_opt) {
    notif_conf["option"] = records[CONF_INDEX.notif_opt][0]; //18. メール通知
  }
  if (CONF_INDEX.notif_ws) {
    notif_conf["ws_name"] = records[CONF_INDEX.notif_ws][0]; //19. メール通知設定シート名
  }

  RUNTIME_CONFIG.sheet_conf = sheet_conf;
  RUNTIME_CONFIG.notif_conf = notif_conf;

  return {
    self_ss: self_ss,
    ss_id: ss_id,
    sheet_conf: sheet_conf,
    notif_conf: notif_conf
  };
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * NLCインスタンスオブジェクトの生成
 * @param       {String} p_username 資格情報のusername
 * @param       {String} p_password 資格情報のpassword
 * @param       {String} p_url      資格情報のURL
 * @return      {NaturalLanguageClassifierV1} インスタンスオブジェクト
 */
function NLCAPP_create_instance(p_username, p_password, p_url) {
  NLCAPP_log_debug({ record: ["NLCAPP_create_instance", "START"] });

  var params = {
    url: p_url
  };

  if (p_username === "apikey") {
    params.iam_apikey = p_password;
  } else {
    params.username = p_username;
    params.password = p_password;
  }

  var nlc = new NaturalLanguageClassifierV1(params);

  NLCAPP_log_debug({ record: ["NLCAPP_create_instance", "END"] });
  return nlc;
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * @typedef {Object} ExpandResult 展開結果
 * @property {String} code 結果コード
 * @property {String} text 結果テキスト
 */
/**
 * 通知ルール用の埋め込みタグを展開する
 * <p>対象テキスト中に埋め込まれたタグをインデックスに該当する対象フィールドに置換する</p>
 * <p>埋め込みタグの形式 [[#インデックス]] ※インデックスは1以上の列として有効な整数</p>
 * @param {String} target 変換対象テキスト
 * @param {String[]} fields 展開対象フィールド
 * @return {ExpandResult} 展開結果
 */
function NLCAPP_expand_tags(target, fields) {
  NLCAPP_log_debug({ record: ["NLCAPP_expand_tags", "START"] });

  var xbody = target;
  var buf = "";

  for (var idx = 0; idx < fields.length; idx += 1) {
    buf = xbody.replace(
      new RegExp("\\[\\[#" + String(idx + 1) + "\\]\\]", "g"),
      fields[idx]
    );
    xbody = buf;
  }

  xbody.match(new RegExp("\\[\\[.+\\]\\]", "g"));

  return {
    code: "OK",
    text: xbody
  };
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * @typedef {Object} Mail メール設定
 * @property {String} from 送信元メールアドレス
 * @property {String} to 送信先メールアドレス
 * @property {String} cc ccメールアドレス
 * @property {String} bcc bccメールアドレス
 * @property {String} subject 件名
 * @property {String} body 本文
 */
/**
 * @typedef {Object} Rules 通知条件
 * @property {String[]} res_int 分類結果1to3
 * @property {Mail} mail メール設定
 */
/**
 * 通知条件の取得
 * <p>
 * </p>
 * @param       {SheetSet} config_set 設定情報
 * @return      {Rules[]} 通知条件
 */
function NLCAPP_load_notif_rules(config_set) {
  NLCAPP_log_debug({ record: ["NLCAPP_load_notif_rules", "START"] });

  var sheet = config_set.self_ss.getSheetByName(config_set.ws_name);
  if (sheet === null) {
    return [];
  }

  var nb_conf = Object.keys(NOTIF_INDEX).length;

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastCol < config_set.start_col + nb_conf - 1) {
    return [];
  }

  if (lastRow < config_set.start_row) {
    return [];
  }

  var records = sheet
    .getRange(
      config_set.start_row,
      config_set.start_col,
      lastRow - config_set.start_row + 1,
      nb_conf
    )
    .setNumberFormat("@")
    .getValues();

  var rules = [];
  for (var i = 0; i < records.length; i += 1) {
    rules.push({
      res_int: [
        String(records[i][NOTIF_INDEX.result1]),
        String(records[i][NOTIF_INDEX.result2]),
        String(records[i][NOTIF_INDEX.result3])
      ],
      mail: {
        from: records[i][NOTIF_INDEX.from],
        to: records[i][NOTIF_INDEX.to],
        cc: records[i][NOTIF_INDEX.cc],
        bcc: records[i][NOTIF_INDEX.bcc],
        subject: records[i][NOTIF_INDEX.subject],
        body: records[i][NOTIF_INDEX.body]
      }
    });
  }
  return rules;
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * @typedef {Object} ClassifierInfoPayload 分類器情報
 * @property {String} classifier_id ID
 * @property {String} name 名称
 * @property {String} created 作成日
 */
/**
 * @typedef {Object} ClfVers 分類器バージョン一覧
 * @property {Integer} count 件数
 * @property {Integer} min_ver 最小バージョン
 * @property {Integer} max_ver 最大バージョン
 * @property {ClassifierInfoPayload[]} clfs 分類器情報
 */
/**
 * 分類器のバージョン一覧を取得
 * <p>分類器一覧から分類器名(ex. CLF1)に該当するバージョン一覧を生成する</p>
 * @param       {Object} p_params params
 * @param       {ClassifierInfoPayload[]} clf_list 分類器一覧
 * @param       {String} target_name 分類器名
 * @return      {ClfVers}  バージョン一覧
 */
function NLCAPP_clf_vers(p_params) {
  NLCAPP_log_debug({ record: ["NLCAPP_clf_vers", "START"] });
  NLCAPP_log_debug({
    record: ["NLCAPP_clf_vers", JSON.stringify(p_params.clf_list)]
  });

  var clfs = [];
  var max_ver = 0;
  var min_ver = 99999999;
  var count = 0;

  for (var i = 0; i < p_params.clf_list.length; i += 1) {
    var base = p_params.clf_list[i].name.split(CLF_SEP);

    if (p_params.target_name === base[0]) {
      clfs[parseInt(base[1], 10)] = p_params.clf_list[i];
      count += 1;

      if (parseInt(base[1], 10) > max_ver) {
        max_ver = parseInt(base[1], 10);
      }
      if (parseInt(base[1], 10) < min_ver) {
        min_ver = parseInt(base[1], 10);
      }
    }
  }

  var result = {
    count: count,
    min_ver: min_ver,
    max_ver: max_ver,
    clfs: clfs
  };
  return result;
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * @typedef {Object} ClfInfo 分類器情報
 * @property {String} clf_id ID
 * @property {String} status ステータス
 */
/**
 * 利用可能な分類器の最新バージョンを取得する
 * <p>ステータスコードが200以外の場合、IDに空白、ステータスにNLCの実行ステータスをセットする</p>
 * <p>バージョン件数が０件の場合、IDに空白、ステータスに'Nothing'をセットする</p>
 * <p>各バージョンの状態を取得する</p>
 * <p>状態が'Available'でバージョンが最新の分類器情報を返す</p>
 * @param       {NaturalLanguageClassifierV1} p_nlc NLC
 * @param       {String} clfs           分類器一覧
 * @param       {String} clf_name       分類器名 ex.CLF1
 * @return      {ClfInfo}  分類器情報
 */
function NLCAPP_select_clf(p_nlc, clfs, clf_name) {
  NLCAPP_log_debug({ record: ["NLCAPP_select_clf", "START"] });
  NLCAPP_log_debug({ record: ["NLCAPP_select_clf", JSON.stringify(clfs)] });

  if (clfs.status !== 200) {
    return {
      clf_id: "",
      status: "Error",
      code: clfs.status,
      description: clfs.body.error
    };
  }

  var clf_info = NLCAPP_clf_vers({
    clf_list: clfs.body.classifiers,
    target_name: clf_name
  });
  if (clf_info.count === 0) {
    return {
      clf_id: "",
      status: "Nothing"
    };
  }

  var clf;
  var res;
  for (var i = clf_info.max_ver; i >= clf_info.min_ver; i -= 1) {
    clf = clf_info.clfs[i];

    res = p_nlc.getClassifier({
      classifier_id: clf.classifier_id
    });
    if (res.body.status === "Available") {
      return {
        clf_id: clf.classifier_id,
        status: res.body.status
      };
    }
  }

  return {
    clf_id: clf.classifier_id,
    status: res.body.status
  };
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * メール通知条件チェック
 * <p>通知条件にマッチした場合はメールを送信する</p>
 * <p>件名と本文の埋め込みタグを展開する</p>
 * <p>インテントがブランクの場合はワイルドカード扱いする</p>
 * @param       {NotifSet} notif_set 通知設定
 * @param       {String[]} record    通知対象データ
 * @param       {Integer[]} upd_flg   更新フラグ
 */
function NLCAPP_check_notify(notif_set, record, upd_flg) {
  NLCAPP_log_debug({ record: ["NLCAPP_check_notify", "START"] });

  for (var i = 0; i < notif_set.rules.length; i += 1) {
    var chk_cnt = 0;
    var upd_chk = 0;

    for (var j = 0; j < NB_CLFS; j += 1) {
      if (notif_set.rules[i].res_int[j] === "") {
        chk_cnt += 1;
      } else {
        if (upd_flg[j] === 1) {
          upd_chk = 1;
        }
        if (
          String(record[notif_set.result_cols[j] - 1]) ===
          notif_set.rules[i].res_int[j]
        ) {
          chk_cnt += 1;
        }
      }
    }

    if (chk_cnt === NB_CLFS && upd_chk === 1) {
      var res;
      res = NLCAPP_expand_tags(notif_set.rules[i].mail.body, record);
      var body = res.text;

      res = NLCAPP_expand_tags(notif_set.rules[i].mail.subject, record);
      var subject = res.text;

      var mail_set = {
        from: notif_set.rules[i].mail.from,
        to: notif_set.rules[i].mail.to,
        cc: notif_set.rules[i].mail.cc,
        bcc: notif_set.rules[i].mail.bcc,
        subject: subject,
        body: body
      };

      GASLIB_Mail_send(mail_set);
    }
  }
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * 分類処理の初期化
 * @param {Object} CONFIG_SET CONFIG_SET
 * @return {Object} conf
 * @throws {Error} データシートが不明です
 */
function NLCAPP_init_classify(CONFIG_SET) {
  NLCAPP_log_debug({ record: ["NLCAPP_init_classify", "START"] });

  var creds = NLCAPP_load_creds();
  var conf = NLCAPP_load_config(CONFIG_SET);

  if (!RUNTIME_OPTION.UI_DISABLE || RUNTIME_OPTION.UI_DISABLE === false) {
    var SS_UI;
    try {
      SS_UI = SpreadsheetApp.getUi();
    } catch (e) {
      SS_UI = null;
    }

    if (SS_UI !== null) {
      var res = GASLIB_Dialog_open(
        "分類",
        "分類を開始します。よろしいですか？",
        SS_UI.ButtonSet.OK_CANCEL
      );
      if (res === SS_UI.Button.CANCEL) {
        GASLIB_Dialog_open("分類", "分類を中止しました。", SS_UI.ButtonSet.OK);
        return {};
      }
      var msg =
        "分類を開始しました。ログは「" +
        conf.sheet_conf.log_ws +
        "」シートをご参照ください。";
      GASLIB_Dialog_open("分類", msg, SS_UI.ButtonSet.OK);
    }
  }

  var notif_conf = {
    self_ss: conf.self_ss,
    ws_name: conf.notif_conf.ws_name,
    start_col: CONFIG_SET.notif_start_col,
    start_row: CONFIG_SET.notif_start_row
  };

  //TODO
  var rules = NLCAPP_load_notif_rules(notif_conf);

  var notif_set = {
    rules: rules,
    from: CONFIG_SET.notif_from,
    sender: CONFIG_SET.notif_sender,
    result_cols: conf.sheet_conf.result_col
  };

  var log_sheet = conf.self_ss.getSheetByName(conf.sheet_conf.log_ws);
  if (log_sheet === null) {
    log_sheet = conf.self_ss.insertSheet(conf.sheet_conf.log_ws);
  }

  var log_set = {
    sheet: log_sheet,
    start_col: CONFIG_SET.log_start_col,
    start_row: CONFIG_SET.log_start_row
  };

  var data_sheet = conf.self_ss.getSheetByName(conf.sheet_conf.ws_name);
  if (data_sheet === null) {
    throw new Error("データシートが不明です");
  }

  var test_set = {
    sheet: data_sheet,
    ws_name: conf.sheet_conf.ws_name,
    start_col: conf.sheet_conf.start_col,
    start_row: conf.sheet_conf.start_row,
    end_row: -1,
    text_col: conf.sheet_conf.text_col,
    notif_opt: conf.notif_conf.option,
    notif_set: notif_set
  };

  var nlc = NLCAPP_create_instance(creds.username, creds.password, creds.url);

  var clf_ids = NLCAPP_get_classifiers({
    nlc: nlc,
    log_set: log_set,
    test_set: test_set
  });

  return {
    nlc: nlc,
    clf_ids: clf_ids,
    sheet_conf: conf.sheet_conf,
    test_set: test_set,
    log_set: log_set,
    notif_set: notif_set
  };
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * 外部トリガー用分類処理
 */
function NLCAPP_classify_all() { // eslint-disable-line no-unused-vars
  NLCAPP_log_debug({ record: ["NLCAPP_classify_all", "START"] });

  if (RUNTIME_ACTIVE !== "NLCAPP") {
    throw new Error("このアプリでは実行できません");
  }

  var conf = NLCAPP_init_classify(CONFIG_SET);

  // 分類データの作成
  var classify_data = NLCAPP_create_classify_data(conf);

  var classify_results = NLCAPP_classify_base({
    nlc: conf.nlc,
    clf_ids: conf.clf_ids,
    classify_data: classify_data,
    test_set: conf.test_set,
    notif_set: conf.notif_set,
    sheet: conf.sheet,
    sheet_conf: conf.sheet_conf
  });

  NLCAPP_log_classify_results({
    clf_ids: conf.clf_ids,
    classify_results: classify_results,
    sheet_conf: conf.sheet_conf,
    log_set: conf.log_set,
    test_set: conf.test_set
  });
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * 分類用データの作成
 * @param {Object} p_params params
 * @param {Object} p_params.test_set 分類設定情報
 * @param {Object} p_params.test_set.sheet 分類データシート
 * @param {Object}  p_params.sheet_conf 設定シート情報
 * @param {Object}  p_params.sheet_conf.start_row 開始行
 * @param {Object}  p_params.sheet_conf.start_col 開始列
 * @param {Object}  p_params.sheet_conf.text_col テキスト列
 * @param {Object}  p_params.sheet_conf.result_col 結果列
 * @return {String[]} classify_data
 */
function NLCAPP_create_classify_data(p_params) {
  NLCAPP_log_debug({ record: ["NLCAPP_create_classify_data", "START"] });

  // 対象データの読み込み
  var lastRow = p_params.test_set.sheet.getLastRow();
  var lastCol = p_params.test_set.sheet.getLastColumn();

  var entries;
  if (
    lastRow < p_params.sheet_conf.start_row ||
    lastCol < p_params.sheet_conf.start_col
  ) {
    entries = [];
  } else {
    entries = p_params.test_set.sheet
      .getRange(
        p_params.sheet_conf.start_row,
        p_params.sheet_conf.start_col,
        lastRow - p_params.sheet_conf.start_row + 1,
        lastCol - p_params.sheet_conf.start_col + 1
      )
      .getValues();
  }

  var classify_data = [];
  for (var cnt = 0; cnt < entries.length; cnt += 1) {
    var test_text =
      entries[cnt][
        p_params.sheet_conf.text_col - p_params.sheet_conf.start_col
      ];
    test_text = GASLIB_Text_normalize(String(test_text)).trim();

    if (test_text.length > NLCLIB_MAX_TRAIN_STRINGS) {
      test_text = test_text.substring(0, NLCLIB_MAX_TRAIN_STRINGS);
    }

    // 分類実行フラグ 1:する 0:スキップ
    var flags = [1, 1, 1];
    if (test_text.length === 0) {
      flags = [0, 0, 0];
    } else {
      for (var j = 0; j < NB_CLFS; j += 1) {
        var clf_no = j;
        var result_text;
        var result_col = p_params.sheet_conf.result_col[clf_no];
        if (lastCol < p_params.sheet_conf.result_col[clf_no]) {
          result_text = "";
        } else {
          result_text =
            entries[cnt][result_col - p_params.sheet_conf.start_col];
        }

        // 結果が存在し、上書き不可
        if (result_text !== "" && CONFIG_SET.result_override !== true) {
          flags[j] = 0;
        }
      }
    }

    classify_data.push({
      test_text: test_text,
      flags: flags
    });
  }
  return classify_data;
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * イベント実行用分類処理
 * @param  {Object} p_params params
 * @param  {Object} p_params.nlc params
 * @param  {Object} p_params.clf_ids clf_ids
 * @param  {Object} p_params.clf_ids.status clf_ids
 * @param  {Object} p_params.clf_ids.clf_name clf_ids
 * @param  {Object} p_params.clf_ids.id clf_ids
 * @param  {Object} p_params.classify_data params
 * @param  {Object} p_params.classify_data.test_text params
 * @param  {Object} p_params.classify_data.flags params
 * @param  {Object} p_params.test_set params
 * @param  {Object} p_params.test_set.notif_opt params
 * @param  {Object} p_params.notif_set params
 * @param  {Object} p_params.sheet params
 * @param  {Object} p_params.sheet_conf params
 * @param  {Object} p_params.sheet_conf.result_col params
 * @param  {Object} p_params.sheet_conf.start_row params
 * @param  {Object} p_params.sheet_conf.restime_col params
 * @param  {Object} p_params.sheet_conf.resconf_col params
 * @return {Object} classify results
 */
function NLCAPP_classify_base(p_params) {
  NLCAPP_log_debug({ record: ["NLCAPP_classify_base", "START"] });
  NLCAPP_log_debug({
    record: ["NLCAPP_classify_base", JSON.stringify(p_params.clf_ids)]
  });

  // 分類
  var hasError = 0;
  var err_res;
  var nlc_res;
  var res_rows = [0, 0, 0]; // 件数
  for (var cnt = 0; cnt < p_params.classify_data.length; cnt += 1) {
    var test_text = p_params.classify_data[cnt].test_text;
    var flags = p_params.classify_data[cnt].flags;

    var updates = 0;
    var upd_flg = [0, 0, 0];
    // 各分類器
    for (var clf_no = 0; clf_no < NB_CLFS; clf_no += 1) {
      if (p_params.clf_ids[clf_no].status !== "Available") continue;
      if (flags[clf_no] === 0) continue;

      NLCAPP_log_debug({ record: ["nlc.classify", "START"] });
      nlc_res = p_params.nlc.classify({
        classifier_id: p_params.clf_ids[clf_no].id,
        text: test_text
      });
      NLCAPP_log_debug({ record: ["nlc.classify", JSON.stringify(nlc_res)] });
      if (nlc_res.status !== 200) {
        hasError = 1;
        err_res = nlc_res;
      } else {
        //NLCAPP_write_classify_results()
        // データシート
        var r = nlc_res.body.top_class;
        p_params.test_set.sheet
          .getRange(
            p_params.sheet_conf.start_row + cnt,
            p_params.sheet_conf.result_col[clf_no],
            1,
            1
          )
          .setNumberFormat("@")
          .setValue(r);

        var t = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss");
        p_params.test_set.sheet
          .getRange(
            p_params.sheet_conf.start_row + cnt,
            p_params.sheet_conf.restime_col[clf_no],
            1,
            1
          )
          .setValue(t);

        var c = nlc_res.body.classes[0].confidence;
        p_params.test_set.sheet
          .getRange(
            p_params.sheet_conf.start_row + cnt,
            p_params.sheet_conf.resconf_col[clf_no],
            1,
            1
          )
          .setValue(c);

        updates += 1;
        res_rows[clf_no] += 1;
        upd_flg[clf_no] = 1;
      }
    } // end for clf_no

    // 要通知
    if (p_params.test_set.notif_opt === NOTIF_OPT.ON) {
      if (updates > 0) {
        var record = p_params.test_set.sheet
          .getRange(
            p_params.sheet_conf.start_row + cnt,
            1,
            1,
            p_params.test_set.sheet.getLastColumn()
          )
          .getValues();

        NLCAPP_check_notify(p_params.notif_set, record[0], upd_flg);
      } // end if
    } // end if
  } // end for cnt

  return {
    res_rows: res_rows,
    hasError: hasError,
    nlc_res: nlc_res,
    err_res: err_res
  };
}
// ------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * [NLCAPP_log_classify_results description]
 * @param       {Object} p_params params
 * @param       {Object} p_params.clf_ids params
 * @param       {Object} p_params.clf_ids.status params
 * @param       {Object} p_params.classify_results params
 * @param       {Object} p_params.classify_results.res_rows params
 * @param       {Object} p_params.classify_results.nlc_res params
 * @param       {Object} p_params.classify_results.nlc_res.status params
 * @param       {Object} p_params.classify_results.hasError params
 * @param       {Object} p_params.classify_results.err_res params
 * @param       {Object} p_params.classify_results.err_res.status params
 * @param       {Object} p_params.sheet_conf params
 * @param       {Object} p_params.sheet_conf.result_col params
 * @param       {Object} p_params.sheet_conf.restime_col params
 * @param       {Object} p_params.test_set params
 * @param       {Object} p_params.log_set params
 */
function NLCAPP_log_classify_results(p_params) {
  NLCAPP_log_debug({ record: ["NLCAPP_log_classify_results", "START"] });
  NLCAPP_log_debug({
    record: [
      "NLCAPP_log_classify_results",
      JSON.stringify(p_params.classify_results.res_rows)
    ]
  });

  // 分類を実行した件数を分類器ごとにログ出力
  for (var k = 0; k < NB_CLFS; k += 1) {
    if (p_params.clf_ids[k].status !== "Available") continue;

    var clfname = CLFNAME_PREFIX + String(k + 1);
    var test_set = p_params.test_set;
    test_set["clf_no"] = k + 1;
    test_set["clf_name"] = clfname;
    test_set["result_col"] = p_params.sheet_conf.result_col[k];
    test_set["restime_col"] = p_params.sheet_conf.restime_col[k];

    var test_result;
    // ０件
    if (p_params.classify_results.res_rows[k] === 0) {
      test_result = {
        status: 0,
        rows: 0
      };
    } else {
      //エラーあり
      if (p_params.classify_results.hasError === 1) {
        test_result = {
          status: p_params.classify_results.err_res.status,
          rows: p_params.classify_results.res_rows[k],
          nlc: p_params.classify_results.err_res
        };
        // エラーなし
      } else {
        test_result = {
          status: p_params.classify_results.nlc_res.status,
          rows: p_params.classify_results.res_rows[k],
          nlc: p_params.classify_results.nlc_res
        };
      } // end if hasError
    } // end if res_rows

    NLCAPP_log_classify({
      log_set: p_params.log_set,
      test_set: test_set,
      test_result: test_result
    });
  } // end for
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * 設定シートの分類器情報をセットする
 * @param       {Object} p_params パラメータ
 * @param       {Int} p_params.clf_no 分類器番号
 * @param       {String} p_params.status 分類器のステータス
 * @param       {Spreadsheet} p_params.sheet 設定シート
 */
function NLCAPP_set_classifier_status(p_params) {
  p_params.sheet
    .getRange(
      CONFIG_SET.clfs_start_row + (p_params.clf_no - 1),
      CONFIG_SET.clfs_start_col,
      1,
      2
    )
    .setValues([[p_params.classifier_id, p_params.status]]);
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * タイマー起動用分類器状態チェック
 */
function NLCAPP_exec_check_clfs() { // eslint-disable-line no-unused-vars
  NLCAPP_log_debug({ record: ["NLCAPP_exec_check_clfs", "START"] });

  var creds = NLCAPP_load_creds();
  var conf = NLCAPP_load_config(CONFIG_SET);

  var nlc = NLCAPP_create_instance(creds.username, creds.password, creds.url);
  var clfs = NLCAPP_list_classifiers({
    nlc: nlc,
    logging: false
  });

  var log_sheet = conf.self_ss.getSheetByName(conf.sheet_conf.log_ws);
  if (log_sheet === null) {
    log_sheet = conf.self_ss.insertSheet(conf.sheet_conf.log_ws);
  }

  var log_set = {
    sheet: log_sheet,
    start_col: CONFIG_SET.log_start_col,
    start_row: CONFIG_SET.log_start_row
  };

  NLCAPP_check_classifiers({
    nlc: nlc,
    clfs: clfs,
    sheet_conf: conf.sheet_conf,
    system_conf: CONFIG_SET,
    log_set: log_set
  });
  NLCAPP_log_debug({ record: ["NLCAPP_exec_check_clfs", "END"] });
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * トレーニングデータの作成
 * @param {Object} p_params params
 * @param {Object} p_params.train_set. params
 * @param {Object} p_params.train_set.sheet params
 * @param {Object} p_params.train_set.start_row. params
 * @param {Object} p_params.train_set.class_col. params
 * @param {Object} p_params.train_set.text_col. params
 * @return {Object} トレーニングデータ
 */
function NLCAPP_create_training_data(p_params) {
  NLCAPP_log_debug({ record: ["NLCAPP_create_training_data", "START"] });

  var EOL = "\r\n";

  var lastRow = p_params.train_set.sheet.getLastRow();
  var lastCol = p_params.train_set.sheet.getLastColumn();

  var entries;
  //class_col = conf.sheet_conf.intent_col[clf_no - 1],

  if (
    lastRow < p_params.train_set.start_row ||
    lastCol < p_params.train_set.class_col ||
    lastCol < p_params.train_set.text_col
  ) {
    entries = [];
  } else {
    entries = p_params.train_set.sheet
      .getRange(
        p_params.train_set.start_row,
        1,
        lastRow - p_params.train_set.start_row + 1,
        lastCol
      )
      .setNumberFormat("@")
      .getValues();
  }
  var train_buf = [];
  var row_cnt = 0;
  var csvString = "";
  for (var i = 0; i < entries.length; i += 1) {
    var class_name = GASLIB_Text_normalize(
      String(entries[i][p_params.train_set.class_col - 1])
    );
    if (class_name.length === 0) continue;

    var train_text = String(entries[i][p_params.train_set.text_col - 1]);
    train_text = GASLIB_Text_normalize(train_text);
    if (train_text.length === 0) continue;

    if (train_text.length > NLCLIB_MAX_TRAIN_STRINGS) {
      train_text = train_text.substring(0, NLCLIB_MAX_TRAIN_STRINGS);
    }

    train_buf.push({
      text: train_text,
      class: class_name
    });

    row_cnt += 1;
  }

  if (row_cnt === 0) {
    return {
      row_cnt: 0,
      csvString: ""
    };
  }

  var train_data = [];
  if (row_cnt > NLCLIB_MAX_TRAIN_RECORDS) {
    train_data = train_buf.splice(
      row_cnt - NLCLIB_MAX_TRAIN_RECORDS,
      row_cnt - 1
    );
  } else {
    train_data = train_buf;
  }

  // クラス件数のチェック
  /*
  var uniq_classes = [];
  train_data.forEach(function (record) {
    if (uniq_classes.indexOf(record.class) === -1) {
      uniq_classes.push(record.class);
    }
  });

  if (uniq_classes.length > MAX_TRAIN_CLASSES) {
    return {
      row_cnt: row_cnt,
      status: 1,
      description:
        "ユニークなクラスが" + MAX_TRAIN_CLASSES + "件を超過しています",
      error_desc: "クラスの種類:" + uniq_classes.length
    };
  }
  */

  for (var tcnt = 0; tcnt < train_data.length; tcnt += 1) {
    csvString =
      csvString +
      '"' +
      train_data[tcnt].text +
      '","' +
      train_data[tcnt].class +
      '"' +
      EOL;
  }

  return {
    row_cnt: row_cnt,
    csvString: csvString
  };
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * 学習処理の初期化
 * @return {Object} conf
 *     nlc: nlc,
    clfs: clfs,
    sheet_conf: conf.sheet_conf,
    sheet: sheet,
    log_set: log_set,

 */
function NLCAPP_init_train() {
  NLCAPP_log_debug({ record: ["NLCAPP_init_train", "START"] });

  var creds = NLCAPP_load_creds();
  var conf = NLCAPP_load_config(CONFIG_SET);

  if (!RUNTIME_OPTION.UI_DISABLE || RUNTIME_OPTION.UI_DISABLE === false) {
    var ss_ui;
    try {
      ss_ui = SpreadsheetApp.getUi();
    } catch (e) {
      ss_ui = null;
    }
    if (ss_ui != null) {
      var res = GASLIB_Dialog_open(
        "学習",
        "学習を開始します。よろしいですか？",
        ss_ui.ButtonSet.OK_CANCEL
      );
      if (res === ss_ui.Button.CANCEL) {
        GASLIB_Dialog_open("学習", "学習を中止しました。", ss_ui.ButtonSet.OK);
        return {};
      }

      var msg =
        "学習を開始しました。ログは「" +
        conf.sheet_conf.log_ws +
        "」シートをご参照ください。";
      msg +=
        "\nステータスは「" + CONFIG_SET.ws_name + "」シートをご参照ください。";
      GASLIB_Dialog_open("学習", msg, ss_ui.ButtonSet.OK);
    }
  }

  var log_sheet = conf.self_ss.getSheetByName(conf.sheet_conf.log_ws);
  if (log_sheet === null) {
    log_sheet = conf.self_ss.insertSheet(conf.sheet_conf.log_ws);
  }

  var log_set = {
    sheet: log_sheet,
    start_col: CONFIG_SET.log_start_col,
    start_row: CONFIG_SET.log_start_row
  };

  var nlc = NLCAPP_create_instance(creds.username, creds.password, creds.url);

  var clfs = NLCAPP_list_classifiers({
    nlc: nlc
  });
  if (clfs.status !== 200) {
    NLCAPP_log_train({
      log_set: log_set,
      train_set: {
        clf_no: "N/A",
        ws_name: "N/A",
        text_col: "N/A",
        class_col: "N/A"
      },
      train_result: {
        status: clfs.status,
        description: "エラー",
        error_desc:
          typeof clfs.body === "object" && "error" in clfs.body
            ? clfs.body.error
            : clfs.body
      }
    });
    throw new Error("分類器のエラーです。ログを確認してください");
  }

  return {
    nlc: nlc,
    clfs: clfs,
    conf: conf,
    sheet_conf: conf.sheet_conf,
    log_set: log_set
  };
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * 学習処理
 * @param       {Object} p_params params
 * @param       {Object} p_params.nlc params
 * @param       {Object} p_params.conf params
 * @param       {Object} p_params.conf.clfs params
 * @param       {Object} p_params.conf.SS_ID params
 * @param       {Object} p_params.conf.sheet params
 * @param       {Object} p_params.sheet_conf params
 * @param       {Object} p_params.mail_conf params
 * @param       {String} creds_username クレデンシャル
 * @param       {String} creds_password クレデンシャル
 * @throws      {Error}  データシートが不明です
 * @throws      {Error}  学習・分類対象が不正です
 */
function NLCAPP_train_common(p_params) {
  NLCAPP_log_debug({ record: ["NLCAPP_train_common", "START"] });

  var clf_info = NLCAPP_clf_vers({
    clf_list: p_params.clfs.body.classifiers,
    target_name: p_params.train_set.clf_name
  });
  var new_version = clf_info.max_ver + 1;
  var new_name = p_params.train_set.clf_name + CLF_SEP + new_version;

  NLCAPP_log_debug({
    record: [
      "NLCAPP_train_common",
      "ROWS:" + String(p_params.training_data.row_cnt)
    ]
  });

  var del_params = {};
  var train_result = {};
  if (p_params.training_data.row_cnt === 0) {
    train_result = {
      status: 0,
      description: "学習データなし",
      rows: 0,
      version: new_version
    };

    // 分類器の削除
    if (clf_info.count > 0) {
      del_params = {
        classifier_id: clf_info.clfs[clf_info.min_ver].classifier_id
      };
      p_params.nlc.deleteClassifier(del_params);
    }

    NLCAPP_set_classifier_status({
      sheet: p_params.sheet_conf.sheet,
      clf_no: p_params.train_set.clf_no,
      classifier_id: "",
      status: ""
    });

    NLCAPP_log_train({
      log_set: p_params.log_set,
      train_set: p_params.train_set,
      train_result: train_result
    });

    return;
  }

  // 分類器の作成
  var train_params = {
    metadata: {
      name: new_name,
      language: "ja"
    },
    training_data: p_params.training_data.csvString
  };

  NLCAPP_log_debug({
    record: ["nlc.createClassifier", "START", JSON.stringify(train_params)]
  });
  var nlc_res = p_params.nlc.createClassifier(train_params);
  NLCAPP_log_debug({
    record: ["nlc.createClassifier", "END", JSON.stringify(nlc_res)]
  });
  if (nlc_res.status === 200) {
    NLCAPP_set_classifier_status({
      sheet: p_params.sheet_conf.sheet,
      clf_no: p_params.train_set.clf_no,
      classifier_id: nlc_res.body.classifier_id,
      status: nlc_res.body.status
    });

    // バージョンが複数ある場合
    if (clf_info.count >= 2 && nlc_res.status === 200) {
      // 分類器の削除
      del_params = {
        classifier_id: clf_info.clfs[clf_info.min_ver].classifier_id
      };
      p_params.nlc.deleteClassifier(del_params);
    }
  }

  train_result = {
    status: nlc_res.status,
    nlc: nlc_res,
    rows: p_params.training_data.row_cnt,
    version: new_version
  };

  NLCAPP_log_train({
    log_set: p_params.log_set,
    train_set: p_params.train_set,
    train_result: train_result
  });
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * イベント実行用学習処理
 */
function NLCAPP_train_all() { // eslint-disable-line no-unused-vars
  NLCAPP_log_debug({ record: ["NLCAPP_train_all", "START"] });

  if (RUNTIME_ACTIVE !== "NLCAPP") {
    throw new Error("このアプリでは実行できません");
  }

  var conf = NLCAPP_init_train();

  var data_sheet = conf.conf.self_ss.getSheetByName(conf.sheet_conf.ws_name);

  for (var clf_no = 1; clf_no <= NB_CLFS; clf_no += 1) {
    var train_set = {
      sheet: data_sheet,
      ws_name: conf.sheet_conf.ws_name,
      start_row: conf.sheet_conf.start_row,
      start_col: conf.sheet_conf.start_col,
      end_row: -1,
      text_col: conf.sheet_conf.text_col,
      class_col: conf.sheet_conf.intent_col[clf_no - 1],
      clf_no: clf_no,
      clf_name: CLFNAME_PREFIX + clf_no
    };

    var training_data = NLCAPP_create_training_data({
      train_set: train_set
    });

    NLCAPP_train_common({
      nlc: conf.nlc,
      clfs: conf.clfs,
      train_set: train_set,
      training_data: training_data,
      log_set: conf.log_set,
      sheet_conf: conf.sheet_conf
    });
  }

  NLCAPP_log_debug({ record: ["NLCAPP_train_all", "TRIGGER SET"] });
  GASLIB_Trigger_set("NLCAPP_exec_check_clfs", 1);

  NLCAPP_log_debug({ record: ["NLCAPP_train_all", "END"] });
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * 分類器削除
 * @param       {Integer} clf_no 分類器番号
 * @throws      {Error} 分類器のエラーです。ログを確認してください
 */
function NLCAPP_delete_classifier(clf_no) {
  NLCAPP_log_debug({ record: ["NLCAPP_delete_classifier", "START"] });

  var creds = NLCAPP_load_creds();
  var conf = NLCAPP_load_config(CONFIG_SET);

  var SS_UI;
  try {
    SS_UI = SpreadsheetApp.getUi();
  } catch (e) {
    SS_UI = null;
  }

  var log_sheet = conf.self_ss.getSheetByName(conf.sheet_conf.log_ws);
  if (log_sheet === null) {
    log_sheet = conf.self_ss.insertSheet(conf.sheet_conf.log_ws);
  }

  var log_set = {
    sheet: log_sheet,
    start_col: CONFIG_SET.log_start_col,
    start_row: CONFIG_SET.log_start_row
  };

  var nlc = NLCAPP_create_instance(creds.username, creds.password, creds.url);

  var clfs = NLCAPP_list_classifiers({
    nlc: nlc
  });
  if (clfs.status !== 200) {
    NLCAPP_log_delete({
      log_set: log_set,
      del_set: {
        clf_no: clf_no,
        clf_id: "N/A"
      },
      del_result: {
        status: clfs.body.error,
        code: clfs.body.code
      }
    });
    throw new Error("分類器のエラーです。ログを確認してください");
  }

  var msg;
  var clf_name = CLFNAME_PREFIX + String(clf_no);
  var clf_info = NLCAPP_clf_vers({
    clf_list: clfs.body.classifiers,
    target_name: clf_name
  });

  if (!RUNTIME_OPTION.UI_DISABLE || RUNTIME_OPTION.UI_DISABLE === false) {
    if (SS_UI != null) {
      if (clf_info.count === 0) {
        msg = "分類器" + String(clf_no) + "は存在しません。";
        GASLIB_Dialog_open("削除", msg, SS_UI.ButtonSet.OK);
        return;
      }

      msg = "分類器" + String(clf_no) + "を削除します。よろしいですか？";
      var res = GASLIB_Dialog_open("削除", msg, SS_UI.ButtonSet.OK_CANCEL);
      if (res === SS_UI.Button.CANCEL) {
        GASLIB_Dialog_open("削除", "削除を中止しました。", SS_UI.ButtonSet.OK);
        return;
      }
    }
  }

  for (var i = clf_info.min_ver; i <= clf_info.max_ver; i += 1) {
    var nlc_res = nlc.deleteClassifier({
      classifier_id: clf_info.clfs[i].classifier_id
    });

    var del_set = {
      clf_no: clf_no,
      clf_id: clf_info.clfs[i].classifier_id
    };
    var del_result = {
      status: nlc_res.status,
      nlc: nlc_res
    };

    NLCAPP_log_delete({
      log_set: log_set,
      del_set: del_set,
      del_result: del_result
    });
  }

  clfs = NLCAPP_list_classifiers({
    nlc: nlc
  });
  if (clfs.status !== 200) {
    NLCAPP_log_delete({
      log_set: log_set,
      del_set: {
        clf_no: clf_no,
        clf_id: "N/A"
      },
      del_result: {
        status: clfs.body.error,
        code: clfs.body.code
      }
    });
    throw new Error("分類器のエラーです。ログを確認してください");
  }

  NLCAPP_check_classifiers({
    nlc: nlc,
    clfs: clfs,
    sheet_conf: conf.sheet_conf,
    system_conf: CONFIG_SET,
    log_set: log_set
  });
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * 分類器1削除
 */
function NLCAPP_del_clf1() { // eslint-disable-line no-unused-vars
  NLCAPP_log_debug({ record: ["NLCAPP_del_clf1", "START"] });
  NLCAPP_delete_classifier(1);
}

/**
 * 分類器2削除
 */
function NLCAPP_del_clf2() { // eslint-disable-line no-unused-vars
  NLCAPP_log_debug({ record: ["NLCAPP_del_clf2", "START"] });
  NLCAPP_delete_classifier(2);
}

/**
 * 分類器3削除
 */
function NLCAPP_del_clf3() { // eslint-disable-line no-unused-vars
  NLCAPP_log_debug({ record: ["NLCAPP_del_clf3", "START"] });
  NLCAPP_delete_classifier(3);
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * 分類器の一覧取得
 * <p>分類器の一覧を取得する</p>
 * @param       {Object} p_params params
 * @param       {NaturalLanguageClassifierV1} p_params.nlc NLC
 * @param       {Boolean} p_params.logging ログ出力オプション
 * @return      {Object} clfs
 */
function NLCAPP_list_classifiers(p_params) {
  NLCAPP_log_debug({ record: ["NLCAPP_list_classifiers", "START"] });

  var conf = NLCAPP_load_config(CONFIG_SET);
  var log_sheet = conf.self_ss.getSheetByName(conf.sheet_conf.log_ws);
  if (log_sheet === null) {
    log_sheet = conf.self_ss.insertSheet(conf.sheet_conf.log_ws);
  }

  NLCAPP_log_debug({ record: ["nlc.listClassifiers", "START"] });
  var clfs = p_params.nlc.listClassifiers({
    muteHttpExceptions: true
  });
  NLCAPP_log_debug({ record: ["nlc.listClassifiers", JSON.stringify(clfs)] });
  if (clfs.status !== 200) {
    return clfs;
  }

  var clf_list = clfs.body.classifiers;

  var norm_list = [];
  var settings;
  var record;
  var colors;
  var params;
  for (var i = 0; i < clf_list.length; i += 1) {
    if (!clf_list[i].name) {
      if (
        typeof p_params.logging === "undefined" ||
        p_params.logging === true
      ) {
        settings = {
          sheet: log_sheet,
          start_row: CONFIG_SET.log_start_row,
          start_col: CONFIG_SET.log_start_col
        };
        record = [
          "共通",
          Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss"),
          "スクリプトで管理されていない分類器が存在します",
          "(名前なし)",
          "",
          "",
          "",
          "",
          clf_list[i].classifier_id
        ];
        colors = [
          "red",
          "red",
          "red",
          "red",
          "red",
          "red",
          "red",
          "red",
          "red"
        ];
        params = {
          settings: settings,
          record: record,
          colors: colors
        };

        GASLIB_SheetLog_write(params);
      }
      continue;
    }

    var base = clf_list[i].name.split(CLF_SEP);
    if (base.length !== 2) {
      if (
        typeof p_params.logging === "undefined" ||
        p_params.logging === true
      ) {
        settings = {
          sheet: log_sheet,
          start_row: CONFIG_SET.log_start_row,
          start_col: CONFIG_SET.log_start_col
        };
        record = [
          "共通",
          Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss"),
          "スクリプトで管理されていない分類器が存在します",
          clf_list[i].name,
          "",
          "",
          "",
          "",
          clf_list[i].classifier_id
        ];
        colors = [
          "red",
          "red",
          "red",
          "red",
          "red",
          "red",
          "red",
          "red",
          "red"
        ];
        params = {
          settings: settings,
          record: record,
          colors: colors
        };
        GASLIB_SheetLog_write(params);
      }
      continue;
    }

    norm_list.push(clf_list[i]);
  }

  clfs.body.classifiers = norm_list;

  //NLCAPP_log_debug({ record: ["NLCAPP_list_classifiers", "END"] });
  return clfs;
}

// ------------------------------------------------------------------------
/**
 * 分類器情報の取得
 * @param       {Object} p_params params
 * @param       {NaturalLanguageClassifierV1} p_params.nlc NLC
 * @param       {Object} p_params.log_set ログ設定
 * @param       {Object} p_params.test_set テスト設定
 * @return      {Object} 分類器情報
 */
function NLCAPP_get_classifiers(p_params) {
  NLCAPP_log_debug({ record: ["NLCAPP_get_classifiers", "START"] });

  var clfs = NLCAPP_list_classifiers({
    nlc: p_params.nlc
  });

  var clf_ids = [];
  var test_set = p_params.test_set;
  for (var i = 0; i < NB_CLFS; i += 1) {
    var clf_name = CLFNAME_PREFIX + String(i + 1);
    test_set.clf_no = i + 1;
    test_set.result_col = "";

    var clf = NLCAPP_select_clf(p_params.nlc, clfs, clf_name);
    if (clf.status === "Training") {
      NLCAPP_log_classify({
        log_set: p_params.log_set,
        test_set: test_set,
        test_result: {
          status: 900,
          description: "トレーニング中",
          clf_id: clf.clf_id
        }
      });
    } else if (clf.status === "Nothing") {
      NLCAPP_log_classify({
        log_set: p_params.log_set,
        test_set: test_set,
        test_result: {
          status: 900,
          description: "分類器なし",
          clf_id: ""
        }
      });
    } else if (clf.status === "Error") {
      NLCAPP_log_classify({
        log_set: p_params.log_set,
        test_set: test_set,
        test_result: {
          status: clf.code,
          description: clf.description,
          clf_id: clf.clf_id
        }
      });
    } else if (clf.status !== "Available") {
      NLCAPP_log_classify({
        log_set: p_params.log_set,
        test_set: test_set,
        test_result: {
          status: 800,
          description: clf.status,
          clf_id: ""
        }
      });
    }

    clf_ids.push({
      id: clf.clf_id,
      name: clf_name,
      status: clf.status
    });
  }

  return clf_ids;
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * @typedef {Object} SheetSet 資格情報
 * @property {String}  ユーザー名
 */
/**
 * 分類器の状態確認
 * @param       {Object} p_params params
 * @param       {Object} p_params.sheet_conf params
 * @param       {Object} p_params.sheet_conf.sheet params
 * @param       {Object} p_params.sheet_conf.start_row params
 * @param       {Object} p_params.sheet_conf.start_col params
 * @param       {Object} p_params.clfs params
 * @param       {Object} p_params.clfs.status params
 * @param       {Object} p_params.clfs.body params
 * @param       {Object} p_params.clfs.classifiers params
 * @param       {Object} p_params.clfs.body.error params
 * @throws      {Error} 設定シートが不明です
 */
function NLCAPP_check_classifiers(p_params) {
  NLCAPP_log_debug({ record: ["NLCAPP_check_classifiers", "START"] });

  var log_set = {
    sheet: p_params.log_set.sheet,
    start_col: p_params.log_set.start_col,
    start_row: p_params.log_set.start_row
  };

  var stats_range = p_params.sheet_conf.sheet.getRange(
    p_params.system_conf.clfs_start_row,
    p_params.system_conf.clfs_start_col,
    NB_CLFS,
    2
  );

  var curr_stats = stats_range.getValues();
  stats_range.clear();

  var all_status = 0;
  for (var cnt = 1; cnt <= NB_CLFS; cnt += 1) {
    if (p_params.clfs.status !== 200) {
      p_params.sheet_conf.sheet
        .getRange(
          p_params.system_conf.clfs_start_row + (cnt - 1),
          p_params.system_conf.clfs_start_col,
          1,
          2
        )
        .setValues([["ERROR", p_params.clfs.body.error]]);
    } else {
      var clf_name = CLFNAME_PREFIX + String(cnt);

      var clf_info = NLCAPP_clf_vers({
        clf_list: p_params.clfs.body.classifiers,
        target_name: clf_name
      });
      if (clf_info.count === 0) {
        p_params.sheet_conf.sheet
          .getRange(
            p_params.system_conf.clfs_start_row + (cnt - 1),
            p_params.system_conf.clfs_start_col,
            1,
            2
          )
          .setValues([["", ""]]);
        all_status += 1;
      } else {
        var clf = clf_info.clfs[clf_info.max_ver];

        var get_params = {
          classifier_id: clf.classifier_id
        };
        NLCAPP_log_debug({ record: ["NLC.getClassifier", "START"] });
        var res = p_params.nlc.getClassifier(get_params);
        if (res.body.status === "Failed") {
          all_status += 1;
          var train_set2 = {
            clf_no: cnt,
            ws_name: "",
            text_col: "",
            class_col: ""
          };
          var train_result2 = {
            status: 999,
            description: "Failed",
            version: clf_info.max_ver,
            error_desc: res.body.status_description
          };

          NLCAPP_log_train({
            log_set: log_set,
            train_set: train_set2,
            train_result: train_result2
          });
        } else if (res.body.status === "Available") {
          all_status += 1;
          if (
            curr_stats[cnt - 1][0] === clf.classifier_id &&
            curr_stats[cnt - 1][1] === "Training"
          ) {
            var train_set = {
              clf_no: cnt
            };
            var train_result = {
              nlc: res,
              status: 2000,
              version: clf_info.max_ver
            };
            NLCAPP_log_train({
              log_set: log_set,
              train_set: train_set,
              train_result: train_result
            });
          }
        }

        p_params.sheet_conf.sheet
          .getRange(
            p_params.system_conf.clfs_start_row + (cnt - 1),
            p_params.system_conf.clfs_start_col,
            1,
            2
          )
          .setValues([[clf.classifier_id, res.body.status]]);
      }
    }
  }

  // 全てAvailableでタイマー解除
  if (NB_CLFS === all_status) {
    NLCAPP_log_debug({ record: ["NLCAPP_check_classifiers", "DEL TRIGGER"] });
    GASLIB_Trigger_del("NLCAPP_exec_check_clfs");
  }

  NLCAPP_log_debug({ record: ["NLCAPP_check_classifiers", "END"] });
}
// ----------------------------------------------------------------------------
