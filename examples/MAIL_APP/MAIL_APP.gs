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

// ----------------------------------------------------------------------------
// グローバル変数
/* globals GASLIB_Dialog_open */
/* globals GASLIB_Text_normalize */
/* globals GASLIB_Text_escape_formula */
/* globals GASLIB_Trigger_set */

/* globals CONFIG_SET */
/* globals CONF_INDEX */
/* globals NB_CLFS */
/* globals CONF_INDEX */
/* globals CLFNAME_PREFIX */
/* globals RUNTIME_CONFIG */
/* globals RUNTIME_OPTION */
/* globals NLCAPP_load_creds */
/* globals NLCAPP_log_train */
/* globals NLCAPP_load_notif_rules */
/* globals NLCAPP_list_classifiers */
/* globals NLCAPP_get_classifiers */
/* globals NLCAPP_load_config */
/* globals NLCAPP_create_instance */
/* globals NLCAPP_log_classify_results */
/* globals NLCAPP_classify_base */
/* globals NLCAPP_train_common */
/* globals NLCAPP_log_debug */

/* globals NLCLIB_MAX_TRAIN_RECORDS */
/* globals NLCLIB_MAX_TRAIN_STRINGS */

/**
 * メールデータフィールドインデックス
 * @type {Object}
 * @property {Integer} ID メールID
 * @property {Integer} DATE 受信日時
 * @property {Integer} SUBJECT 件名
 * @property {Integer} FROM 送信元メールアドレス
 * @property {Integer} TO 送信先メールアドレス
 * @property {Integer} CC CCメールアドレス
 * @property {Integer} BODY 本文
 * @property {Integer} NORM_BODY 本文(処理対象)
 */
var MAIL_FIELDS = {
  ID: 0,
  DATE: 1,
  SUBJECT: 2,
  FROM: 3,
  TO: 4,
  CC: 5,
  BODY: 6,
  NORM_BODY: 7
};

/**
 * 学習対象選択オプション
 * @type {Object}
 * @property {String} SUBJECT 件名のみ
 * @property {String} BODY 本文
 * @property {String} BOTH 両方
 */
var TRAIN_COLUMN = {
  SUBJECT: "件名のみ",
  BODY: "本文のみ",
  BOTH: "件名・本文両方"
};

/**
 * GmailApp search 最大スレッド数
 * @type {Integer}
 */
var MAX_THREADS = 500;
// ----------------------------------------------------------------------------

/**
 * メール本文の文字数上限(超過は切り捨て)
 * @type {Integer}
 */
var BODY_LENGTH_LIMIT = 2000; // eslint-disable-line no-unused-vars

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
 * @property {String}    ws_name      データシート名
 * @property {Integer}   start_row    定義開始行
 * @property {Integer}   start_col    定義開始列
 * @property {Integer[]} intent_col   インテント列1to3
 * @property {Integer[]} result_col   分類結果列1to3
 * @property {Integer[]} resconf_col  確信度列1to3
 * @property {Integer[]} restime_col  分類日時列1to3
 * @property {String}    log_ws       ログシート名
 * @property {String}    query        フィルタクエリ
 * @property {Integer}   search_limit 最大取得スレッド数
 * @property {Integer}   ago_days     過去分取得日数
 */
/**
 * @typedef {Object} NotifConf 通知設定
 * @property {String} notif_opt 通知オプション{On,Off}
 * @property {String} notif_ws 設定シート名
 */
/**
 * @typedef {Object} ExcConf 本文除外設定
 * @property {String[]} re_list 正規表現リスト
 */
/**
 * @typedef {Object} Config 設定情報
 * @property {SheetConf} sheet_conf データシート設定
 * @property {NotifConf} notif_conf 通知設定
 * @property {ExcConf}   exc_conf   本文除外設定
 */
/**
 * 設定情報の取得
 * @param       {ConfigMeta} config_set メタデータ
 * @return      {Config} 設定情報
 * @throws      {Error}  設定シートが不明です
 * @throws      {Error}  設定シートに問題があります
 */
function MAILAPP_load_config(config_set) {
  NLCAPP_log_debug({ record: ["MAILAPP_load_config", "START"] });

  var nlc_conf = NLCAPP_load_config(config_set);

  var nb_conf = Object.keys(CONF_INDEX).length;

  var lastRow = nlc_conf.sheet_conf.sheet.getLastRow();
  var lastCol = nlc_conf.sheet_conf.sheet.getLastColumn();

  if (
    lastRow < config_set.st_start_row + nb_conf - 1 ||
    lastCol < config_set.st_start_col
  ) {
    throw new Error("設定シートに問題があります");
  }

  var conf_list = nlc_conf.sheet_conf.sheet
    .getRange(config_set.st_start_row, config_set.st_start_col, nb_conf, 1)
    .getValues();

  var mail_conf = {};
  mail_conf["train_column"] = conf_list[CONF_INDEX.train_column][0];
  mail_conf["query"] = conf_list[CONF_INDEX.query][0]; //フィルタ
  mail_conf["search_limit"] = parseInt(
    conf_list[CONF_INDEX.search_limit][0],
    10
  ); //取得スレッド数 ※1
  if (mail_conf.search_limit < 1) {
    mail_conf.search_limit = 1;
  }
  if (mail_conf.search_limit > MAX_THREADS) {
    mail_conf.search_limit = MAX_THREADS;
  }
  mail_conf["ago_days"] = conf_list[CONF_INDEX.ago_days][0]; //過去メール取得日数
  mail_conf["top_msg_only"] = conf_list[CONF_INDEX.top_msg_only][0]; //スレッドの先頭メッセージのみ取得

  // 本文除外設定
  if (
    lastRow < config_set.exc_start_row ||
    lastCol < config_set.exc_start_col
  ) {
    throw new Error("設定シートに問題があります");
  }
  var exc_list = nlc_conf.sheet_conf.sheet
    .getRange(
      config_set.exc_start_row,
      config_set.exc_start_col,
      lastRow - config_set.exc_start_row + 1,
      1
    )
    .getValues();

  var exc_conf = {};
  var re_list = [];
  for (var i = 0; i < exc_list.length; i += 1) {
    var exc_re = exc_list[i][0];
    if (exc_re !== "") {
      re_list.push(exc_re);
    }
  }
  exc_conf["re_list"] = re_list;

  RUNTIME_CONFIG.sheet_conf = nlc_conf.sheet_conf;
  RUNTIME_CONFIG.notif_conf = nlc_conf.notif_conf;
  RUNTIME_CONFIG.mail_conf = mail_conf;
  RUNTIME_CONFIG.exc_conf = exc_conf;

  return {
    self_ss: nlc_conf.self_ss,
    ss_id: nlc_conf.ss_id,
    sheet_conf: nlc_conf.sheet_conf,
    notif_conf: nlc_conf.notif_conf,
    mail_conf: mail_conf,
    exc_conf: exc_conf
  };
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * @typedef {Object} MailSet
 * @property {String} ss_id SS_ID
 * @property {String} ws_name シート名
 * @property {String} start_col 開始列
 * @property {String} start_row 開始行
 * @property {String} query クエリ
 * @property {String} search_limit 最大スレッド数
 * @property {String} exc_res 除外設定
 * @property {String} msgs メッセージ
 */
/**
 * 最新
 * @param       {MailSet} mail_set メール設定
 * @return      {Date} 最新日付
 * @throws      {Error} データシートが不明です
 */
function MAILAPP_get_newest(mail_set) {
  NLCAPP_log_debug({ record: ["MAILAPP_get_newest", "START"] });

  var lastRow = mail_set.sheet.getLastRow();
  var lastCol = mail_set.sheet.getLastColumn();

  if (lastRow < mail_set.start_row || lastCol < mail_set.start_col) {
    return new Date();
  }

  var entries = [];
  if (lastRow >= mail_set.start_row) {
    entries = mail_set.sheet
      .getRange(
        mail_set.start_row,
        mail_set.start_col + MAIL_FIELDS.DATE,
        lastRow - mail_set.start_row + 1,
        1
      )
      .getValues();

    var from_max = entries.map(function (flds) {
      if (flds[0] === "") {
        return 0;
      }
      return flds[0];
    });
    // 該当なしの場合は当日のみ
    var max_date = Math.max.apply(null, from_max);
    if (max_date === 0) {
      return new Date();
    }
    return new Date(max_date);
  }
  return new Date();
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * 本文を正規表現で除外する
 * @param       {MailSet} p_mail_set 設定情報
 * @return      {MailSet} 編集結果
 */
function MAILAPP_trim_exc(p_mail_set) {
  NLCAPP_log_debug({ record: ["MAILAPP_trim_exc", "START"] });

  var mail_set = p_mail_set;

  for (var i = 0; i < mail_set.msgs.length; i += 1) {
    var msg = mail_set.msgs[i];
    var buf = msg[MAIL_FIELDS.BODY];

    for (var j = 0; j < mail_set.exc_res.length; j += 1) {
      var regexp = new RegExp(mail_set.exc_res[j], "gm");

      buf = buf.replace(regexp, "");
    }

    mail_set.msgs[i][MAIL_FIELDS.NORM_BODY] = buf.trim();
  }

  return mail_set;
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * シートのメールデータを更新する
 * @param       {Object} p_params params
 * @param       {MailSet} mail_set メール情報
 */
function MAILAPP_update_data(p_params) {
  NLCAPP_log_debug({ record: ["MAILAPP_update_data", "START"] });

  var lastRow = p_params.sheet.getLastRow();

  var nb_fields = Object.keys(MAIL_FIELDS).length;

  var entries = [];
  if (lastRow >= p_params.mail_set.start_row) {
    entries = p_params.sheet
      .getRange(
        p_params.mail_set.start_row,
        p_params.mail_set.start_col,
        lastRow - p_params.mail_set.start_row + 1,
        nb_fields
      )
      .getValues();
  } else {
    lastRow = p_params.mail_set.start_row - 1;
  }

  var mail_set = p_params.mail_set;

  var row_cnt = 0;
  for (var i = 0; i < p_params.mail_set.msgs.length; i += 1) {
    var isMatch = 0;
    for (var j = 0; j < entries.length; j += 1) {
      if (
        entries[j][MAIL_FIELDS.ID] === p_params.mail_set.msgs[i][MAIL_FIELDS.ID]
      ) {
        isMatch = 1;
        break;
      }
    }

    if (isMatch === 0) {
      if (
        p_params.mail_set.msgs[i][MAIL_FIELDS.BODY].length > BODY_LENGTH_LIMIT
      ) {
        mail_set.msgs[i][MAIL_FIELDS.BODY] = p_params.mail_set.msgs[i][
          MAIL_FIELDS.BODY
        ].substring(0, BODY_LENGTH_LIMIT);
      }

      if (
        p_params.mail_set.msgs[i][MAIL_FIELDS.NORM_BODY].length >
        BODY_LENGTH_LIMIT
      ) {
        mail_set.msgs[i][MAIL_FIELDS.NORM_BODY] = p_params.mail_set.msgs[i][
          MAIL_FIELDS.NORM_BODY
        ].substring(0, BODY_LENGTH_LIMIT);
      }

      var record = [
        mail_set.msgs[i][MAIL_FIELDS.ID],
        mail_set.msgs[i][MAIL_FIELDS.DATE],
        mail_set.msgs[i][MAIL_FIELDS.SUBJECT],
        mail_set.msgs[i][MAIL_FIELDS.FROM],
        mail_set.msgs[i][MAIL_FIELDS.TO],
        mail_set.msgs[i][MAIL_FIELDS.CC],
        mail_set.msgs[i][MAIL_FIELDS.BODY],
        mail_set.msgs[i][MAIL_FIELDS.NORM_BODY]
      ];

      p_params.sheet
        .getRange(lastRow + row_cnt + 1, mail_set.start_col, 1, record.length)
        .setValues([record]);
      row_cnt += 1;
    }
  }
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * 取得メールをシートに配置する
 * @throws {Error} 過去分取得日数が不正です
 */
function MAILAPP_load_messages() { // eslint-disable-line no-unused-vars

  NLCAPP_log_debug({ record: ["MAILAPP_load_messages", "START"] });

  var conf = MAILAPP_load_config(CONFIG_SET);

  var SELF_SS = SpreadsheetApp.getActiveSpreadsheet();
  var SS_ID = SELF_SS.getId();

  var sheet = SELF_SS.getSheetByName(conf.sheet_conf.ws_name);
  if (sheet === null) {
    throw new Error("データシートが不明です");
  }

  var SS_UI;
  try {
    SS_UI = SpreadsheetApp.getUi();
  } catch (e) {
    SS_UI = null;
  }

  if (!RUNTIME_OPTION.UI_DISABLE || RUNTIME_OPTION.UI_DISABLE === false) {
    if (SS_UI !== null) {
      var res = GASLIB_Dialog_open(
        "取得",
        "メールの取得を開始します。よろしいですか？",
        SS_UI.ButtonSet.OK_CANCEL
      );
      if (res === SS_UI.Button.CANCEL) {
        GASLIB_Dialog_open(
          "取得",
          "メールの取得を中止しました。",
          SS_UI.ButtonSet.OK
        );
        return;
      }
      var msg = "メールの取得を開始しました。";
      GASLIB_Dialog_open("分類", msg, SS_UI.ButtonSet.OK);
    }
  }

  var mail_set = {
    SELF_SS: SELF_SS,
    SS_ID: SS_ID,
    sheet: sheet,
    ws_name: conf.sheet_conf.ws_name,
    start_col: conf.sheet_conf.start_col,
    start_row: conf.sheet_conf.start_row,
    query: conf.mail_conf.query,
    search_limit: conf.mail_conf.search_limit,
    exc_res: conf.exc_conf.re_list,
    top_msg_only: conf.sheet_conf.top_msg_only,
    msgs: []
  };

  NLCAPP_log_debug({ record: ["ago_days", conf.mail_conf.ago_days] });
  var from_date;
  if (conf.mail_conf.ago_days === 0) {
    from_date = MAILAPP_get_newest(mail_set);
    from_date.setHours(0, 0, 0, 0);
  } else if (conf.mail_conf.ago_days > 0) {
    from_date = new Date();
    from_date.setHours(0, 0, 0, 0);
    from_date.setDate(from_date.getDate() - conf.mail_conf.ago_days);
  } else {
    throw new Error("過去メール取得日数が不正です");
  }
  NLCAPP_log_debug({ record: ["from_date", from_date] });

  mail_set.from_date = from_date;

  var msgs = MAILAPP_get_messages(mail_set);

  mail_set.msgs = msgs;

  mail_set = MAILAPP_trim_exc(mail_set);

  MAILAPP_update_data({
    sheet: sheet,
    mail_set: mail_set
  });
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * [MAILAPP_create_training_data description]
 * @param       {[type]} p_params [description]
 * @param       {TrainSet} train_set      学習情報
 * @constructor
 */
function MAILAPP_create_training_data(p_params) {
  NLCAPP_log_debug({ record: ["MAILAPP_create_training_data", "START"] });

  var lastRow = p_params.sheet.getLastRow();
  var lastCol = p_params.sheet.getLastColumn();

  var entries;
  if (
    lastRow < p_params.train_set.start_row ||
    lastCol < p_params.train_set.start_col ||
    lastCol < p_params.train_set.class_col
  ) {
    entries = [];
  } else {
    entries = p_params.sheet
      .getRange(
        p_params.train_set.start_row,
        p_params.train_set.start_col,
        lastRow - p_params.train_set.start_row + 1,
        lastCol - p_params.train_set.start_col
      )
      .setNumberFormat("@")
      .getValues();
  }

  var train_buf = [];
  var row_cnt = 0;
  var csvString = "";

  for (var cnt = 0; cnt < entries.length; cnt += 1) {
    var class_name = GASLIB_Text_normalize(
      String(
        entries[cnt][
          p_params.train_set.class_col - p_params.train_set.start_col
        ]
      )
    ).trim();
    if (class_name === "") continue;

    var subject_text = GASLIB_Text_normalize(
      String(entries[cnt][MAIL_FIELDS.SUBJECT])
    ).trim();
    var body_text = GASLIB_Text_normalize(
      String(entries[cnt][MAIL_FIELDS.NORM_BODY])
    ).trim();

    var train_text;
    if (p_params.train_set.train_column === TRAIN_COLUMN.SUBJECT) {
      train_text = subject_text;
    } else if (p_params.train_set.train_column === TRAIN_COLUMN.BODY) {
      train_text = body_text;
    } else if (p_params.train_set.train_column === TRAIN_COLUMN.BOTH) {
      if (subject_text === "") {
        train_text = body_text;
      } else if (body_text === "") {
        train_text = subject_text;
      } else {
        train_text = subject_text + " " + body_text;
      }
    } else {
      throw new Error("学習・分類対象が不正です");
    }

    if (train_text.length === 0) continue;

    if (train_text.length > 1024) {
      train_text = train_text.substring(0, 1024);
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

  for (var tcnt = 0; tcnt < train_data.length; tcnt += 1) {
    csvString =
      csvString +
      '"' +
      train_data[tcnt].text +
      '","' +
      train_data[tcnt].class +
      '"' +
      "\r\n";
  }

  return {
    row_cnt: row_cnt,
    csvString: csvString
  };
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * 対象分類器を学習する
 * @param       {Integer} clf_no 分類器番号
 */
/*
function MAILUTIL_train_set(clf_no) {

  Logger.log("### MAILUTIL_train_set", clf_no);

  var CREDS = NLCAPP_load_creds();

  var conf = MAILUTIL_load_config(CONFIG_SET);

  var train_set = {
    ss_id: SS_ID,
    ws_name: conf.sheet_conf.ws_name,
    start_row: conf.sheet_conf.start_row,
    start_col: conf.sheet_conf.start_col,
    end_row: -1,
    train_column: conf.mail_conf.train_column,
    text_col: conf.sheet_conf.train_column,
    class_col: conf.sheet_conf.intent_col[clf_no - 1],
    clf_no: clf_no,
    clf_name: CLFNAME_PREFIX + clf_no,
  };

  var train_result = MAILUTIL_train(train_set, CREDS.username, CREDS.password);

  var log_set = {
    ss_id: SS_ID,
    ws_name: conf.sheet_conf.log_ws,
    start_col: CONFIG_SET.log_start_col,
    start_row: CONFIG_SET.log_start_row,
  };

  NLCAPP_log_train(log_set, train_set, train_result);

}
*/
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * [MAILAPP_init_train description]
 * @param {Object} p_params params
 * @return {Object} conf
 */
function MAILAPP_init_train(p_params) {
  NLCAPP_log_debug({ record: ["MAILAPP_init_train", "START"] });

  var creds = NLCAPP_load_creds();

  var conf = MAILAPP_load_config(CONFIG_SET);

  var SELF_SS = SpreadsheetApp.getActiveSpreadsheet();

  var sheet = SELF_SS.getSheetByName(conf.sheet_conf.ws_name);
  if (sheet === null) {
    throw new Error("データシートが不明です");
  }

  var SS_UI;
  try {
    SS_UI = SpreadsheetApp.getUi();
  } catch (e) {
    SS_UI = null;
  }

  if (!RUNTIME_OPTION.UI_DISABLE || RUNTIME_OPTION.UI_DISABLE === false) {
    if (SS_UI != null) {
      var res = GASLIB_Dialog_open(
        "学習",
        "学習を開始します。よろしいですか？",
        SS_UI.ButtonSet.OK_CANCEL
      );
      if (res === SS_UI.Button.CANCEL) {
        GASLIB_Dialog_open("学習", "学習を中止しました。", SS_UI.ButtonSet.OK);
        return {};
      }

      var msg =
        "学習を開始しました。ログは「" +
        conf.sheet_conf.log_ws +
        "」シートをご参照ください。";
      msg +=
        "\nステータスは「" + CONFIG_SET.ws_name + "」シートをご参照ください。";
      GASLIB_Dialog_open("学習", msg, SS_UI.ButtonSet.OK);
    }
  }

  var log_sheet = conf.self_ss.getSheetByName(conf.sheet_conf.log_ws);
  if (log_sheet === null) {
    log_sheet = conf.self_ss.insertSheet(conf.sheet_conf.log_ws);
  }

  var log_set = {
    sheet: log_sheet,
    ws_name: conf.sheet_conf.log_ws,
    start_col: p_params.CONFIG_SET.log_start_col,
    start_row: p_params.CONFIG_SET.log_start_row
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
        error_desc: clfs.body.error
      }
    });
    throw new Error("分類器のエラーです。ログを確認してください");
  }

  return {
    nlc: nlc,
    clfs: clfs,
    conf: conf,
    sheet_conf: conf.sheet_conf,
    mail_conf: conf.mail_conf,
    sheet: sheet,
    log_set: log_set
  };
}

// ----------------------------------------------------------------------------
/**
 * 全分類器を学習
 */
function MAILAPP_train_all() { // eslint-disable-line no-unused-vars

  NLCAPP_log_debug({ record: ["MAILAPP_train_all", "START"] });

  var conf = MAILAPP_init_train({
    CONFIG_SET: CONFIG_SET
  });

  NLCAPP_log_debug({ record: ["MAILAPP_train_all", conf.mail_conf] });

  for (var clf_no = 1; clf_no <= NB_CLFS; clf_no += 1) {
    var train_set = {
      ss_id: conf.SS_ID,
      ws_name: conf.sheet_conf.ws_name,
      start_row: conf.sheet_conf.start_row,
      start_col: conf.sheet_conf.start_col,
      end_row: -1,
      train_column: conf.mail_conf.train_column,
      text_col: conf.mail_conf.train_column,
      class_col: conf.sheet_conf.intent_col[clf_no - 1],
      clf_no: clf_no,
      clf_name: CLFNAME_PREFIX + clf_no
    };

    var training_data = MAILAPP_create_training_data({
      sheet: conf.sheet,
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

  NLCAPP_log_debug({ record: ["MAILAPP_train_all", "TRIGGER SET"] });
  GASLIB_Trigger_set("NLCAPP_exec_check_clfs", 1);
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * [MAILUTIL_create_classify_data description]
 * @param {Object} p_params params
 * @return {Object} classify_data
 */
function MAILAPP_create_classify_data(p_params) {
  NLCAPP_log_debug({ record: ["MAILAPP_create_classify_data", "START"] });

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
    var subject_text = GASLIB_Text_normalize(
      String(entries[cnt][MAIL_FIELDS.SUBJECT])
    ).trim();
    var body_text = GASLIB_Text_normalize(
      String(entries[cnt][MAIL_FIELDS.NORM_BODY])
    ).trim();

    var test_text;
    if (p_params.mail_conf.train_column === TRAIN_COLUMN.SUBJECT) {
      test_text = subject_text;
    } else if (p_params.mail_conf.train_column === TRAIN_COLUMN.BODY) {
      test_text = body_text;
    } else if (p_params.mail_conf.train_column === TRAIN_COLUMN.BOTH) {
      if (subject_text === "") {
        test_text = body_text;
      } else if (body_text === "") {
        test_text = subject_text;
      } else {
        test_text = subject_text + " " + body_text;
      }
    } else {
      throw new Error("学習・分類対象が不正です");
    }

    if (test_text.length === 0) continue;

    if (test_text.length > NLCLIB_MAX_TRAIN_STRINGS) {
      test_text = test_text.substring(0, NLCLIB_MAX_TRAIN_STRINGS);
    }

    // 分類結果のチェック
    var flags = [1, 1, 1];
    for (var j = 0; j < NB_CLFS; j += 1) {
      var clf_no = j;
      var result_text;
      var result_col = p_params.sheet_conf.result_col[clf_no];
      if (lastCol < p_params.sheet_conf.result_col[clf_no]) {
        result_text = "";
      } else {
        result_text = entries[cnt][result_col - p_params.sheet_conf.start_col];
      }

      if (result_text !== "" && CONFIG_SET.result_override !== true) {
        flags[j] = 0;
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
 * [MAILUTIL_init_classify description]
 * @return {Object} conf
 */
function MAILAPP_init_classify() {
  NLCAPP_log_debug({ record: ["MAILAPP_init_classify", "START"] });

  var creds = NLCAPP_load_creds();
  var conf = MAILAPP_load_config(CONFIG_SET);

  var self_ss = SpreadsheetApp.getActiveSpreadsheet();
  var ss_id = self_ss.getId();

  var data_sheet = self_ss.getSheetByName(conf.sheet_conf.ws_name);
  if (data_sheet === null) {
    throw new Error("データシートが不明です");
  }

  var ss_ui;
  try {
    ss_ui = SpreadsheetApp.getUi();
  } catch (e) {
    ss_ui = null;
  }

  if (!RUNTIME_OPTION.UI_DISABLE || RUNTIME_OPTION.UI_DISABLE === false) {
    if (ss_ui !== null) {
      var res = GASLIB_Dialog_open(
        "分類",
        "分類を開始します。よろしいですか？",
        ss_ui.ButtonSet.OK_CANCEL
      );
      if (res === ss_ui.Button.CANCEL) {
        GASLIB_Dialog_open("分類", "分類を中止しました。", ss_ui.ButtonSet.OK);
        return {};
      }
      var msg =
        "分類を開始しました。ログは「" +
        conf.sheet_conf.log_ws +
        "」シートをご参照ください。";
      GASLIB_Dialog_open("分類", msg, ss_ui.ButtonSet.OK);
    }
  }

  var notif_conf = {
    self_ss: self_ss,
    ss_id: ss_id,
    ws_name: conf.notif_conf.ws_name,
    start_col: CONFIG_SET.notif_start_col,
    start_row: CONFIG_SET.notif_start_row
  };

  var rules = NLCAPP_load_notif_rules(notif_conf);

  var notif_set = {
    rules: rules,
    from: CONFIG_SET.notif_from,
    sender: CONFIG_SET.notif_sender,
    result_cols: conf.sheet_conf.result_col
  };

  var log_sheet = self_ss.getSheetByName(conf.sheet_conf.log_ws);
  if (log_sheet === null) {
    log_sheet = self_ss.insertSheet(conf.sheet_conf.log_ws);
  }

  var log_set = {
    sheet: log_sheet,
    ws_name: conf.sheet_conf.log_ws,
    start_col: CONFIG_SET.log_start_col,
    start_row: CONFIG_SET.log_start_row
  };

  var test_set = {
    sheet: data_sheet,
    ws_name: conf.sheet_conf.ws_name,
    start_col: conf.sheet_conf.start_col,
    start_row: conf.sheet_conf.start_row,
    end_row: -1,
    text_col: conf.mail_conf.train_column,
    notif_set: notif_set,
    notif_opt: conf.notif_conf.option
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
    mail_conf: conf.mail_conf,
    test_set: test_set,
    log_set: log_set,
    notif_set: notif_set
  };
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * 全分類器で分類
 * @throws      {Error}  学習・分類対象が不正です
 * @throws      {Error}  データシートが不明です
 */
function MAILAPP_classify_all() { // eslint-disable-line no-unused-vars

  NLCAPP_log_debug({ record: ["MAILAPP_classify_all", "START"] });

  var conf = MAILAPP_init_classify();

  // 分類データの作成
  var classify_data = MAILAPP_create_classify_data(conf);

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
 * @type {String[]} MailData
 * ID, date, subject, from, to, cc, body
 */
/**
 * メールを取得する
 * @param       {MailSet} mail_set 設定情報
 * @param       {String} mail_set.query query
 * @param       {String} mail_set.search_limit search_limit
 * @return      {MailData} メールデータ
 */
function MAILAPP_get_messages(mail_set) {
  NLCAPP_log_debug({ record: ["MAILAPP_get_messages", "START"] });

  var result = [];
  var threads = GmailApp.search(mail_set.query, 0, mail_set.search_limit);

  NLCAPP_log_debug({ record: ["threads", threads.length] });

  for (var i = 0; i < threads.length; i += 1) {
    var thread = threads[i];
    var msgs = thread.getMessages();

    var nb_msgs = 1;
    if (mail_set.top_msg_only === "Off") {
      nb_msgs = msgs.length;
    }

    NLCAPP_log_debug({ record: ["nb_msgs", nb_msgs] });

    for (var j = 0; j < nb_msgs; j += 1) {
      var msg = msgs[j];

      var res_msg = [
        msg.getId(),
        msg.getDate(),
        GASLIB_Text_escape_formula(msg.getSubject()),
        msg.getFrom(),
        msg.getTo(),
        msg.getCc(),
        GASLIB_Text_escape_formula(msg.getPlainBody())
      ];

      NLCAPP_log_debug({
        record: ["MAIL_FIELDS.DATE", res_msg[MAIL_FIELDS.DATE]]
      });
      if (res_msg[MAIL_FIELDS.DATE] >= mail_set.from_date) {
        result.push(res_msg);
      }
    }
  }

  return result;
}
// ----------------------------------------------------------------------------
