// The MIT License (MIT)
//
// Copyright (c) 2017 SoftBank Corp.
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
// ----------------------------------------------------------------------------
// グローバル

/* globals RUNTIME_CONFIG */
/* globals RUNTIME_OPTION */

/* globals GASLIB_RSS_get_feeds */
/* globals GASLIB_Trigger_set */

/* globals NLCLIB_MAX_TRAIN_STRINGS */

/* globals NLCAPP_load_config */
/* globals NLCAPP_log_debug */
/* globals NLCAPP_create_instance */
/* globals NLCAPP_list_classifiers */
/* globals NLCAPP_train_common */
/* globals NLCAPP_get_classifiers */
/* globals NLCAPP_classify_base */
/* globals NLCAPP_log_classify_results */

/**
 * 分類器数
 * @type {Integer}
 */
var NB_CLFS = 3; // eslint-disable-line no-unused-vars

/**
 * 設定シートフィールドインデックス
 * @type {Object}
 * @property {Integer} ws_name      シート名
 * @property {Integer} start_col    定義開始列
 * @property {Integer} start_row    定義開始行
 * @property {Integer} train_column 学習テキスト選択列
 * @property {Integer} intent1_col  インテント列1
 * @property {Integer} result1_col  分類結果列1
 * @property {Integer} resconf1_col 確信度列1
 * @property {Integer} restime1_col 分類日時列1
 * @property {Integer} intent2_col  インテント列2
 * @property {Integer} result2_col  分類結果列2
 * @property {Integer} resconf2_col 確信度列2
 * @property {Integer} restime2_col 分類日時列2
 * @property {Integer} intent3_col  インテント列3
 * @property {Integer} result3_col  分類結果列3
 * @property {Integer} resconf3_col 確信度列3
 * @property {Integer} restime3_col 分類日時列3
 * @property {Integer} log_ws       ログシート名
 * @property {Integer} notif_opt    通知オプション
 * @property {Integer} notif_ws     通知設定シート名
 */
var CONF_INDEX = {
  // eslint-disable-line no-unused-vars
  ws_name: 0,
  start_col: 1,
  start_row: 2,
  train_column: 3,
  intent1_col: 4,
  result1_col: 5,
  resconf1_col: 6,
  restime1_col: 7,
  intent2_col: 8,
  result2_col: 9,
  resconf2_col: 10,
  restime2_col: 11,
  intent3_col: 12,
  result3_col: 13,
  resconf3_col: 14,
  restime3_col: 15,
  log_ws: 16,
  notif_opt: 17,
  notif_ws: 18
};

/**
 * バインドされているスプレッドシートオブジェクト
 * @type {Spreadsheet}}
 */
var self_ss = SpreadsheetApp.getActiveSpreadsheet();

/**
 * バインドされているスプレッドシートのID
 * @type {String}
 */
var ss_id = self_ss.getId();

/**
 * 設定メタデータ
 * @type {ConfigSet} CONFIG_SET
 * @property {String} ss_id           スプレッドシートID
 * @property {String} ws_name         設定シート名
 * @property {Integer} st_start_row    設定シート定義開始行
 * @property {Integer} st_start_col    設定シート定義開始列
 * @property {Integer} notif_start_row 通知設定定義開始行
 * @property {Integer} notif_start_col 通知設定定義開始列
 * @property {Boolean} result_override 分類結果上書きオプション
 * @property {Integer} clfs_start_col  分類器表示開始列
 * @property {Integer} clfs_start_row  分類器表示開始行
 * @property {Integer} log_start_col   ログ開始列
 * @property {Integer} log_start_row   ログ開始行
 */
var CONFIG_SET = {
  // eslint-disable-line no-unused-vars
  ss_id: ss_id,
  ws_name: "設定",
  st_start_row: 2,
  st_start_col: 2,
  rss_start_row: 3,
  rss_start_col: 5,
  notif_start_row: 2,
  notif_start_col: 1,
  result_override: false,
  clfs_start_col: 9,
  clfs_start_row: 3,
  log_start_col: 1,
  log_start_row: 2
};
// ----------------------------------------------------------------------------

// グローバル変数
/* globals CLFNAME_PREFIX */
/* globals GASLIB_Text_normalize */
/* globals NLCAPP_load_creds */
/* globals GASLIB_Dialog_open */
/* globals NLCAPP_log_train */
/* globals NLCAPP_load_notif_rules */
/* globals NLCAPP_escape_formula */

/**
 * RSSデータフィールドインデックス
 * @type {Object}
 */
var RSS_FIELDS = {
  NAME: 0,
  TITLE: 1,
  URL: 2,
  CREATED: 3,
  SUMMARY: 4
};

/**
 * フィード取得用シート名
 * @type {String}
 */
var FEED_WS_NAME = "RSS_WORK";

/**
 * 取得フィード項目名
 * @type {String[]}
 */
var RSS_NAMES = ["title", "url", "created", "summary"];

/**
 * 学習対象選択オプション
 * @type {Object}
 */
var TRAIN_COLUMN = {
  TITLE: "タイトルのみ",
  SUMMARY: "サマリのみ",
  BOTH: "タイトル・サマリ両方"
};
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * 設定情報を取得する
 * @param       {Object} config_set メタデータ
 * @return      {Object} 設定情報
 * @throws      {Error} 設定シートが不明です
 * @throws      {Error} 設定シートに問題があります
 */
function RSSAPP_load_config(config_set) {
  var nlc_conf = NLCAPP_load_config(config_set);

  var nb_conf = Object.keys(CONF_INDEX).length;

  var lastCol = nlc_conf.sheet_conf.sheet.getLastColumn();
  var lastRow = nlc_conf.sheet_conf.sheet.getLastRow();
  if (lastRow < config_set.st_start_row + nb_conf - 1) {
    throw new Error("設定シートに問題があります");
  }

  var conf_list = nlc_conf.sheet_conf.sheet
    .getRange(config_set.st_start_row, config_set.st_start_col, nb_conf, 1)
    .getValues();

  var rss_conf = {};
  rss_conf["train_column"] = conf_list[CONF_INDEX.train_column][0];

  if (
    lastRow < config_set.rss_start_row ||
    lastCol < config_set.rss_start_col
  ) {
    throw new Error("設定シートに問題があります");
  }

  var rss_list = nlc_conf.sheet_conf.sheet
    .getRange(
      config_set.rss_start_row,
      config_set.rss_start_col,
      lastRow - config_set.rss_start_row + 1,
      2
    )
    .getValues();

  var rss_srcs = [];
  for (var i = 0; i < rss_list.length; i += 1) {
    var rss_name = rss_list[i][0];
    var rss_url = rss_list[i][1];
    if (rss_name !== "" && rss_url !== "") {
      rss_srcs.push({
        name: rss_name,
        url: rss_url
      });
    }
  }

  RUNTIME_CONFIG.sheet_conf = nlc_conf.sheet_conf;
  RUNTIME_CONFIG.rss_conf = rss_conf;

  return {
    self_ss: nlc_conf.self_ss,
    ss_id: nlc_conf.ss_id,
    sheet_conf: nlc_conf.sheet_conf,
    notif_conf: nlc_conf.notif_conf,
    rss_conf: rss_conf,
    rss_srcs: rss_srcs
  };
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * RSSフィードを取得する
 * @param       {Object} feed_set 設定情報
 * @return      {Object} フィード
 */
function RSSAPP_get_feeds(feed_set) { // eslint-disable-line no-unused-vars

  var sheet = self_ss.getSheetByName(feed_set.ws_name);

  var formulastring = "";
  formulastring = '=Importfeed( "' + feed_set.url + '", "items", TRUE )';
  sheet.getRange("A1").setFormula(formulastring);

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  var rss_data = sheet.getRange(1, 1, lastRow, lastCol).getValues();

  var indices = {};
  indices.title = rss_data[0].indexOf("Title");
  indices.url = rss_data[0].indexOf("URL");
  indices.created = rss_data[0].indexOf("Date Created");
  indices.summary = rss_data[0].indexOf("Summary");

  var result = [];
  for (var cnt = 1; cnt < rss_data.length; cnt += 1) {
    var localDate;
    if (indices.created !== -1) {
      var dateObj = new Date(rss_data[cnt][indices.created]);
      localDate = Utilities.formatDate(dateObj, "JST", "yyyy_MMdd_HHmmss");
    } else {
      localDate = "";
    }

    result.push({
      name: feed_set.name,
      title:
        indices.title !== -1
          ? NLCAPP_escape_formula(rss_data[cnt][indices.title])
          : "",
      url: indices.url !== -1 ? rss_data[cnt][indices.url] : "",
      created: localDate,
      summary:
        indices.summary !== -1
          ? NLCAPP_escape_formula(rss_data[cnt][indices.summary].trim())
          : ""
    });
  }

  return result;
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * ワークエリアを消去する
 * @param       {Object} feed_set 取得設定
 */
function RSSAPP_clear_work(feed_set) { // eslint-disable-line no-unused-vars

  var sheet = self_ss.getSheetByName(feed_set.ws_name);
  if (sheet === null) {
    sheet = self_ss.insertSheet(feed_set.ws_name);
  }

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastRow < 1 && lastCol < 1) {
    return;
  }

  sheet.getRange(1, 1, lastRow, lastCol).clear();
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * RSSフィードを追記する
 * @param       {Object} data_set 配置情報
 */
function RSSAPP_update_data(data_set) {
  var sheet = self_ss.getSheetByName(data_set.ws_name);
  if (sheet === null) {
    sheet = self_ss.insertSheet(data_set.ws_name);
  }

  var lastRow = sheet.getLastRow();

  var nb_fields = Object.keys(RSS_FIELDS).length;
  var entries = [];
  if (lastRow >= data_set.start_row) {
    entries = sheet
      .getRange(
        data_set.start_row,
        data_set.start_col,
        lastRow - data_set.start_row + 1,
        nb_fields
      )
      .getValues();
  } else {
    lastRow = data_set.start_row - 1;
  }

  var row_cnt = 0;
  for (var i = 0; i < data_set.feeds.length; i += 1) {
    var isMatch = 0;

    for (var j = 0; j < entries.length; j += 1) {
      if (
        entries[j][RSS_FIELDS.NAME] === data_set.feeds[i].name &&
        entries[j][RSS_FIELDS.TITLE] === data_set.feeds[i].title &&
        entries[j][RSS_FIELDS.URL] === data_set.feeds[i].url &&
        entries[j][RSS_FIELDS.CREATED] === data_set.feeds[i].created
      ) {
        isMatch = 1;
        break;
      }
    }

    if (isMatch === 0) {
      var record = [
        data_set.feeds[i].name,
        data_set.feeds[i].title,
        data_set.feeds[i].url,
        data_set.feeds[i].created,
        data_set.feeds[i].summary
      ];

      sheet
        .getRange(lastRow + row_cnt + 1, data_set.start_col, 1, nb_fields)
        .setValues([record]);
      row_cnt += 1;
    }
  }
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * RSSフォードを取得してシートに追記する
 * @param       {Object} feed_set 取得設定
 * @param       {Object} rss_set  配置設定
 */
function RSSAPP_load_rss(feed_set, rss_set) {
  var rss_feeds = GASLIB_RSS_get_feeds({
    self_ss: feed_set.self_ss,
    ws_name: feed_set.ws_name,
    feed_name: feed_set.name,
    feed_url: feed_set.url
  });

  if (rss_feeds === null) {
    throw new Error(feed_set.name + "のフィードは取得できませんでした");
  }

  var data_set = {
    ss_id: ss_id,
    ws_name: rss_set.ws_name,
    start_col: rss_set.start_col,
    start_row: rss_set.start_row,
    rss_name: feed_set.name,
    feeds: rss_feeds
  };

  RSSAPP_update_data(data_set);
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * 全てのRSSフィードを取得してシートに追記する
 */
function RSSAPP_crawl() { // eslint-disable-line no-unused-vars

  Logger.log("### RSSAPP_crawl");

  var conf = RSSAPP_load_config(CONFIG_SET);

  var ss_ui;
  try {
    ss_ui = SpreadsheetApp.getUi();
  } catch (e) {
    ss_ui = null;
  }

  if (!RUNTIME_OPTION.UI_DISABLE || RUNTIME_OPTION.UI_DISABLE === false) {
    if (ss_ui !== null) {
      var res = GASLIB_Dialog_open(
        "データ取得",
        "データ取得を開始します。よろしいですか？",
        ss_ui.ButtonSet.OK_CANCEL
      );
      if (res === ss_ui.Button.CANCEL) {
        GASLIB_Dialog_open(
          "データ取得",
          "データ取得を中止しました。",
          ss_ui.ButtonSet.OK
        );
        return;
      }
      var msg = "データ取得を開始しました。";
      GASLIB_Dialog_open("データ取得", msg, ss_ui.ButtonSet.OK);
    }
  }

  var rss_set = {
    ss_id: conf.ss_id,
    ws_name: conf.sheet_conf.ws_name,
    start_col: conf.sheet_conf.start_col,
    start_row: conf.sheet_conf.start_row
  };

  for (var cnt = 0; cnt < conf.rss_srcs.length; cnt += 1) {
    var feed_set = {
      self_ss: conf.self_ss,
      ss_id: conf.ss_id,
      ws_name: FEED_WS_NAME,
      name: conf.rss_srcs[cnt].name,
      url: conf.rss_srcs[cnt].url,
      columns: RSS_NAMES
    };

    RSSAPP_load_rss(feed_set, rss_set);
  }
}
// ----------------------------------------------------------------------------

/**
 * [RSSAPP_create_training_data description]
 * @param {Object} p_params params
 * @return {Object} result
 */
function RSSAPP_create_training_data(p_params) {
  NLCAPP_log_debug({ record: ["RSSAPP_create_training_data", "START"] });

  var lastRow = p_params.data_sheet.getLastRow();
  var lastCol = p_params.data_sheet.getLastColumn();

  var entries;
  if (
    lastRow < p_params.train_set.start_row ||
    lastCol < p_params.train_set.start_col
  ) {
    entries = [];
  } else {
    entries = p_params.data_sheet
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

    if (entries[cnt][RSS_FIELDS.NAME].length === 0) continue;
    if (entries[cnt][RSS_FIELDS.NAME] === 0) continue;

    var train_text = "";
    var title_text = GASLIB_Text_normalize(
      String(entries[cnt][RSS_FIELDS.TITLE])
    ).trim();
    var summary_text = GASLIB_Text_normalize(
      String(entries[cnt][RSS_FIELDS.SUMMARY])
    ).trim();

    if (p_params.train_set.train_column === TRAIN_COLUMN.TITLE) {
      train_text = title_text;
    } else if (p_params.train_set.train_column === TRAIN_COLUMN.SUMMARY) {
      train_text = summary_text;
    } else if (p_params.train_set.train_column === TRAIN_COLUMN.BOTH) {
      if (title_text === "") {
        train_text = summary_text;
      } else if (summary_text === "") {
        train_text = title_text;
      } else {
        train_text = title_text + " " + summary_text;
      }
    } else {
      throw new Error("学習・分類対象が不正です");
    }

    if (train_text.length === 0) continue;

    if (train_text.length > 1024) {
      train_text = train_text.substring(0, 1024);
    }

    NLCAPP_log_debug({ record: ["RSSAPP_create_training_data", train_text] });
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

  var LIMIT = 15000;
  var train_data = [];
  if (row_cnt > LIMIT) {
    train_data = train_buf.splice(row_cnt - LIMIT, row_cnt - 1);
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
/**
 * RSSフィードのテキストを学習する
 * @param       {Object} p_params      学習設定
 * @param       {Object} p_params.train_set      学習設定
 * @throws      {Error} データシートが不明です
 * @throws      {Error} 学習・分類対象が不正です
 */
/*
function RSSAPP_train(p_params) {
  NLCAPP_log_debug({ record: ["RSSAPP_train", "START"] });

  for (var clf_no = 1; clf_no <= NB_CLFS; clf_no += 1) {
    var train_set = {
      ss_id: p_params.conf.SS_ID,
      ws_name: p_params.sheet_conf.ws_name,
      start_row: p_params.sheet_conf.start_row,
      start_col: p_params.sheet_conf.start_col,
      end_row: -1,
      train_column: p_params.rss_conf.train_column,
      text_col: p_params.rss_conf.train_column,
      class_col: p_params.sheet_conf.intent_col[clf_no - 1],
      clf_no: clf_no,
      clf_name: CLFNAME_PREFIX + clf_no
    };

    var training_data = RSSAPP_create_training_data({
      data_sheet: p_params.conf.sheet,
      train_set: train_set
    });

    var clf_name = CLFNAME_PREFIX + clf_no;
    var clf_info = NLCAPP_clf_vers({
      clf_list: p_params.clfs.body.classifiers,
      target_name: clf_name
    });
    var new_version = clf_info.max_ver + 1;
    var new_name = clf_name + CLF_SEP + new_version;

    var train_result = {};
    if (training_data.row_cnt === 0) {
      train_result = {
        status: 0,
        description: "学習データなし",
        rows: 0,
        version: new_version
      };

      NLCAPP_log_train({
        log_set: p_params.log_set,
        train_set: train_set,
        train_result: train_result
      });
      continue;
    }

    // 分類器の作成
    var train_params = {
      metadata: {
        name: new_name,
        language: "ja"
      },
      training_data: training_data.csvString
    };
    NLCAPP_log_debug({
      record: [
        "NLC createClassifier",
        "START",
        JSON.stringify(train_params.metadata)
      ]
    });
    var nlc_res = p_params.nlc.createClassifier(train_params);

    if (nlc_res.status === 200) {
      NLCAPP_set_classifier_status({
        sheet: p_params.sheet_conf.sheet,
        clf_no: clf_no,
        classifier_id: nlc_res.body.classifier_id,
        status: nlc_res.body.status
      });

      // バージョンが複数ある場合
      if (clf_info.count >= 2 && nlc_res.status === 200) {
        // 分類器の削除
        var del_params = {
          classifier_id: clf_info.clfs[clf_info.min_ver].classifier_id
        };
        p_params.nlc.deleteClassifier(del_params);
      }
    }

    train_result = {
      status: nlc_res.status,
      nlc: nlc_res,
      rows: training_data.row_cnt,
      version: new_version
    };

    NLCAPP_log_train({
      log_set: p_params.log_set,
      train_set: train_set,
      train_result: train_result
    });
  }
}
*/
// ----------------------------------------------------------------------------

/**
 * [RSSAPP_init_train description]
 * @return {Object} result
 */
function RSSAPP_init_train() {
  var creds = NLCAPP_load_creds();

  var conf = RSSAPP_load_config(CONFIG_SET);

  var sheet = conf.self_ss.getSheetByName(conf.sheet_conf.ws_name);
  if (sheet === null) {
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
    ws_name: conf.sheet_conf.log_ws,
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
    rss_conf: conf.rss_conf,
    sheet: sheet,
    log_set: log_set
  };
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * イベント実行用学習処理
 */
function RSSAPP_train_all() { // eslint-disable-line no-unused-vars

  NLCAPP_log_debug({ record: ["RSSAPP_train_all", "START"] });

  var conf = RSSAPP_init_train();

  var data_sheet = conf.conf.self_ss.getSheetByName(conf.sheet_conf.ws_name);

  for (var clf_no = 1; clf_no <= NB_CLFS; clf_no += 1) {
    var train_set = {
      ws_name: conf.sheet_conf.ws_name,
      start_row: conf.sheet_conf.start_row,
      start_col: conf.sheet_conf.start_col,
      end_row: -1,
      train_column: conf.rss_conf.train_column,
      text_col: conf.rss_conf.train_column,
      class_col: conf.sheet_conf.intent_col[clf_no - 1],
      clf_no: clf_no,
      clf_name: CLFNAME_PREFIX + clf_no
    };

    var training_data = RSSAPP_create_training_data({
      data_sheet: data_sheet,
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

  NLCAPP_log_debug({ record: ["RSSAPP_train_all", "TRIGGER SET"] });
  GASLIB_Trigger_set("NLCAPP_exec_check_clfs", 1);
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * RSSフィードのテキストを分類する
 * @param       {Object} test_set       分類設定
 * @param       {String} creds_username クレデンシャル
 * @param       {String} creds_password クレデンシャル
 * @param       {Boolean} override       上書きオプション
 * @return      {Object} 分類結果
 * @throws      {Error}  データシートが不明です
 * @throws      {Error} 学習・分類対象が不正です
 */
/*
function RSSAPP_classify(test_set, creds_username, creds_password, override) {
  Logger.log("### RSSAPP_classify");

  var clf_id = NLCAPP_select_clf(
    test_set.clf_name,
    creds_username,
    creds_password
  );
  if (clf_id === "") {
    return {
      status: 900,
      description: "classifier does not found"
    };
  }

  var sheet = self_ss.getSheetByName(test_set.ws_name);
  if (sheet === null) {
    throw new Error("データシートが不明です");
  }

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  var entries;
  if (lastRow < test_set.start_row || lastCol < test_set.start_col) {
    entries = [];
  } else {
    entries = sheet
      .getRange(
        test_set.start_row,
        test_set.start_col,
        lastRow - test_set.start_row + 1,
        lastCol - test_set.start_col + 1
      )
      .getValues();
  }

  var hasError = 0;
  var row_cnt = 0;
  var nlc_res;
  var err_res;
  for (var cnt = 0; cnt < entries.length; cnt += 1) {
    var rss_name = entries[cnt][RSS_FIELDS.NAME];
    if (rss_name.length === 0) continue;
    if (rss_name === 0) continue;

    var result_text;
    if (lastCol < test_set.result_col) {
      result_text = "";
    } else {
      result_text = entries[cnt][test_set.result_col - test_set.start_col];
    }
    if (result_text !== "" && override !== true) continue;

    var title_text = GASLIB_Text_normalize(
      String(entries[cnt][RSS_FIELDS.TITLE])
    ).trim();
    var summary_text = GASLIB_Text_normalize(
      String(entries[cnt][RSS_FIELDS.SUMMARY])
    ).trim();

    var test_text;
    if (test_set.train_column === TRAIN_COLUMN.TITLE) {
      test_text = title_text;
    } else if (test_set.train_column === TRAIN_COLUMN.SUMMARY) {
      test_text = summary_text;
    } else if (test_set.train_column === TRAIN_COLUMN.BOTH) {
      if (title_text === "") {
        test_text = summary_text;
      } else if (summary_text === "") {
        test_text = title_text;
      } else {
        test_text = title_text + " " + summary_text;
      }
    } else {
      throw new Error("学習・分類対象が不正です");
    }

    if (test_text.length === 0) continue;

    if (test_text.length > 1024) {
      test_text = test_text.substring(0, 1024);
    }

    nlc_res = NLCAPI_post_classify(
      creds_username,
      creds_password,
      clf_id,
      test_text
    );
    if (nlc_res.status !== 200) {
      hasError = 1;
      err_res = nlc_res;
    }

    var r = nlc_res.body.top_class;
    sheet
      .getRange(test_set.start_row + cnt, test_set.result_col, 1, 1)
      .setNumberFormat("@")
      .setValue(r);

    var t = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss");
    sheet
      .getRange(test_set.start_row + cnt, test_set.restime_col, 1, 1)
      .setValue(t);

    // TODO: confidence

    row_cnt += 1;

    if (test_set.notif_opt === NOTIF_OPT.ON) {
      var record = sheet
        .getRange(test_set.start_row + cnt, 1, 1, lastCol)
        .getValues();

      NLCAPP_check_notify(test_set.notif_set, record[0]);
    }
  }

  if (row_cnt === 0) {
    return {
      status: 0,
      rows: 0
    };
  }
  if (hasError === 1) {
    return {
      status: err_res.status,
      rows: row_cnt,
      nlc: err_res
    };
  }
  return {
    status: nlc_res.status,
    rows: row_cnt,
    nlc: nlc_res
  };
}
*/
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * 特定分類器で分類して結果をログに出力する
 * @param       {Integer} clf_no 分類器番号
 */
/*
function RSSAPP_classify_set(clf_no) {
  Logger.log("### RSSAPP_classify_set " + clf_no);

  var creds = NLCAPP_load_creds();

  var conf = RSSAPP_load_config(CONFIG_SET);

  var notif_conf = {
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

  var test_set = {
    ss_id: ss_id,
    ws_name: conf.sheet_conf.ws_name,
    start_col: conf.sheet_conf.start_col,
    start_row: conf.sheet_conf.start_row,
    end_row: -1,
    text_col: conf.sheet_conf.train_column,
    result_col: conf.sheet_conf.result_col[clf_no - 1],
    restime_col: conf.sheet_conf.restime_col[clf_no - 1],
    clf_name: CLFNAME_PREFIX + String(clf_no),
    notif_set: notif_set,
    notif_opt: conf.notif_conf.option
  };

  var test_result = RSSAPP_classify(test_set, creds.username, creds.password);

  var log_set = {
    ss_id: ss_id,
    ws_name: conf.sheet_conf.log_ws,
    start_col: CONFIG_SET.log_start_col,
    start_row: CONFIG_SET.log_start_row
  };

  NLCAPP_log_classify(log_set, test_set, test_result);
}
*/
// ----------------------------------------------------------------------------

/**
 * [RSSAPP_init_classify description]
 * @return {Object} result
 */
function RSSAPP_init_classify() {
  NLCAPP_log_debug({ record: ["RSSAPP_init_classify", "START"] });

  var creds = NLCAPP_load_creds();

  var conf = RSSAPP_load_config(CONFIG_SET);

  var data_sheet = conf.self_ss.getSheetByName(conf.sheet_conf.ws_name);
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
    self_ss: conf.self_ss,
    ss_id: conf.ss_id,
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

  var log_sheet = conf.self_ss.getSheetByName(conf.sheet_conf.log_ws);
  if (log_sheet === null) {
    log_sheet = conf.self_ss.insertSheet(conf.sheet_conf.log_ws);
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
    text_col: conf.rss_conf.train_column,
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
    rss_conf: conf.rss_conf,
    test_set: test_set,
    log_set: log_set,
    notif_set: notif_set
  };
}

// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * [RSSAPP_create_classify_data description]
 * @param       {[type]} p_params [description]
 * @constructor
 */
function RSSAPP_create_classify_data(p_params) {
  NLCAPP_log_debug({ record: ["RSSAPP_create_classify_data", "START"] });

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
    var rss_name = entries[cnt][RSS_FIELDS.NAME];
    if (rss_name.length === 0) continue;
    if (rss_name === 0) continue;

    var title_text = GASLIB_Text_normalize(
      String(entries[cnt][RSS_FIELDS.TITLE])
    ).trim();
    var summary_text = GASLIB_Text_normalize(
      String(entries[cnt][RSS_FIELDS.SUMMARY])
    ).trim();

    var test_text;
    if (p_params.rss_conf.train_column === TRAIN_COLUMN.TITLE) {
      test_text = title_text;
    } else if (p_params.rss_conf.train_column === TRAIN_COLUMN.SUMMARY) {
      test_text = summary_text;
    } else if (p_params.rss_conf.train_column === TRAIN_COLUMN.BOTH) {
      if (title_text === "") {
        test_text = summary_text;
      } else if (summary_text === "") {
        test_text = title_text;
      } else {
        test_text = title_text + " " + summary_text;
      }
    } else {
      throw new Error("学習・分類対象が不正です");
    }

    if (test_text.length === 0) continue;

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

    NLCAPP_log_debug({ record: ["RSSAPP_create_classify_data", test_text] });
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
 * [RSSAPP_classify_base description]
 */
/*
function RSSAPP_classify_base() {
  var updates = 0;
  var upd_flg = [0, 0, 0];
  for (var j = 0; j < NB_CLFS; j += 1) {
    if (clf_ids[j].status !== "Available") continue;

    var result_text;
    if (lastCol < conf.sheet_conf.result_col[j]) {
      result_text = "";
    } else {
      result_text =
        entries[cnt][conf.sheet_conf.result_col[j] - conf.sheet_conf.start_col];
    }

    if (result_text !== "" && CONFIG_SET.result_override !== true) continue;

    nlc_res = NLCAPI_post_classify(
      creds.username,
      creds.password,
      clf_ids[j].id,
      test_text
    );
    if (nlc_res.status !== 200) {
      hasError = 1;
      err_res = nlc_res;
    } else {
      var r = nlc_res.body.top_class;
      sheet
        .getRange(
          conf.sheet_conf.start_row + cnt,
          conf.sheet_conf.result_col[j],
          1,
          1
        )
        .setNumberFormat("@")
        .setValue(r);
      var c = nlc_res.body.classes[0].confidence;
      sheet
        .getRange(
          conf.sheet_conf.start_row + cnt,
          conf.sheet_conf.resconf_col[j],
          1,
          1
        )
        .setValue(c);
      var t = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss");
      sheet
        .getRange(
          conf.sheet_conf.start_row + cnt,
          conf.sheet_conf.restime_col[j],
          1,
          1
        )
        .setValue(t);
      updates += 1;
      res_rows[j] += 1;
      upd_flg[j] = 1;
    }
  }

  if (test_set.notif_opt === NOTIF_OPT.ON) {
    if (updates > 0) {
      //row_cnt += 1;
      var record = sheet
        .getRange(conf.sheet_conf.start_row + cnt, 1, 1, lastCol)
        .getValues();
      NLCAPP_check_notify(notif_set, record[0], upd_flg);
    }
  }
}
*/
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * イベント実行用分類処理
 * @throws      {Error}  データシートが不明です
 * @throws      {Error} 学習・分類対象が不正です
 */
function RSSAPP_classify_all() { // eslint-disable-line no-unused-vars
  NLCAPP_log_debug({ record: ["RSSAPP_classify_all", "START"] });

  var conf = RSSAPP_init_classify();

  var classify_data = RSSAPP_create_classify_data(conf);

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
 * [RSSAPP_log_classify_results description]
 */
/*
function RSSAPP_log_classify_results() {

  for (var k = 0; k < NB_CLFS; k += 1) {
    if (clf_ids[k].status !== "Available") continue;

    var clfname = CLFNAME_PREFIX + String(k + 1);
    test_set.clf_no = k + 1;
    test_set.clf_name = clfname;
    test_set.result_col = conf.sheet_conf.result_col[k];
    test_set.restime_col = conf.sheet_conf.restime_col[k];

    var test_result;
    if (res_rows[k] === 0) {
      test_result = {
        status: 0,
        rows: 0
      };
      NLCAPP_log_classify(log_set, test_set, test_result);
    } else {
      if (hasError === 1) {
        test_result = {
          status: err_res.status,
          rows: res_rows[k],
          nlc: err_res
        };
        NLCAPP_log_classify(log_set, test_set, test_result);
      } else {
        test_result = {
          status: nlc_res.status,
          rows: res_rows[k],
          nlc: nlc_res
        };
        NLCAPP_log_classify(log_set, test_set, test_result);
      }
    }
  }
}
*/

// ----------------------------------------------------------------------------
// イベント用ラッパー
// ----------------------------------------------------------------------------
// f807209 - 管理対象外Classifierの対応、共通関数の実行抑止
