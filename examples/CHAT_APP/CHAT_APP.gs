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
/* globals CONF_INDEX */
/* globals CONFIG_SET */
/* globals NB_CLFS */
/* globals CLFNAME_PREFIX */
/* globals CLF_SEP */
/* globals RUNTIME_CONFIG */
/* globals RUNTIME_OPTION */
/* globals RUNTIME_STATUS */

/* globals GASLIB_Text_normalize */
/* globals GASLIB_Dialog_open */
/* globals GASLIB_Trigger_set */

/* globals NLCLIB_MAX_TRAIN_STRINGS */
/* globals NLCLIB_MAX_TRAIN_RECORDS */
/* globals NLCAPP_load_creds */
/* globals NLCAPP_load_config */
/* globals NLCAPP_clf_vers */
/* globals NLCAPP_log_train */
/* globals NLCAPP_load_creds */
/* globals NLCAPP_create_instance */
/* globals NLCAPP_list_classifiers */
/* globals NLCAPP_exec_check_clfs */
/* globals NLCAPP_list_classifiers */
/* globals NLCAPP_get_classifiers */
/* globals NLCAPP_train_common */
/* globals NLCAPP_log_debug */

var IS_DEBUG = false;

/**
 * 応答設定フィールドインデックス
 * @type {Object}
 */
var CONV_INDEX = {
  result1: 0,
  resconf1: 1,
  result2: 2,
  resconf2: 3,
  result3: 4,
  resconf3: 5,
  message: 6,
  question: 7
};

/**
 * LINE応答メッセージ用URL
 * @type {String}
 */
var LINE_REPLY_URL = "https://api.line.me/v2/bot/message/reply"; // eslint-disable-line no-unused-vars
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * @typedef {Object} ChatCreds クレデンシャル情報
 * @property {String} username NLCのユーザー名
 * @property {String} password NLCのパスワード
 * @property {String} url      NLCのエンドポイント
 * @property {String} channel_access_token LINEチャネルアクセストークン
 */
/**
 * クレデンシャル情報の取得
 * <p>利用するNLCインスタンスのクレデンシャル情報をスクリプトプロパティから取得する</p>
 * <p>利用するLINEアカウントのアクセストークンをスクリプトプロパティから取得する</p>
 * @return {ChatCreds} クレデンシャル情報
 * @throws {Error}  NLCクレデンシャルが不明です
 * @throws {Error}  LINEクレデンシャルが不明です
 */
function CHATAPP_load_creds() { // eslint-disable-line no-unused-vars

  var scriptProps = PropertiesService.getScriptProperties();

  var creds = {};

  creds["channel_access_token"] = scriptProps.getProperty(
    "CHANNEL_ACCESS_TOKEN"
  );

  if (creds.channel_access_token === null) {
    throw new Error("LINEクレデンシャルが不明です");
  }

  return creds;
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * @typedef {Object} ChatConfig 設定情報
 * @property {SheetConf} sheet_conf データシート設定
 * @property {ConvConf}  conv_conf 応答設定
 */
/**
 * 設定情報の取得
 * @param       {ConfigMeta} config_set メタデータ
 * @return      {ChatConfig} 設定情報
 * @throws      {Error}  設定シートが不明です
 * @throws      {Error}  設定シートに問題があります
 */
function CHATAPP_load_config(config_set) {
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

  var records = nlc_conf.sheet_conf.sheet
    .getRange(config_set.st_start_row, config_set.st_start_col, nb_conf, 1)
    .getValues();

  var i = 0;
  var chat_conf = {
    text_col: parseInt(records[CONF_INDEX.text_col][i], 10),
    start_msg: records[CONF_INDEX.start_msg][i],
    other_msg: records[CONF_INDEX.other_msg][i],
    error_msg: records[CONF_INDEX.error_msg][i],
    avatar_url: records[CONF_INDEX.avatar_url][i],
    giveup_msg: records[CONF_INDEX.giveup_msg][i],
    show_suggests: records[CONF_INDEX.show_suggests][i]
  };

  var conv_conf = {};
  conv_conf["ws_name"] = records[CONF_INDEX.conv_ws][0];

  RUNTIME_CONFIG.sheet_conf = nlc_conf.sheet_conf;
  RUNTIME_CONFIG.conv_conf = conv_conf;

  return {
    self_ss: nlc_conf.self_ss,
    ss_id: nlc_conf.ss_id,
    sheet_conf: nlc_conf.sheet_conf,
    chat_conf: chat_conf,
    conv_conf: conv_conf
  };
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * 応答設定の取得
 * @param       {Object} p_params メタデータ
 * @param       {ConfigMeta} config_set メタデータ
 * @return      {Config} 応答設定
 * @throws      {Error}  応答設定シートに問題があります
 */
function CHATAPP_load_conv_rules(p_params) {
  Logger.log(">>> CHATAPP_load_conv_rules");

  var self_ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = self_ss.getSheetByName(p_params.ws_name);
  if (sheet === null) {
    return [];
  }

  var nb_conf = Object.keys(CONV_INDEX).length;

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastCol < p_params.start_col + nb_conf - 1) {
    throw new Error("応答設定シートに問題があります");
  }

  if (lastRow < p_params.start_row) {
    return [];
  }

  var records = sheet
    .getRange(
      p_params.start_row,
      p_params.start_col,
      lastRow - p_params.start_row + 1,
      nb_conf
    )
    .setNumberFormat("@")
    .getValues();

  var rules = [];
  for (var i = 0; i < records.length; i += 1) {
    rules.push({
      res_int: [
        String(records[i][CONV_INDEX.result1]),
        String(records[i][CONV_INDEX.result2]),
        String(records[i][CONV_INDEX.result3])
      ],
      res_conf: [
        String(records[i][CONV_INDEX.resconf1]),
        String(records[i][CONV_INDEX.resconf2]),
        String(records[i][CONV_INDEX.resconf3])
      ],
      message: records[i][CONV_INDEX.message],
      question: records[i][CONV_INDEX.question]
    });
  }

  Logger.log("<<< CHATAPP_load_conv_rules");

  return rules;
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * 展開用辞書
 * @type {Object} EnvVars
 * @property {String} input 入力テキスト
 * @property {String} date 現在日付
 * @property {String} time 現在時刻
 */
/**
 * 展開用辞書
 * @type {Object} ExpandResult
 * @property {String} code
 * @property {String} text
 */
/**
 * 埋め込みタグ展開
 * @param       {String} p_temp 対象文字列
 * @param       {EnvVars} p_dict 辞書
 * @return      {Object} 展開結果
 */
function CHATAPP_expand_tags(p_temp, p_dict) {
  var xbody = p_temp;
  var buf = "";
  Object.keys(p_dict).forEach(function (key) {
    buf = xbody.replace(
      new RegExp("\\[\\[#" + key + "\\]\\]", "g"),
      p_dict[key]
    );
    xbody = buf;
  });

  xbody.match(new RegExp("\\[\\[#.+\\]\\]", "g"));

  return {
    code: "OK",
    text: xbody
  };
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * メッセージ選択
 * @param       {ConvSet} conv_set    応答設定
 * @param       {Object} res_classes 分類結果
 * @return      {String[]} 応答メッセージ
 */
function CHATAPP_select_message(conv_set, res_classes) {
  NLCAPP_log_debug({ record: ["CHATAPP_select_message", "START"] });

  IS_DEBUG = false;

  var NB_MESSAGES = 1;
  if (conv_set.show_suggests === "ON") {
    NB_MESSAGES = 4;
  }

  // DEBUG
  var row;
  var debug;
  if (IS_DEBUG === true) {
    debug = SpreadsheetApp.getActive().getSheetByName("DEBUG");
    var n = debug.getLastRow();
    if (n >= 2) {
      debug.deleteRows(2, n - 1);
    }
    row = 2;
    res_classes[0].classes.forEach(function (item) {
      var rec = [item.class_name, item.confidence];
      debug.getRange(row, 1, 1, rec.length).setValues([rec]);
      row += 1;
    });
    row = 2;
    res_classes[1].classes.forEach(function (item) {
      var rec = [item.class_name, item.confidence];
      debug.getRange(row, 3, 1, rec.length).setValues([rec]);
      row += 1;
    });
    row = 2;
    res_classes[2].classes.forEach(function (item) {
      var rec = [item.class_name, item.confidence];
      debug.getRange(row, 5, 1, rec.length).setValues([rec]);
      row += 1;
    });
    row = 10;
  }

  if (res_classes[0].classes.length === 0) {
    res_classes[0].classes.push({
      class_name: "",
      confidence: 0.0
    });
  }
  if (res_classes[1].classes.length === 0) {
    res_classes[1].classes.push({
      class_name: "",
      confidence: 0.0
    });
  }
  if (res_classes[2].classes.length === 0) {
    res_classes[2].classes.push({
      class_name: "",
      confidence: 0.0
    });
  }

  var refs = [];
  res_classes[0].classes.forEach(function (item1) {
    res_classes[1].classes.forEach(function (item2) {
      res_classes[2].classes.forEach(function (item3) {
        refs.push({
          names: [item1.class_name, item2.class_name, item3.class_name],
          confs: [item1.confidence, item2.confidence, item3.confidence],
          score: item1.confidence + item2.confidence + item3.confidence
        });
      });
    });
  });

  var sorted = refs.sort(function (elem1, elem2) {
    if (elem1.score > elem2.score) return -1;
    if (elem1.score < elem2.score) return 1;
    return 0;
  });

  var messages = [];
  var match_cnt = 0;
  var rec;
  sorted.forEach(function (ref) {
    if (IS_DEBUG === true) {
      row += 1;
      rec = [
        ref.names[0],
        ref.confs[0],
        ref.names[1],
        ref.confs[1],
        ref.names[2],
        ref.confs[2],
        ref.score
      ];
      debug.getRange(row, 1, 1, rec.length).setValues([rec]);
    }

    if (match_cnt >= NB_MESSAGES) return;

    for (var i = 0; i < conv_set.rules.length; i += 1) {
      if (conv_set.rules[i].checked === true) continue;

      var chk_cnt = 0;
      for (var j = 0; j < NB_CLFS; j += 1) {
        if (conv_set.rules[i].res_int[j] === "") {
          chk_cnt += 1;
        } else {
          if (ref.names[j] === conv_set.rules[i].res_int[j]) {
            if (conv_set.rules[i].res_conf[j] === "") {
              chk_cnt += 1;
            } else {
              if (ref.confs[j] >= conv_set.rules[i].res_conf[j]) {
                chk_cnt += 1;
              }
            }
          }
        }
      }
      if (chk_cnt === NB_CLFS) {
        var env_vars = {
          input: conv_set.input_text,
          date: Utilities.formatDate(new Date(), "JST", "yyyy年MM月dd日"),
          time: Utilities.formatDate(new Date(), "JST", "HH時mm分ss秒")
        };

        if (IS_DEBUG === true) {
          rec = [
            conv_set.rules[i].message,
            conv_set.rules[i].question,
            conv_set.rules[i].res_int[0],
            conv_set.rules[i].res_conf[0],
            conv_set.rules[i].res_int[1],
            conv_set.rules[i].res_conf[1],
            conv_set.rules[i].res_int[2],
            conv_set.rules[i].res_conf[2]
          ];

          debug.getRange(row, 10, 1, rec.length).setValues([rec]);
        }

        var res = CHATAPP_expand_tags(conv_set.rules[i].message, env_vars);
        messages.push({
          message: res.text,
          question: conv_set.rules[i].question
        });
        match_cnt += 1;
        conv_set.rules[i].checked = true;
        break;
      }
    }
  });

  NLCAPP_log_debug({
    record: ["CHATAPP_select_message", "match_cnt", match_cnt]
  });

  if (match_cnt === 0) {
    messages.push({
      message: conv_set.other_msg,
      question: ""
    });
  }

  NLCAPP_log_debug({ record: ["CHATAPP_select_message", "END"] });

  return messages;
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * 会話記録
 * @param       {ConvSet} conv_set    応答設定
 * @param       {Object} res_classes 分類結果
 */
function CHATAPP_store_dialog(conv_set, res_classes) {
  var self_ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = self_ss.getSheetByName(conv_set.ws_name);
  if (sheet === null) {
    sheet = self_ss.insertSheet(conv_set.ws_name);
  }

  var lastRow = sheet.getLastRow();
  lastRow += 1;
  if (lastRow < conv_set.start_row) {
    lastRow = conv_set.start_row;
  }
  sheet.appendRow([conv_set.timestamp, conv_set.input_text, conv_set.messages]);

  sheet
    .getRange(lastRow, 5, 1, 1)
    .setNumberFormat("@")
    .setValue(res_classes[0].class_name);
  sheet.getRange(lastRow, 6, 1, 1).setValue(res_classes[0].confidence);
  sheet.getRange(lastRow, 7, 1, 1).setValue(res_classes[0].timestamp);

  sheet
    .getRange(lastRow, 9, 1, 1)
    .setNumberFormat("@")
    .setValue(res_classes[1].class_name);
  sheet.getRange(lastRow, 10, 1, 1).setValue(res_classes[1].confidence);
  sheet.getRange(lastRow, 11, 1, 1).setValue(res_classes[1].timestamp);

  sheet
    .getRange(lastRow, 13, 1, 1)
    .setNumberFormat("@")
    .setValue(res_classes[2].class_name);
  sheet.getRange(lastRow, 14, 1, 1).setValue(res_classes[2].confidence);
  sheet.getRange(lastRow, 15, 1, 1).setValue(res_classes[2].timestamp);

  //lastRow = sheet.getLastRow();
  // sheet.setActiveRange(sheet.getRange(lastRow + 1, 1));
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * 応答記録
 * @param       {String} input   入力情報
 * @param       {String} res_msg 応答メッセージ
 */
function CHATAPP_store_reply(input, res_msg) { // eslint-disable-line no-unused-vars

  var timestamp = Utilities.formatDate(
    new Date(),
    "JST",
    "yyyy/MM/dd HH:mm:ss"
  );

  var conf = CHATAPP_load_config(CONFIG_SET);

  var sheet = SpreadsheetApp.getActive().getgetSheetByName(
    conf.sheet_conf.ws_name
  );
  if (sheet === null) {
    sheet = SpreadsheetApp.getActive().insertSheet(conf.sheet_conf.ws_name);
  }

  var lastRow = sheet.getLastRow();
  lastRow += 1;
  if (lastRow < conf.sheet_conf.start_row) {
    lastRow = conf.sheet_conf.start_row;
  }

  var record = [
    timestamp,
    input,
    res_msg,
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    ""
  ];
  sheet.appendRow(record);
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * 環境情報の取得
 */
function CHATAPP_prepare_chat() { // eslint-disable-line no-unused-vars

  NLCAPP_log_debug({ record: ["CHATAPP_prepare_chat", "START"] });

  var userProperties = PropertiesService.getUserProperties();

  var conf = CHATAPP_load_config(CONFIG_SET);
  userProperties.setProperty("CONF", JSON.stringify(conf));

  var creds = {};
  try {
    creds = NLCAPP_load_creds();
  } catch (e) {
    Logger.log(e);
    throw e;
  }
  userProperties.setProperty("CREDS", JSON.stringify(creds));

  var log_sheet = conf.self_ss.getSheetByName(conf.sheet_conf.log_ws);
  if (log_sheet === null) {
    log_sheet = conf.self_ss.insertSheet(conf.sheet_conf.log_ws);
  }

  var log_set = {
    sheet: log_sheet,
    self_ss: conf.self_ss,
    ss_id: conf.ss_id,
    ws_name: conf.sheet_conf.log_ws,
    start_col: CONFIG_SET.log_start_col,
    start_row: CONFIG_SET.log_start_row
  };

  var test_set = {
    self_ss: conf.self_ss,
    ss_id: conf.ss_id,
    ws_name: conf.sheet_conf.ws_name,
    start_col: conf.sheet_conf.start_col,
    start_row: conf.sheet_conf.start_row,
    end_row: -1,
    text_col: conf.sheet_conf.text_col
  };

  var nlc = NLCAPP_create_instance(creds.username, creds.password, creds.url);

  var clf_ids = NLCAPP_get_classifiers({
    nlc: nlc,
    log_set: log_set,
    test_set: test_set
  });

  NLCAPP_log_debug({
    record: ["CHATAPP_prepare_chat", JSON.stringify(clf_ids)]
  });

  userProperties.setProperty("CLF_IDS", JSON.stringify(clf_ids));
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * メッセージ送信
 * @param       {String} input_text ユーザー入力
 * @return      {Object} 応答メッセージ
 */
function CHATAPP_send_message(input_text) { // eslint-disable-line no-unused-vars

  NLCAPP_log_debug({ record: ["CHATAPP_send_message", "START"] });

  var userProperties = PropertiesService.getUserProperties();
  var prop = userProperties.getProperty("CONF");
  var conf = JSON.parse(prop);

  NLCAPP_log_debug({
    record: ["CHATAPP_send_message", "conf", JSON.stringify(conf)]
  });

  var conv_conf = {
    ws_name: conf.conv_conf.ws_name,
    start_col: CONFIG_SET.conv_start_col,
    start_row: CONFIG_SET.conv_start_row
  };
  var rules = CHATAPP_load_conv_rules(conv_conf);

  prop = userProperties.getProperty("CREDS");
  var creds = JSON.parse(prop);

  prop = userProperties.getProperty("CLF_IDS");
  var clf_ids = JSON.parse(prop);

  var nlc = NLCAPP_create_instance(creds.username, creds.password, creds.url);

  var conv_set = {
    self_ss: conf.sheet_conf.self_ss,
    ss_id: conf.ss_id,
    ws_name: conf.sheet_conf.ws_name,
    start_col: conf.sheet_conf.start_col,
    start_row: conf.sheet_conf.start_row,
    rules: rules,
    other_msg: conf.chat_conf.other_msg,
    show_suggests: conf.chat_conf.show_suggests,
    input_text: input_text
  };

  NLCAPP_log_debug({
    record: ["CHATAPP_send_message", "other_msg", conv_set.other_msg]
  });

  var timestamp = Utilities.formatDate(
    new Date(),
    "JST",
    "yyyy/MM/dd HH:mm:ss"
  );

  // ３つの分類器にリクエストを投げる
  var res_classes = [];
  var nlc_res;
  //var err_res;
  var has_error = 0;
  for (var j = 0; j < NB_CLFS; j += 1) {
    if (clf_ids[j].status !== "Available") {
      res_classes.push({
        class_name: "",
        confidence: "",
        timestamp: "",
        classes: []
      });
      if (clf_ids[j].status !== "Nothing") {
        has_error = 1;
      }
      continue;
    }

    nlc_res = nlc.classify({
      classifier_id: clf_ids[j].id,
      text: input_text
    });

    if (nlc_res.status !== 200) {
      //err_res = nlc_res;
      has_error = 2;
      res_classes.push({
        class_name: "",
        confidence: "",
        timestamp: "",
        classes: []
      });
    } else {
      res_classes.push({
        class_name: nlc_res.body.top_class,
        confidence: nlc_res.body.classes[0].confidence,
        timestamp: Utilities.formatDate(
          new Date(nlc_res.from),
          "JST",
          "yyyy/MM/dd HH:mm:ss"
        ),
        classes: nlc_res.body.classes
      });
    }
  }

  var msgs = [];
  if (has_error !== 0) {
    msgs.push({
      message: conf.chat_conf.error_msg,
      question: ""
    });
  } else {
    msgs = CHATAPP_select_message(conv_set, res_classes);
  }

  conv_set.messages = msgs[0].message;
  conv_set.timestamp = timestamp;

  CHATAPP_store_dialog(conv_set, res_classes);

  // 履歴をシートに保存
  var result = {
    response: msgs
  };

  RUNTIME_STATUS["CHATAPP_send_message"] = msgs;

  NLCAPP_log_debug({ record: ["CHATAPP_send_message", "END"] });

  return result;
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * [CHATAPP_create_training_data description]
 * @param {Object} p_params params
 * @param {Object} p_params.sheet params
 * @param {Object} p_params.train_set params
 * @return {Object} result
 */
function CHATAPP_create_training_data(p_params) {
  NLCAPP_log_debug({ record: ["CHATAPP_create_training_data", "START"] });

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

    var train_text = String(
      entries[i][p_params.train_set.text_col - p_params.train_set.start_col]
    );
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
 * 分類器の学習
 * @param       {Object} p_params 学習設定
 * @param       {TrainSet} p_params.train_set 学習設定
 * @param       {LogSet} p_params.log_set   ログ設定
 * @throws      {Error}  データシートが不明です
 */
function CHATAPP_train(p_params) { // eslint-disable-line no-unused-vars
  for (var clf_no = 1; clf_no <= NB_CLFS; clf_no += 1) {
    var train_set = {
      ws_name: p_params.conf.sheet_conf.ws_name,
      start_row: p_params.sheet_conf.start_row,
      start_col: p_params.sheet_conf.start_col,
      text_col: p_params.sheet_conf.text_col,
      class_col: p_params.sheet_conf.intent_col[clf_no - 1],
      clf_no: clf_no,
      clf_name: CLFNAME_PREFIX + String(clf_no),
      clfs: p_params.clfs
    };

    var training_data = CHATAPP_create_training_data({
      sheet: p_params.conf.sheet,
      train_set: train_set
    });

    var clf_name = CLFNAME_PREFIX + clf_no;
    var clf_info = NLCAPP_clf_vers({
      clf_list: p_params.clfs.body.classifiers,
      target_name: clf_name
    });
    var new_version = clf_info.max_ver + 1;
    var new_name = clf_name + CLF_SEP + new_version;

    // 分類器の作成
    var train_params = {
      metadata: {
        name: new_name,
        language: "ja"
      },
      training_data: training_data.csvString
    };
    var nlc_res = p_params.nlc.createClassifier(train_params);

    // バージョンが複数ある場合
    if (clf_info.count >= 2 && nlc_res.status === 200) {
      // 分類器の削除
      var del_params = {
        classifier_id: clf_info.clfs[clf_info.min_ver].classifier_id
      };
      p_params.nlc.deleteClassifier(del_params);
    }

    var train_result = {
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
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * [CHATAPP_init_train description]
 * @return {Object} result
 */
function CHATAPP_init_train() {
  var creds = NLCAPP_load_creds();

  var conf = CHATAPP_load_config(CONFIG_SET);

  var sheet = conf.self_ss.getSheetByName(conf.sheet_conf.ws_name);
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
    chat_conf: conf.chat_conf,
    conv_conf: conf.conv_conf,
    sheet: sheet,
    log_set: log_set
  };
}
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
/**
 * 全分類器を学習
 */
function CHATAPP_train_all() { // eslint-disable-line no-unused-vars

  NLCAPP_log_debug({ record: ["CHATAPP_train_all", "START"] });

  var conf = CHATAPP_init_train({
    CONFIG_SET: CONFIG_SET
  });

  for (var clf_no = 1; clf_no <= NB_CLFS; clf_no += 1) {
    var train_set = {
      ss_id: conf.SS_ID,
      ws_name: conf.sheet_conf.ws_name,
      start_row: conf.sheet_conf.start_row,
      start_col: conf.sheet_conf.start_col,
      end_row: -1,
      text_col: conf.chat_conf.text_col,
      class_col: conf.sheet_conf.intent_col[clf_no - 1],
      clf_no: clf_no,
      clf_name: CLFNAME_PREFIX + clf_no
    };

    var training_data = CHATAPP_create_training_data({
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

  NLCAPP_exec_check_clfs();

  NLCAPP_log_debug({ record: ["CHATAPP_train_all", "TRIGGER SET"] });

  GASLIB_Trigger_set("NLCAPP_exec_check_clfs", 1);
}
// ----------------------------------------------------------------------------
