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


/* globals BaseService */

/**
 * NaturalLanguageClassifierV1オブジェクトを構築します。
 *
 * @param {Object} p_options - オプションオブジェクト
 * @param {String} [p_options.url=https://gateway.watsonplatform.net/natural-language-classifier/api] - 資格情報のURL
 * @param {String} [p_options.username] - 認証に使用する資格情報のユーザー名
 * @param {String} [p_options.password] - 認証に使用する資格情報のパスワード
 * @param {String} [p_options.iam_apikey] - 認証に使用する資格情報のAPIキー
 * @param {String} [p_options.iam_access_token] - 認証に使用するIAMアクセストークン
 * @param {Object} [p_options.headers] - デフォルトリクエストヘッダ
 * @param {Boolean} [p_options.headers.X-Watson-Learning-Opt-Out=true] - データ収集オプション `false`に設定すると、すべてのIBM Watsonサービスが要求とその結果をログに記録します。
 * @param {Boolean} [p_options.muteHttpExceptions=false] - UrlFetchApp.fetchのmuteHttpExceptionsオプション
 * @constructor
 * @extends {BaseService}
 * @throws {Error} コンストラクタの呼び出しにはnew演算子が必要です
 */
function NaturalLanguageClassifierV1(p_options) {

  // ====================================================
  // # Classify text

  /**
   * 分類クラス情報
   * @typedef {Object} ClassifiedClass
   * @property {Double} confidence - このクラスに対するWatsonの信頼度を表す10進パーセンテージ。より高い値はより高い信頼度を表します。
   * @property {String} class_name - クラス名
   */

  /**
   * フレーズの分類結果
   * @typedef {Object} Classification
   * @property {String} classifier_id - 分類器のID
   * @property {String} url - 分類器へのリンク
   * @property {String} text - 分類対象のテキスト
   * @property {String} top_class - 最も信頼度の高かったクラス名
   * @property {ClassifiedClass[]} classes - 信頼度の降順にソートされた最大10個の、信頼度とクラス名が列挙された配列。
   */

  // ----------------------------------------------------
  // ## Classify a phrase (GET)
  // GET /v1/classifiers/{classifier_id}/classify
  // ----------------------------------------------------

  // ----------------------------------------------------
  // ## Classify a phrase
  // POST /v1/classifiers/{classifier_id}/classify

  /**
   * フレーズ分類の実行結果
   * @typedef {Object} ClassifyResult
   * @property {Integer} status レスポンスコード
   * @property {Classification}body レスポンスボディ
   * @property {Long} from 開始時刻
   * @property {Long} to 終了時刻
   * @property {Long} delta 応答時間
   */

  /**
   * フレーズの分類
   * <p>使用する分類器は、ステータスが`Available`でなければなりません。<p>
   * @param {Object} p_params - パラメータ
   * @param {string} p_params.classifier_id - 使用する分類器のID
   * @param {string} p_params.text - 分類対象のテキスト(最大2048文字)
   * @param {Object} [params.headers] - カスタムリクエストヘッダ
   * @param {Boolean} [p_params.muteHttpExceptions] - UrlFetchApp.fetchのmuteHttpExceptionsオプション
   * @return {ClassifyResult} 実行結果
   */
  this.classify = function (p_params) {
    this.check_params(p_params, ["classifier_id", "text"]);

    var resource = {
      template: "classifiers/{classifier_id}/classify",
      params: {
        classifier_id: p_params.classifier_id
      }
    };

    var body = {
      text: p_params.text
    };

    var headers = {
      Accept: "application/json"
    };

    // リクエストオプション
    var options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(body),
      headers: headers
    };

    var params = {
      options: options,
      resource: resource
    };

    var res = this.send_request(params);
    return res;
  };
  // ----------------------------------------------------

  // ----------------------------------------------------
  // ## Classify multiple phrases
  // POST /v1/classifiers/{classifier_id}/classify_collection
  // ----------------------------------------------------

  // ====================================================

  // ====================================================
  // # Manage classifiers

  /**
   * 分類器情報
   * @typedef {Object} Classifier
   * @property {String} url - 分類器へのリンク
   * @property {String} classifier_id - 分類器のID
   * @property {String} name - 分類器の名前
   * @property {String} status - 分類器の状態 [Non Existent,Training,Failed,Available,Unavailable]
   * @property {DateTime} created - 分類器の作成日時(UTC)
   * @property {String} status_description - 分類器のステータスに関する追加の詳細。
   * @property {String} language - 分類器が使用する言語
   */

  /**
   * 分類器情報の一覧
   * @typedef {Object} ClassifierList
   * @property {Classifier[]} classifiers - 分類器情報の配列
   */

  // ----------------------------------------------------
  // ## Create classifier
  // POST /v1/classifiers

  /**
   * 分類器作成の実行結果
   * @typedef {Object} CreateClassifierResult
   * @property {Integer} status レスポンスコード
   * @property {Classifier}body レスポンスボディ
   * @property {Long} from 開始時刻
   * @property {Long} to 終了時刻
   * @property {Long} delta 応答時間
   */

  /**
   * 分類器の作成
   *
   * <p>分類器を作成およびトレーニングするためのデータを送信し、新しい分類器に関する情報を返します。</p>
   *
   * @param {Object} p_params - パラメータオブジェクト
   * @param {Object} p_params.metadata - メタデータオブジェクト
   * @param {String} p_params.metadata.language - 言語コード
   * @param {String} [p_params.metadata.name] - 分類器の名前
   * @param {String} p_params.training_data - CSV形式のトレーニングデータ。データには、最大3,000のクラスと20,000のレコードを含めることができます。
   * @param {Object} [params.headers] - カスタムリクエストヘッダ
   * @param {Boolean} [p_params.muteHttpExceptions] - UrlFetchApp.fetchのmuteHttpExceptionsオプション
   * @return {CreateClassifierResult} 実行結果
   */
  this.createClassifier = function (p_params) {
    this.check_params(p_params, ["metadata", "training_data"]);

    var resource = {
      template: "classifiers"
    };

    // メタデータ
    var training_metadata = {
      language: p_params.metadata.language,
      name: p_params.metadata.name
    };

    // バウンダリデータの生成
    var boundary = this.create_boundary();
    var EOL = "\r\n";

    var postbody = Utilities.newBlob(
      "--" +
        boundary +
        EOL +
        'Content-Disposition: form-data; name="training_data"; filename="training.csv"' +
        EOL +
        "Content-Type: application/octet-stream" +
        EOL +
        "Content-Transfer-Encoding: binary" +
        EOL +
        EOL
    ).getBytes();
    postbody = postbody.concat(
      Utilities.newBlob(p_params.training_data).getBytes()
    );
    postbody = postbody.concat(
      Utilities.newBlob(
        "--" +
          boundary +
          EOL +
          'Content-Disposition: form-data; name="training_metadata"' +
          EOL +
          EOL
      ).getBytes()
    );
    postbody = postbody.concat(
      Utilities.newBlob(JSON.stringify(training_metadata)).getBytes()
    );
    postbody = postbody.concat(
      Utilities.newBlob(EOL + "--" + boundary + "--" + EOL).getBytes()
    );

    var headers = {
      Accept: "application/json"
    };

    // リクエストオプション
    var options = {
      method: "post",
      contentType: "multipart/form-data; boundary=" + boundary,
      payload: postbody,
      headers: headers
    };

    var params = {
      options: options,
      resource: resource
    };

    var res = this.send_request(params);
    return res;
  };
  // ----------------------------------------------------

  // ----------------------------------------------------
  // ## List classifiers
  // GET /v1/classifiers

  /**
   * 分類器一覧取得の実行結果
   * @typedef {Object} ListClassifiersResult
   * @property {Integer} status レスポンスコード
   * @property {ClassifierList}body レスポンスボディ
   * @property {Long} from 開始時刻
   * @property {Long} to 終了時刻
   * @property {Long} delta 応答時間
   */

  /**
   * 分類器一覧の取得
   * <p>利用可能な分類器がない場合、実行結果のbodyには空の配列がセットされます。</p>
   *
   * @param {Object} [p_params] - パラメータオブジェクト
   * @param {Object} [p_params.headers] - カスタムリクエストヘッダ
   * @param {Boolean} [p_params.muteHttpExceptions] - UrlFetchApp.fetchのmuteHttpExceptionsオプション
   * @return {ListClassifiersResult} 実行結果
   */
  this.listClassifiers = function (p_params) {
    var resource = {
      template: "classifiers"
    };

    var headers = {
      Accept: "application/json"
    };

    // リクエストオプション
    var options = {
      method: "get",
      contentType: "application/json",
      headers: headers
    };

    var params = {
      options: options,
      resource: resource,
      user_params: p_params
    };

    var res = this.send_request(params);
    return res;
  };
  // ----------------------------------------------------

  // ----------------------------------------------------
  // ## Get information about a classifier
  // GET /v1/classifiers/{classifier_id}

  /**
   * 分類器情報取得の実行結果
   * @typedef {Object} GetClassifierResult
   * @property {Integer} status レスポンスコード
   * @property {Classifier} body レスポンスボディ
   * @property {Long} from 開始時刻
   * @property {Long} to 終了時刻
   * @property {Long} delta 応答時間
   */

  /**
   * 分類器情報の取得
   *
   * <p>分類器に関するステータスとその他の情報を返します。</p>
   *
   * @param {Object} p_params - パラメータオブジェクト
   * @param {String} p_params.classifier_id - 照会する分類器のID
   * @param {Object} [p_params.headers] - カスタムリクエストヘッダ
   * @param {Boolean} [p_params.muteHttpExceptions] - UrlFetchApp.fetchのmuteHttpExceptionsオプション
   * @return {GetClassifierResult} 実行結果
   */
  this.getClassifier = function (p_params) {
    this.check_params(p_params, ["classifier_id"]);

    var resource = {
      template: "classifiers/{classifier_id}",
      params: {
        classifier_id: p_params.classifier_id
      }
    };

    var headers = {
      Accept: "application/json"
    };

    // リクエストオプション
    var options = {
      method: "get",
      contentType: "application/json",
      headers: headers
    };

    var params = {
      options: options,
      resource: resource
    };

    var res = this.send_request(params);
    return res;
  };
  // ----------------------------------------------------

  // ----------------------------------------------------
  // ## Delete classifier
  // DELETE /v1/classifiers/{classifier_id}

  /**
   * 分類器削除の実行結果
   * @typedef {Object} DeleteClassifierResult
   * @property {Integer} status レスポンスコード
   * @property {{}} body レスポンスボディ
   * @property {Long} from 開始時刻
   * @property {Long} to 終了時刻
   * @property {Long} delta 応答時間
   */

  /**
   * 分類器の削除
   *
   * @param {Object} p_params - パラメータオブジェクト
   * @param {string} p_params.classifier_id - 削除する分類器のID
   * @param {Object} [p_params.headers] - カスタムリクエストヘッダ
   * @param {Boolean} [p_params.muteHttpExceptions] - UrlFetchApp.fetchのmuteHttpExceptionsオプション
   * @return {DeleteClassifierResult} 実行結果
   */
  this.deleteClassifier = function (p_params) {
    this.check_params(p_params, ["classifier_id"]);

    var resource = {
      template: "classifiers/{classifier_id}",
      params: {
        classifier_id: p_params.classifier_id
      }
    };

    var headers = {
      Accept: "application/json"
    };

    // リクエストオプション
    var options = {
      method: "delete",
      contentType: "application/json",
      headers: headers
    };

    var params = {
      options: options,
      resource: resource
    };

    var res = this.send_request(params);
    return res;
  };
  // ----------------------------------------------------

  // ====================================================

  if (!(this instanceof NaturalLanguageClassifierV1)) {
    throw Error("コンストラクタの呼び出しにはnew演算子が必要です");
  }

  // デフォルト値設定
  this.uri_endpoint =
    "https://gateway.watsonplatform.net/natural-language-classifier/api";
  this.api_version = "v1";
  this.muteHttpExceptions = true;
  this.default_headers = {
    "X-Watson-Learning-Opt-Out": true
  };

  BaseService.call(this, p_options);

  // 資格情報のチェック
  var r = this.check_creds(p_options);
  if (r !== 1) {
    throw Error("資格情報が不正です");
  }

  if ("version" in p_options) {
    this.version_date = p_options.version;
  }
}

NaturalLanguageClassifierV1.prototype = Object.create(BaseService.prototype);
NaturalLanguageClassifierV1.prototype.constructor = NaturalLanguageClassifierV1;
