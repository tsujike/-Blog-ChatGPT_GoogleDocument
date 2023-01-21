function onOpen() {

  const ui = DocumentApp.getUi(); 
  ui.createMenu('🌈カスタムメニュー')
    .addItem('🚀教えて！ChatGPT！', 'main')
    .addToUi();
}

/** 実行されるエントリーポイント */
function main() {
  const word = getTargetWord();
  const text = askChatGPT(word)
  setText(text);
}


/**
 * テキストを受け取ってドキュメントに貼り付ける関数
 * @param {string} 文字列
*/
function setText(text) {
  const doc = DocumentApp.getActiveDocument();
  const paragraph = doc.getParagraphs()[0];
  paragraph.appendText(`${text}\n\n`); //\nは改行
}

/**
 * 選択した範囲の単語を取得する関数
 * @return  {string} 単語
*/
function getTargetWord() {
  const doc = DocumentApp.getActiveDocument();
  const word = doc.getSelection().getRangeElements()[0].getElement().getText();
  return word;
}


/**
 * ChatGPTに質問する関数
 * @param {string} 質問する単語
 * @return {string} 回答のテキスト
*/
function askChatGPT(word) {

  // API Key
  const apiKey = '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';
 
  // API Key
  // const apiKey = "API Keyを入力";

  //エンドポイントとパラメータ
  const endpoint = 'https://api.openai.com/v1/completions';
  const params = {
    'prompt': word,
    'model': 'text-davinci-002',
    'max_tokens': 1000,
    'temperature': 0.8
  };

  //ヘッダーとオプション
  const headers = {
    'Content-Type': 'application/json',
    'Authorization': 'Bearer ' + apiKey
  };

  const options = {
    'method': 'post',
    'headers': headers,
    'payload': JSON.stringify(params)
  };
  const response = UrlFetchApp.fetch(endpoint, options);

  //レスポンスのオブジェクト化
  const data = JSON.parse(response.getContentText());

  //オブジェクトからテキスト抽出
  const answer = data.choices[0].text;
  console.log(answer);
  return answer
}