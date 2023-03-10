function onOpen() {

  const ui = DocumentApp.getUi(); 
  ui.createMenu('ðã«ã¹ã¿ã ã¡ãã¥ã¼')
    .addItem('ðæãã¦ï¼ChatGPTï¼', 'main')
    .addToUi();
}

/** å®è¡ãããã¨ã³ããªã¼ãã¤ã³ã */
function main() {
  const word = getTargetWord();
  const text = askChatGPT(word)
  setText(text);
}


/**
 * ãã­ã¹ããåãåã£ã¦ãã­ã¥ã¡ã³ãã«è²¼ãä»ããé¢æ°
 * @param {string} æå­å
*/
function setText(text) {
  const doc = DocumentApp.getActiveDocument();
  const paragraph = doc.getParagraphs()[0];
  paragraph.appendText(`${text}\n\n`); //\nã¯æ¹è¡
}

/**
 * é¸æããç¯å²ã®åèªãåå¾ããé¢æ°
 * @return  {string} åèª
*/
function getTargetWord() {
  const doc = DocumentApp.getActiveDocument();
  const word = doc.getSelection().getRangeElements()[0].getElement().getText();
  return word;
}


/**
 * ChatGPTã«è³ªåããé¢æ°
 * @param {string} è³ªåããåèª
 * @return {string} åç­ã®ãã­ã¹ã
*/
function askChatGPT(word) {

  // API Key
  const apiKey = 'API Keyãã³ãã';

  //ã¨ã³ããã¤ã³ãã¨ãã©ã¡ã¼ã¿
  const endpoint = 'https://api.openai.com/v1/completions';
  const params = {
    'prompt': word,
    'model': 'text-davinci-002',
    'max_tokens': 1000,
    'temperature': 0.8
  };

  //ãããã¼ã¨ãªãã·ã§ã³
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

  //ã¬ã¹ãã³ã¹ã®ãªãã¸ã§ã¯ãå
  const data = JSON.parse(response.getContentText());

  //ãªãã¸ã§ã¯ããããã­ã¹ãæ½åº
  const answer = data.choices[0].text;
  console.log(answer);
  return answer
}
