// =================================================================
// CONFIGURATION & CONSTANTS 
// =================================================================

const FOLDER_ID = '1_M44p2Cqv-1tdBJLsOS_3MtaRZ98IqhL';
const AZURE_OPENAI_API_KEY = PropertiesService.getScriptProperties().getProperty('AZURE_OPENAI_API_KEY');
const AZURE_API_VERSION = '2024-07-01-preview'; 
const AZURE_DEPLOYMENT_MODEL = 'gpt-4o'; 
const AZURE_APIM_BASE_URL = "https://mf-genai-poc-apim.azure-api.net/esotad";
const DOC_LINK_HEADER_JP = "ドキュメントリンク";
const DOC_LINK_HEADER_EN = "Document Link";

// =================================================================
// PURE LOGIC FUNCTION (Testable)
// =================================================================

/**
 * Builds the prompt for the AI model based on spreadsheet data.
 * This is a "pure" function: it has no side effects and doesn't call any external services.
 * @param {string[]} headers - An array of header strings from the sheet.
 * @param {any[]} rowData - An array of data for the selected row.
 * @returns {string|null} The fully formatted prompt string, or null if there's no data.
 */
function buildJdPrompt(headers, rowData) {
  let interviewNotes = "";
  for (let i = 0; i < headers.length; i++) {
    // Check if data exists and the header is not a link column
    if (rowData[i] && rowData[i].toString().trim() !== '' && headers[i].toLowerCase() !== DOC_LINK_HEADER_EN.toLowerCase() && headers[i].toLowerCase() !== DOC_LINK_HEADER_JP.toLowerCase()) {
      interviewNotes += `${headers[i]}:\n${rowData[i]}\n\n`;
    }
  }

  if (interviewNotes.trim() === "") {
    return null; // Return null if there's nothing to process
  }

  // The prompt template is now neatly contained in this function
  const prompt = `
I would like help creating a job description (JD) based on the submitted information from the hiring department.

I will list the JD sections and submitted information down below. 
Please parse out the necessary information from the submitted information and fill in the JD sections accordingly.

Output format:
The output format should start with the JD section header, followed by the submitted information on the next line, as below:
(Section Header)
(Submitted information)

Output language:
I would like you to output two versions of the JD - one in Japanese and one in English. 
The submitted information will be mainly written in one language - please first make the JD in that language, and then use that as a base to translate and make the JD in the other language. Other notes:
- The job description sections have both Japanese and English section headers. (Eg '募集背景 /  Background of the Recruitment’). Please only output the section headers in that language (eg. For the Japanese version only display '募集背景' and leave out 'Background of the Recruitment'). 
- Within the Technology Stack and Tools Used sections, please do the same and only output the appropriate language  (eg. For the Japanese version only display  'リポジトリ管理' and leave out 'Repository Management').
-  For 英語要件 /  English Language Requirements section, please always include the following at the end of the section depending on the language:English: (Note: If you have other qualifications or experiences demonstrating English proficiency, such as EIKEN Pre-1, EIKEN 2nd Grade (CSE score 1950+), TOEFL iBT 60+, IELTS 5.0+, or Cambridge FCE.), feel free to discuss with us) For those without a TOEIC 700+ equivalent score, they will be asked to take a designated test during the interview process (generally after the first interview).Japanese: ※TOEIC以外にも英語力がわかる資格や経験をお持ちの方はご相談ください例：英検準1級、英検2級（英検CSEスコア1950以上）、TOEFL iBT 60以上、IELTS 5.0以上、ケンブリッジ英語検定FCEなど※その他、英語力がわかる資格や経験については応相談※TOEIC 700点相当以上の資格をお持ちでない方については選考の過程で弊社指定の試験を受験いただきます。（原則、一次面接後を想定）
-  For あると望ましいスキル・経験 /  Preferred Skills and Experience section, please always include the following at the end of the section depending on the language:English: Experience in AI development and/or experience in using AI tools to improve development processes.Money Forward recently announced our AI Strategy roadmap which focuses on improving AI-driven operational efficiencies, as well as integrating AI agents into our products to deliver better value to our users. (More information here)Japanese: AIの開発経験もしくはAIツールを使用した開発経験Money Forward AI Vision 2025にて発表の通り、マネーフォワードではAIを使った業務効率化に取り組んでいる状況かつ、将来的には全製品にAIエージェントを導入する想定であるため


If the submitted information is blank for a certain section, that section can be omitted.

If the submitted information is minimal, please expand the information in that section as necessary, so the candidate has a good idea of the position's details.
Please try to aim for 3-4 sentences when expanding a section. 


The JD sections are below:


職種タイトル / Job Title

募集背景 /  Background of the Recruitment

主な業務内容 /  Main Responsibilities* 

仕事のやりがい・得られる経験 /  Job Satisfaction and Experience Gained

期待する役割 /  Expected Role

期待するマインド /  Expected Mindset

求めるスキル・経験 /  Desired Skills and Experience*

あると望ましいスキル・経験 /  Preferred Skills and Experience*

日本語要件 /  Japanese Language Requirements*

英語要件 /  English Language Requirements*

こんな方に仲間になってほしい /  We are looking for someone like this to join our team.

技術スタック /  Technology Stack

使用ツール /  Tools Used

参考URL /  Reference URL


The submitted information is below:


  ${interviewNotes}
  `;
  
  return prompt;
}


// =================================================================
// GOOGLE & AZURE SERVICE FUNCTIONS (The "Doing" parts)
// =================================================================

/**
 * Main function triggered by the menu. Orchestrates the entire process.
 */
function createCandidateSummary() {
  const ui = SpreadsheetApp.getUi();
  try {
    // --- 1. PRE-CHECKS ---
    if (FOLDER_ID === 'YOUR_FOLDER_ID_HERE' || !AZURE_OPENAI_API_KEY) {
      ui.alert("設定エラー: スクリプト内の 'FOLDER_ID' を、そしてスクリプトプロパティで 'AZURE_OPENAI_API_KEY' を更新してください。");
      return;
    }

    // --- 2. GET DATA FROM SHEET ---
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const activeRange = sheet.getActiveRange();
    if (activeRange.getNumRows() !== 1) {
      ui.alert('サマリーを作成したい候補者の行を1行だけ選択してください。');
      return;
    }
    const activeRowIndex = activeRange.getRowIndex();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowData = sheet.getRange(activeRowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // --- 3. BUILD THE PROMPT (Call our new "logic" function) ---
    const prompt = buildJdPrompt(headers, rowData);

    if (!prompt) {
      ui.alert("選択された行は空のようです。サマリーは生成できません。");
      return;
    }
    
    ui.alert('AIサマリーを生成中です... しばらくお待ちください。');

    // --- 4. CALL AI & CREATE DOC ---
    const summaryText = getAiSummary(prompt);
    const jdTitle = rowData[0] || '不明な求人票';
    const docTitle = `【求人票】${jdTitle}`;
    
    const docUrl = createGoogleDoc(docTitle, summaryText);

    // --- 5. WRITE LINK BACK TO SHEET ---
    updateSheetWithLink(sheet, activeRowIndex, headers, docUrl);
    
    ui.alert(`${jdTitle}の求人票が正常に作成されました。シートにリンクが追加されています。`);

  } catch (error) {
    Logger.log(error.toString());
    ui.alert('エラーが発生しました: ' + error.message);
  }
}

/**
 * Creates a Google Doc, moves it to the correct folder, and returns the URL.
 * @param {string} docTitle - The title for the new Google Doc.
 * @param {string} content - The text content to put in the doc.
 * @return {string} The URL of the newly created document.
 */
function createGoogleDoc(docTitle, content) {
  const doc = DocumentApp.create(docTitle);
  doc.getBody().appendParagraph(content);
  doc.saveAndClose();
  
  const docFile = DriveApp.getFileById(doc.getId());
  const folder = DriveApp.getFolderById(FOLDER_ID);
  folder.addFile(docFile);
  DriveApp.getRootFolder().removeFile(docFile); // Important: clean up root
  
  return doc.getUrl();
}

/**
* Finds the "Document Link" column or creates it, then inserts the URL in the correct row.
* @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to modify.
* @param {number} rowIndex - The index of the row to update.
* @param {string[]} headers - The array of header values.
* @param {string} docUrl - The URL to insert.
*/
function updateSheetWithLink(sheet, rowIndex, headers, docUrl) {
    let linkColumnIndex = headers.indexOf(DOC_LINK_HEADER_JP);
    if (linkColumnIndex === -1) {
        linkColumnIndex = headers.indexOf(DOC_LINK_HEADER_EN);
    }
    
    if (linkColumnIndex === -1) { // If column still doesn't exist, create it.
      const newColumnIndex = sheet.getLastColumn() + 1;
      sheet.getRange(1, newColumnIndex).setValue(DOC_LINK_HEADER_JP).setFontWeight('bold');
      sheet.getRange(rowIndex, newColumnIndex).setValue(docUrl);
    } else {
      // The linkColumnIndex is 0-based, but sheet columns are 1-based.
      sheet.getRange(rowIndex, linkColumnIndex + 1).setValue(docUrl);
    }
}


// The getAiSummary function remains unchanged as its only job is to call the API.
/**
 * Calls the Azure OpenAI API to get an AI-generated summary.
 * @param {string} prompt The prompt to send to the model.
 * @return {string} The text content from the AI response.
 */
function getAiSummary(prompt) {
  // --- URL CONSTRUCTION: NOW EXACTLY MATCHING PYTHON CLIENT'S BEHAVIOR ---
  // Combines the base URL, the standard Azure OpenAI path components, and the API version.
  const url = `${AZURE_APIM_BASE_URL}/openai/deployments/${AZURE_DEPLOYMENT_MODEL}/chat/completions?api-version=${AZURE_API_VERSION}`;

  const payload = {
    // The 'model' field is explicitly specified with the deployment name.
    "model": AZURE_DEPLOYMENT_MODEL,
    "messages": [{
      "role": "user",
      "content": prompt
    }]
    // You can add other parameters like temperature, max_tokens, etc. here if needed.
    // "temperature": 0.7,
    // "max_tokens": 800
  };

  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'headers': {
      // --- AUTHENTICATION HEADER: NOW EXACTLY MATCHING PYTHON CLIENT'S BEHAVIOR ---
      // The Python client uses 'api_key', which translates to the 'api-key' header.
      'api-key': AZURE_OPENAI_API_KEY
    },
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true // Important to catch errors
  };

  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  const responseBody = response.getContentText();

  if (responseCode === 200) {
    const json = JSON.parse(responseBody);
    // Safely navigate the JSON structure to get the text
    if (json.choices && json.choices.length > 0 && json.choices[0].message && json.choices[0].message.content) {
        return json.choices[0].message.content;
    } else {
        // Handle cases where the response structure is unexpected
        throw new Error('AIからの応答が無効な形式です。応答: ' + responseBody);
    }
  } else {
    // Throw an error with details from the API response
    throw new Error(`Azure OpenAI APIの呼び出しエラー。ステータス: ${responseCode}, 応答: ${responseBody}`);
  }
}


// The onOpen function remains unchanged.
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('サマリー生成')
      .addItem('AI候補者サマリーを作成 (Azure OpenAI)', 'createCandidateSummary')
      .addToUi();
}

// Add this at the VERY END of your Google Apps Script file
module.exports = { buildJdPrompt };