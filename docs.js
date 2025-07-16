// =================================================================
// CONFIGURATION & CONSTANTS
// =================================================================

const FOLDER_ID = '1_M44p2Cqv-1tdBJLsOS_3MtaRZ98IqhL'; // IMPORTANT: Make sure this is your correct folder ID
const AZURE_OPENAI_API_KEY = PropertiesService.getScriptProperties().getProperty('AZURE_OPENAI_API_KEY');
const AZURE_API_VERSION = '2024-07-01-preview';
const AZURE_DEPLOYMENT_MODEL = 'gpt-4o';
const AZURE_APIM_BASE_URL = "https://mf-genai-poc-apim.azure-api.net/esotad";

// --- Headers for the final documents ---
const DOC_LINK_HEADER_JP = "ドキュメントリンク";
const DOC_LINK_HEADER_EN = "Document Link";
const INTERNAL_NOTES_LINK_HEADER = "Internal Notes Link";

// --- List of columns to include in the Job Description ---
const JD_COLUMN_HEADERS = [
  "職種タイトル / Job Title",
  "募集背景 / Background of the Recruitment",
  "主な業務内容 / Main Responsibilities",
  "仕事のやりがい・得られる経験 / Job Satisfaction and Experience Gained",
  "期待する役割 / Expected Role",
  "期待するマインド / Expected Mindset",
  "求めるスキル・経験 / Desired Skills and Experience",
  "あると望ましいスキル・経験 / Preferred Skills and Experience",
  "日本語要件 / Japanese Language Requirements",
  "英語要件 / English Language Requirements",
  "こんな方に仲間になってほしい / We are looking for someone like this to join our team.",
  "技術スタック / Technology Stack",
  "使用ツール / Tools Used",
  "参考URL / Reference URL"
];

// --- List of columns to include in the Internal Notes document ---
const INTERNAL_NOTES_COLUMN_HEADERS = [
  "担当ハイヤリングマネージャー名 / Name of the Responsible Hiring Manager",
  "担当リクルーター名 / Name of the Responsible Recruiter",
  "職種 / Job Title",
  "勤務形態 / Employment Type",
  "勤務地 / Work Location",
  "配属先部署名 / Assigned Department Name",
  "募集背景 / Background of the Recruitment",
  "具体的な募集背景 / Specific Background of the Recruitment",
  "採用納期 / Hiring Deadline",
  "上記の理由/背景をご教示下さい / Please provide the reason/background for the above.",
  "採用温度感 / Urgency of Hiring",
  "想定グレード / Expected Grade",
  "オファー年収のイメージ / Estimated Annual Salary Offer",
  "年齢 / Ageお任せしたい業務 / Tasks to be Assigned",
  "上記をお任せすることにあたって必要な経験・スキル / Experience and Skills Required for the Above-mentioned Tasks",
  "ターゲット企業や業界 / Target Companies and Industries",
  "技術課題の有無 / Presence of Technical Challenges",
  "＜上記質問で「あり_track」を選択した方＞該当課題のURLを展開して下さい / <For those who selected \"Present (Track)\" in the above question> Please provide the URL for the relevant challenge.",
  "＜上記質問で「あり_track以外」を選択した方＞課題を展開して下さい / <For those who selected \"Present (Other than Track)\" in the above question> Please outline the challenge.",
  "技術課題レビュー担当者 / Reviewer for Technical Challenges",
  "カジュアル面談 担当者 / Casual Interview Representative",
  "カジュアル面談担当者の英語対応可否 / English Proficiency of the Casual Interview Representative",
  "一次面接 担当者 / First Interview Representative",
  "一次面接担当者の英語対応可否 / English Proficiency of the First Interview Representative",
  "二次面接 担当者 / Second Interview Representative",
  "二次面接担当者の英語対応可否 / English Proficiency of the Second Interview Representative",
  "最終面接 担当者 / Final Interview Representative",
  "最終面接担当者の英語対応可否 / English Proficiency of the Final Interview Representative",
  "その他 / Otherオファー面談 担当者 / Offer Meeting Representative",
  "エージェント利用可否 / Availability of Agent Usage",
  "ビザサポートが必要な海外在住者に対してオープンしますか？ / Are you open to candidates residing overseas who require visa support?"
];


// =================================================================
// MAIN TRIGGERED FUNCTION
// =================================================================

/**
 * Main function triggered by a form submission. Orchestrates the entire process.
 * This function will be automatically run by Google when a new form response is received.
 * @param {Object} e The event object passed by the onFormSubmit trigger.
 */
function processFormSubmission(e) {
  const ui = SpreadsheetApp.getUi();
  try {
    // --- 1. GET DATA FROM THE EVENT ---
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const submittedRange = e.range; // Get the range of the newly submitted row
    const submittedRowIndex = submittedRange.getRowIndex();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowData = sheet.getRange(submittedRowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    const jdTitleHeader = "職種タイトル / Job Title";
    const jdTitleIndex = headers.indexOf(jdTitleHeader);
    const jdTitle = jdTitleIndex !== -1 && rowData[jdTitleIndex] ? rowData[jdTitleIndex] : '不明な求人票';
    
    // --- 2. BUILD CONTENT & PROMPT ---
    const jdPrompt = buildJdPrompt(headers, rowData);
    const internalNotesContent = buildInternalNotesContent(headers, rowData);

    if (!jdPrompt && !internalNotesContent) {
      Logger.log("Row " + submittedRowIndex + " appears to be empty. Skipping.");
      return;
    }

    // --- 3. GENERATE JD (if content exists) ---
    let jdDocUrl = "";
    if (jdPrompt) {
      const summaryText = getAiSummary(jdPrompt);
      const docTitle = `【求人票】${jdTitle}`;
      jdDocUrl = createGoogleDoc(docTitle, summaryText);
    }

    // --- 4. CREATE INTERNAL NOTES DOC (if content exists) ---
    let internalNotesDocUrl = "";
    if (internalNotesContent) {
      const docTitle = `[Internal Notes] - ${jdTitle}`;
      internalNotesDocUrl = createGoogleDoc(docTitle, internalNotesContent);
    }
    
    // --- 5. WRITE LINKS BACK TO SHEET ---
    updateSheetWithLinks(sheet, submittedRowIndex, headers, jdDocUrl, internalNotesDocUrl);

  } catch (error) {
    // Log the error for debugging, and optionally alert the user or send an email.
    Logger.log("Error in processFormSubmission: " + error.toString());
    Logger.log("Stack: " + error.stack);
    // ui.alert('自動処理中にエラーが発生しました: ' + error.message); // This might be disruptive, so logging is often better.
  }
}


// =================================================================
// LOGIC & CONTENT-BUILDING FUNCTIONS
// =================================================================

/**
 * Builds the AI prompt using only the columns specified in JD_COLUMN_HEADERS.
 * @param {string[]} headers - All header strings from the sheet.
 * @param {any[]} rowData - Data for the submitted row.
 * @returns {string|null} The fully formatted prompt string, or null if no relevant data is found.
 */
function buildJdPrompt(headers, rowData) {
  let interviewNotes = "";
  // Instead of looping through all headers, we loop through our specific list
  JD_COLUMN_HEADERS.forEach(headerName => {
    const colIndex = headers.indexOf(headerName);
    if (colIndex !== -1 && rowData[colIndex] && rowData[colIndex].toString().trim() !== '') {
      interviewNotes += `${headerName}:\n${rowData[colIndex]}\n\n`;
    }
  });

  if (interviewNotes.trim() === "") {
    return null;
  }

  // The prompt template remains the same
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
- The job description sections have both Japanese and English section headers. (Eg '募集背景 / Background of the Recruitment’). Please only output the section headers in that language (eg. For the Japanese version only display '募集背景' and leave out 'Background of the Recruitment').
- Within the Technology Stack and Tools Used sections, please do the same and only output the appropriate language (eg. For the Japanese version only display 'リポジトリ管理' and leave out 'Repository Management').
- For 英語要件 / English Language Requirements section, please always include the following at the end of the section depending on the language:English: (Note: If you have other qualifications or experiences demonstrating English proficiency, such as EIKEN Pre-1, EIKEN 2nd Grade (CSE score 1950+), TOEFL iBT 60+, IELTS 5.0+, or Cambridge FCE.), feel free to discuss with us) For those without a TOEIC 700+ equivalent score, they will be asked to take a designated test during the interview process (generally after the first interview).Japanese: ※TOEIC以外にも英語力がわかる資格や経験をお持ちの方はご相談ください例：英検準1級、英検2級（英検CSEスコア1950以上）、TOEFL iBT 60以上、IELTS 5.0以上、ケンブリッジ英語検定FCEなど※その他、英語力がわかる資格や経験については応相談※TOEIC 700点相当以上の資格をお持ちでない方については選考の過程で弊社指定の試験を受験いただきます。（原則、一次面接後を想定）
- For あると望ましいスキル・経験 / Preferred Skills and Experience section, please always include the following at the end of the section depending on the language:English: Experience in AI development and/or experience in using AI tools to improve development processes.Money Forward recently announced our AI Strategy roadmap which focuses on improving AI-driven operational efficiencies, as well as integrating AI agents into our products to deliver better value to our users. (More information here)Japanese: AIの開発経験もしくはAIツールを使用した開発経験Money Forward AI Vision 2025にて発表の通り、マネーフォワードではAIを使った業務効率化に取り組んでいる状況かつ、将来的には全製品にAIエージェントを導入する想定であるため


If the submitted information is blank for a certain section, that section can be omitted.

If the submitted information is minimal, please expand the information in that section as necessary, so the candidate has a good idea of the position's details.
Please try to aim for 3-4 sentences when expanding a section.


The JD sections are below:


職種タイトル / Job Title

募集背景 / Background of the Recruitment

主な業務内容 / Main Responsibilities*

仕事のやりがい・得られる経験 / Job Satisfaction and Experience Gained

期待する役割 / Expected Role

期待するマインド / Expected Mindset

求めるスキル・経験 / Desired Skills and Experience*

あると望ましいスキル・経験 / Preferred Skills and Experience*

日本語要件 / Japanese Language Requirements*

英語要件 / English Language Requirements*

こんな方に仲間になってほしい / We are looking for someone like this to join our team.

技術スタック / Technology Stack

使用ツール / Tools Used

参考URL / Reference URL


The submitted information is below:


  ${interviewNotes}
  `;
   return prompt;
}

/**
 * Builds a simple text block using only the columns specified in INTERNAL_NOTES_COLUMN_HEADERS.
 * @param {string[]} headers - All header strings from the sheet.
 * @param {any[]} rowData - Data for the submitted row.
 * @returns {string|null} A simple string of the internal notes data, or null if no relevant data is found.
 */
function buildInternalNotesContent(headers, rowData) {
  let notesContent = "";
  INTERNAL_NOTES_COLUMN_HEADERS.forEach(headerName => {
    const colIndex = headers.indexOf(headerName);
    if (colIndex !== -1 && rowData[colIndex] && rowData[colIndex].toString().trim() !== '') {
      notesContent += `${headerName}:\n${rowData[colIndex]}\n\n`;
    }
  });

  return notesContent.trim() === "" ? null : notesContent;
}

// =================================================================
// GOOGLE & AZURE SERVICE FUNCTIONS
// =================================================================

/**
 * Creates a Google Doc, moves it to the correct folder, and returns the URL.
 * (This function is unchanged)
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
 * Finds or creates the necessary columns and inserts the document URLs in the correct row.
 */
function updateSheetWithLinks(sheet, rowIndex, headers, jdDocUrl, internalNotesDocUrl) {
    // --- Update JD Link ---
    if (jdDocUrl) {
      let linkColumnIndex = headers.indexOf(DOC_LINK_HEADER_JP);
      if (linkColumnIndex === -1) {
          linkColumnIndex = headers.indexOf(DOC_LINK_HEADER_EN);
      }
      
      if (linkColumnIndex === -1) { // If column still doesn't exist, create it.
        const newColumnIndex = sheet.getLastColumn() + 1;
        sheet.getRange(1, newColumnIndex).setValue(DOC_LINK_HEADER_JP).setFontWeight('bold');
        sheet.getRange(rowIndex, newColumnIndex).setValue(jdDocUrl);
        headers.push(DOC_LINK_HEADER_JP); // Update headers array for the next step
      } else {
        sheet.getRange(rowIndex, linkColumnIndex + 1).setValue(jdDocUrl);
      }
    }

    // --- Update Internal Notes Link ---
    if (internalNotesDocUrl) {
        let internalNotesColIndex = headers.indexOf(INTERNAL_NOTES_LINK_HEADER);

        if (internalNotesColIndex === -1) { // If column doesn't exist, create it
            const newColumnIndex = sheet.getLastColumn() + 1;
            sheet.getRange(1, newColumnIndex).setValue(INTERNAL_NOTES_LINK_HEADER).setFontWeight('bold');
            sheet.getRange(rowIndex, newColumnIndex).setValue(internalNotesDocUrl);
        } else {
            sheet.getRange(rowIndex, internalNotesColIndex + 1).setValue(internalNotesDocUrl);
        }
    }
}


/**
 * Calls the Azure OpenAI API to get an AI-generated summary.
 * (This function is unchanged)
 */
function getAiSummary(prompt) {
  const url = `${AZURE_APIM_BASE_URL}/openai/deployments/${AZURE_DEPLOYMENT_MODEL}/chat/completions?api-version=${AZURE_API_VERSION}`;
  const payload = {
    "model": AZURE_DEPLOYMENT_MODEL,
    "messages": [{ "role": "user", "content": prompt }]
  };
  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'headers': { 'api-key': AZURE_OPENAI_API_KEY },
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  const responseBody = response.getContentText();

  if (responseCode === 200) {
    const json = JSON.parse(responseBody);
    if (json.choices && json.choices.length > 0 && json.choices[0].message && json.choices[0].message.content) {
      return json.choices[0].message.content;
    } else {
      throw new Error('AIからの応答が無効な形式です。応答: ' + responseBody);
    }
  } else {
    throw new Error(`Azure OpenAI APIの呼び出しエラー。ステータス: ${responseCode}, 応答: ${responseBody}`);
  }
}

// =================================================================
// ONE-TIME SETUP FUNCTION
// =================================================================

/**
 * You only need to run this function ONCE to set up the automatic trigger.
 */
function createOnSubmitTrigger() {
  const sheet = SpreadsheetApp.getActive();
  ScriptApp.newTrigger('processFormSubmission')
    .forSpreadsheet(sheet)
    .onFormSubmit()
    .create();
  
  SpreadsheetApp.getUi().alert('The trigger has been created! The script will now run automatically on new form submissions.');
}