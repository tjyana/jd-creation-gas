// =================================================================
// CONFIGURATION & CONSTANTS
// =================================================================

// This Folder ID will be used for ALL document creation in this script.
const FOLDER_ID = '1_M44p2Cqv-1tdBJLsOS_3MtaRZ98IqhL'; 
const AZURE_OPENAI_API_KEY = PropertiesService.getScriptProperties().getProperty('AZURE_OPENAI_API_KEY');
const AZURE_API_VERSION = '2024-07-01-preview';
const AZURE_DEPLOYMENT_MODEL = 'gpt-4o';
const AZURE_APIM_BASE_URL = "https://mf-genai-poc-apim.azure-api.net/esotad";

// --- Headers for the final document links in the sheet ---
const DOC_LINK_HEADER_JP = "求人票 (日本語)";
const DOC_LINK_HEADER_EN = "Job Description (English)";

// --- List of columns for the AI to generate the Job Description ---
const JD_COLUMN_HEADERS = [
  "職種タイトル / Job Title", "募集背景 /  Background of the Recruitment", "主な業務内容 /  Main Responsibilities",
  "仕事のやりがい・得られる経験 /  Job Satisfaction and Experience Gained", "期待する役割 /  Expected Role",
  "期待するマインド /  Expected Mindset", "求めるスキル・経験 /  Desired Skills and Experience",
  "あると望ましいスキル・経験 /  Preferred Skills and Experience", "日本語要件 /  Japanese Language Requirements",
  "英語要件 /  English Language Requirements", "こんな方に仲間になってほしい /  We are looking for someone like this to join our team.",
  "技術スタック /  Technology Stack", "使用ツール /  Tools Used", "参考URL /  Reference URL"
];


// =================================================================
// MENU & TRIGGERS
// =================================================================

/**
 * Creates the custom menu in the spreadsheet UI when the file is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('カスタムメニュー')
    .addItem('選択行から求人票を作成 (AI)', 'runAiOnSelectedRows')
    .addSeparator()
    .addItem('選択行を1つのDocsにまとめる (旧機能)', 'exportToSingleDoc_legacy')
    .addToUi();
}

/**
 * This is the AUTOMATIC trigger. It runs when a new form is submitted.
 * It simply finds out which row was added and tells the core engine to process it.
 */
function processFormSubmission(e) {
  const rowIndex = e.range.getRowIndex();
  Logger.log(`Automatic trigger fired for row: ${rowIndex}`);
  generateJdFromRow(rowIndex);
}

/**
 * This is the NEW MANUAL trigger, run from the menu.
 * It finds all the rows with a checked box and tells the core engine to process them one by one.
 */
function runAiOnSelectedRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const checkedRows = [];

  // Start from the second row (index 1) to skip the header
  for (let i = 1; i < data.length; i++) {
    // IMPORTANT: Assuming the checkbox is in the 2nd column (index 1)
    if (data[i][1] === true) {
      // The row index in the sheet is the array index + 1
      checkedRows.push(i + 1);
    }
  }

  if (checkedRows.length === 0) {
    SpreadsheetApp.getUi().alert('選択された回答がありません。\nB列のチェックボックスで回答を選択してください。');
    return;
  }

  // Process each checked row
  checkedRows.forEach(rowIndex => {
    Logger.log(`Manual trigger running for selected row: ${rowIndex}`);
    generateJdFromRow(rowIndex);
  });

  SpreadsheetApp.getUi().alert(`処理を開始しました。\n${checkedRows.length}件の求人票作成が完了すると、各行にリンクが生成されます。`);
}


// =================================================================
// THE CORE AI ENGINE
// =================================================================

/**
 * This is the main "engine" of our script. It takes a row number,
 * performs all the AI and document creation magic, and updates the sheet.
 * Both the automatic and manual triggers use this function.
 * @param {number} rowIndex The row number in the sheet to process.
 */
function generateJdFromRow(rowIndex) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  try {
    // --- 1. GET DATA FOR THE SPECIFIC ROW ---
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowData = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];

    // --- 2. TRANSLATE TITLE ---
    const originalJdTitle = getSheetData(headers, rowData, "職種タイトル / Job Title", '不明な求人');
    const japaneseJdTitle = LanguageApp.translate(originalJdTitle, '', 'ja');
    const englishJdTitle = LanguageApp.translate(originalJdTitle, '', 'en');
    const jdDeptName = getSheetData(headers, rowData, "配属先部署名 /  Assigned Department Name", '不明な部署');
    const jdWorkLocation = getSheetData(headers, rowData, "勤務地 /  Work Location", '不明な勤務地');

    // --- 3. BUILD PROMPT & GET AI RESPONSE ---
    const jdPrompt = buildJdPrompt(headers, rowData);
    if (!jdPrompt) {
      Logger.log(`Row ${rowIndex} is empty or has no data for JD creation. Skipping.`);
      return;
    }
    const fullAiResponse = getAiSummary(jdPrompt);

    // --- 4. PARSE AND BUILD FINAL CONTENT ---
    const parsedJds = parseAiResponse(fullAiResponse);
    const japaneseInternalNotes = buildInternalNotesContent('jp', headers, rowData);
    const englishInternalNotes = buildInternalNotesContent('en', headers, rowData);
    const finalJapaneseContent = parsedJds.japanese + "\n\n---\n\n## フォーム回答内容\n\n" + japaneseInternalNotes;
    const finalEnglishContent = parsedJds.english + "\n\n---\n\n## Form Response Content\n\n" + englishInternalNotes;

    // --- 5. CREATE THE TWO GOOGLE DOCS ---
    let japaneseDocUrl = "";
    if (parsedJds.japanese) {
      const docTitle = `(JP) 【${japaneseJdTitle}】${jdDeptName}_${jdWorkLocation}`;
      japaneseDocUrl = createGoogleDoc(docTitle, finalJapaneseContent);
    }
    let englishDocUrl = "";
    if (parsedJds.english) {
      const docTitle = `(EN) - 【${englishJdTitle}】${jdDeptName}_${jdWorkLocation}`;
      englishDocUrl = createGoogleDoc(docTitle, finalEnglishContent);
    }
    
    // --- 6. WRITE LINKS BACK TO THE SHEET ---
    updateSheetWithLinks(sheet, rowIndex, headers, japaneseDocUrl, englishDocUrl);

  } catch (error) {
    Logger.log(`Error processing row ${rowIndex}: ${error.toString()}`);
    Logger.log(`Stack for row ${rowIndex}: ${error.stack}`);
  }
}

// =================================================================
// HELPER FUNCTIONS
// =================================================================

function getSheetData(headers, rowData, headerName, defaultValue = '') {
  const index = headers.indexOf(headerName);
  return index !== -1 && rowData[index] ? rowData[index] : defaultValue;
}

function buildJdPrompt(headers, rowData) {
  const rawInformationBlock = JD_COLUMN_HEADERS.map(header => {
      const value = getSheetData(headers, rowData, header, '記載なし');
      return `- ${header}: ${value}`;
    }).join('\n');

  if (!rawInformationBlock || rawInformationBlock.trim() === "") { return null; }

  const preferredSkillsBoilerplateEN = `Experience in AI development and/or experience in using AI tools to improve development processes.
Money Forward recently announced our AI Strategy roadmap which focuses on improving AI-driven operational efficiencies, as well as integrating AI agents into our products to deliver better value to our users. (More information here)
`;
  const preferredSkillsBoilerplateJP = `AIの開発経験もしくはAIツールを使用した開発経験
Money Forward AI Vision 2025にて発表の通り、マネーフォワードではAIを使った業務効率化に取り組んでいる状況かつ、将来的には全製品にAIエージェントを導入する想定であるため
`;

  const englishRequirementsBoilerplateEN = `(Note: If you have other qualifications or experiences demonstrating English proficiency, such as EIKEN Pre-1, EIKEN 2nd Grade (CSE score 1950+), TOEFL iBT 60+, IELTS 5.0+, or Cambridge FCE.), feel free to discuss with us) 
For those without a TOEIC 700+ equivalent score, they will be asked to take a designated test during the interview process (generally after the first interview)
`;
  const englishRequirementsBoilerplateJP = `※TOEIC以外にも英語力がわかる資格や経験をお持ちの方はご相談ください
例：英検準1級、英検2級（英検CSEスコア1950以上）、TOEFL iBT 60以上、IELTS 5.0以上、ケンブリッジ英語検定FCEなど
※その他、英語力がわかる資格や経験については応相談
※TOEIC 700点相当以上の資格をお持ちでない方については選考の過程で弊社指定の試験を受験いただきます。（原則、一次面接後を想定）
`;
  
  const prompt = `
I would like help creating a job description (JD) based on information from the hiring department.
Your goal is to create an English version and Japanese version of the JD based on their respective output formats below. 

### RAW INFORMATION
Here is the data provided by the hiring manager. Use this information to fill out the templates below.
${rawInformationBlock}

### YOUR TASK 
1. **Analyze**: Determine the primary language used in the RAW INFORMATION above. 
2. **Write**: Create the complete job description in that primary language first, following its template. 
3. **Translate**: Translate the version you just wrote into the other language, following its template. 
4. **Expand**: If any section has minimal information, expand it to be 3-4 professional sentences. 
5. **Omit**: If a section's information is blank, omit the entire section from the output. 
6. **Format the Technology Lists**: For the 'Technology Stack' and 'Tools Used' sections, take the unstructured information from the information and organize it into a categorized list. **The final output should use the same categories as the example shown in the templates below.** If the manager's notes don't mention a category, omit that category from the final list.


Final output format:

### ENGLISH OUTPUT TEMPLATE

Job Title 
{job_title_content} 

Background of the Recruitment 
{recruitment_background_content} 

Main Responsibilities 
{main_responsibilities_content} 

Job Satisfaction and Experience Gained 
{experience_gained_content} 

Expected Role 
{expected_role_content} 

Expected Mindset 
{expected_mindset_content} 

Desired Skills and Experience 
{desired_skills_content} 

Preferred Skills and Experience 
{preferred_skills_content} 
${preferredSkillsBoilerplateEN} 

Japanese Language Requirements 
{japanese_requirements_content} 

English Language Requirements 
{english_requirements_content} 
${englishRequirementsBoilerplateEN} 

We are looking for someone like this to join our team. 
{ideal_candidate_content} 

Technology Stack 
・Web Server-side：Java (Jersey, Guice, jOOQ) 
・Database：MySQL ・Middleware：Docker, Nginx, Consul 
・Platform：AWS, オンプレミス 
{additional_tech_stack_content} 

Tools Used 
・Repository Management ：GitHub ・CI/CD：CircleCI, Jenkins, Github Actions 
・Development Environment ：Docker, Terraform Enterprise 
・Monitoring ：DataDog, Rollbar, Sentry 
・Communication ：Slack 
・Security ：Dependabot {additional_tools_used_content} 

Reference URL {reference_url_content}


### JAPANESE OUTPUT TEMPLATE

職種タイトル
{job_title_content_jp}

募集背景
{recruitment_background_content_jp}

主な業務内容
{main_responsibilities_content_jp}

仕事のやりがい・得られる経験
{experience_gained_content_jp}

期待する役割
{expected_role_content_jp}

期待するマインド
{expected_mindset_content_jp}

求めるスキル・経験
{desired_skills_content_jp}

あると望ましいスキル・経験
{preferred_skills_content_jp}
${preferredSkillsBoilerplateJP}

日本語要件
{japanese_requirements_content_jp}

英語要件
{english_requirements_content_jp}
${englishRequirementsBoilerplateJP}

こんな方に仲間になってほしい
{ideal_candidate_content_jp}

技術スタック
・Webサーバーサイド：Java (Jersey, Guice, jOOQ)
・データベース：MySQL
・ミドルウェア：Docker, Nginx, Consul
・プラットフォーム：AWS, オンプレミス
{additional_tech_stack_content_jp}

使用ツール
・リポジトリ管理：GitHub
・CI/CD：CircleCI, Jenkins, Github Actions
・開発環境：Docker, Terraform Enterprise
・監視：DataDog, Rollbar, Sentry
・コミュニケーション：Slack
・セキュリティ：Dependabot
{additional_tools_used_content_jp}

参考URL
{reference_url_content_jp}
  `;
   return prompt;
}

function buildInternalNotesContent(language, headers, rowData) {
  let notesContent = "";
  headers.forEach((bilingualHeader, colIndex) => {
    const cellData = rowData[colIndex];
    if (cellData && cellData.toString().trim() !== '') {
      const headerParts = bilingualHeader.split(' / ');
      const jpHeader = headerParts[0];
      const enHeader = headerParts.length > 1 ? headerParts[1].trim() : jpHeader;
      const chosenHeader = (language === 'en') ? enHeader : jpHeader;
      notesContent += `${chosenHeader}:\n${cellData}\n\n`;
    }
  });
  return notesContent.trim() === "" ? "No internal notes provided." : notesContent;
}

function parseAiResponse(fullText) {
  const englishMarker = "### ENGLISH OUTPUT TEMPLATE";
  const japaneseMarker = "### JAPANESE OUTPUT TEMPLATE";
  const fallbackEnglishMarker = "### ENGLISH OUTPUT";
  const fallbackJapaneseMarker = "### JAPANESE OUTPUT";
  let englishContent = "";
  let japaneseContent = "";
  let englishStartIndex = fullText.indexOf(englishMarker);
  let actualEnglishMarker = englishMarker;
  if (englishStartIndex === -1) {
    englishStartIndex = fullText.indexOf(fallbackEnglishMarker);
    actualEnglishMarker = fallbackEnglishMarker;
  }
  let japaneseStartIndex = fullText.indexOf(japaneseMarker);
  let actualJapaneseMarker = japaneseMarker;
  if (japaneseStartIndex === -1) {
    japaneseStartIndex = fullText.indexOf(fallbackJapaneseMarker);
    actualJapaneseMarker = fallbackJapaneseMarker;
  }
  if (englishStartIndex !== -1) {
    const contentStartIndex = englishStartIndex + actualEnglishMarker.length;
    const contentEndIndex = (japaneseStartIndex !== -1 && japaneseStartIndex > englishStartIndex) ? japaneseStartIndex : fullText.length;
    englishContent = fullText.substring(contentStartIndex, contentEndIndex).trim();
  }
  if (japaneseStartIndex !== -1) {
    const contentStartIndex = japaneseStartIndex + actualJapaneseMarker.length;
    const contentEndIndex = (englishStartIndex !== -1 && englishStartIndex > japaneseStartIndex) ? englishStartIndex : fullText.length;
    japaneseContent = fullText.substring(contentStartIndex, contentEndIndex).trim();
  }
  return { english: englishContent, japanese: japaneseContent };
}

function createGoogleDoc(docTitle, content) {
  const doc = DocumentApp.create(docTitle);
  doc.getBody().appendParagraph(content);
  doc.saveAndClose();
  const docFile = DriveApp.getFileById(doc.getId());
  const folder = DriveApp.getFolderById(FOLDER_ID);
  folder.addFile(docFile);
  DriveApp.getRootFolder().removeFile(docFile);
  return doc.getUrl();
}

function updateSheetWithLinks(sheet, rowIndex, headers, japaneseDocUrl, englishDocUrl) {
    function findAndSetUrl(headerName, url) {
        if (!url) return;
        let colIndex = headers.indexOf(headerName);
        if (colIndex === -1) { 
            const newColIndex = sheet.getLastColumn() + 1;
            sheet.getRange(1, newColIndex).setValue(headerName).setFontWeight('bold');
            sheet.getRange(rowIndex, newColIndex).setValue(url);
        } else {
            sheet.getRange(rowIndex, colIndex + 1).setValue(url);
        }
    }
    findAndSetUrl(DOC_LINK_HEADER_JP, japaneseDocUrl);
    findAndSetUrl(DOC_LINK_HEADER_EN, englishDocUrl);
}

function getAiSummary(prompt) {
  const url = `${AZURE_APIM_BASE_URL}/openai/deployments/${AZURE_DEPLOYMENT_MODEL}/chat/completions?api-version=${AZURE_API_VERSION}`;
  const payload = { "model": AZURE_DEPLOYMENT_MODEL, "messages": [{ "role": "user", "content": prompt }]};
  const options = { 'method': 'post', 'contentType': 'application/json', 'headers': { 'api-key': AZURE_OPENAI_API_KEY }, 'payload': JSON.stringify(payload), 'muteHttpExceptions': true };
  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  const responseBody = response.getContentText();
  if (responseCode === 200) {
    const json = JSON.parse(responseBody);
    if (json.choices && json.choices.length > 0 && json.choices[0].message && json.choices[0].message.content) {
      return json.choices[0].message.content;
    } else { throw new Error('AI response is in an invalid format. Response: ' + responseBody); }
  } else { throw new Error(`Error calling Azure OpenAI API. Status: ${responseCode}, Response: ${responseBody}`); }
}

// =================================================================
// LEGACY FUNCTION (The old manual script)
// =================================================================

/**
 * This is the original script that combines multiple selected rows into one single Doc.
 * It is preserved here in case it's still needed.
 */
function exportToSingleDoc_legacy() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const folder = DriveApp.getFolderById(FOLDER_ID);
    
    const hasSelectedResponses = data.slice(1).some(row => row[1] === true);
    if (!hasSelectedResponses) {
      SpreadsheetApp.getUi().alert('選択された回答がありません。\nチェックボックスで回答を選択してください。');
      return;
    }

    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm');
    const doc = DocumentApp.create(`選択した回答_${timestamp}`);
    const body = doc.getBody();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === true) { // B列のチェックボックスがチェックされているか確認
        body.appendParagraph(`回答 ${i}`).setHeading(DocumentApp.ParagraphHeading.HEADING1);
        for (let j = 2; j < data[i].length; j++) {
          const content = data[i][j].toString().trim();
          if (content) {
            const questionPara = body.appendParagraph(`${headers[j]}:`);
            questionPara.setBold(true);
            body.appendParagraph(content).setIndentStart(20).setBold(false);
          }
        }
        body.appendParagraph('');
      }
    }
    doc.saveAndClose();
    
    const docFile = DriveApp.getFileById(doc.getId());
    folder.addFile(docFile);
    DriveApp.getRootFolder().removeFile(docFile);
    
    const url = doc.getUrl();
    SpreadsheetApp.getUi().alert('選択した回答を新しいドキュメントに書き出しました。\n\nドキュメントのURL: ' + url);

  } catch (error) {
    console.error('エラーが発生しました:', error);
    SpreadsheetApp.getUi().alert('エラーが発生しました。\n' + error.toString());
  }
}