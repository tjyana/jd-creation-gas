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
    // Log errors with the specific row that failed
    Logger.log(`Error processing row ${rowIndex}: ${error.toString()}`);
    Logger.log(`Stack for row ${rowIndex}: ${error.stack}`);
    // Optionally, write an error message to a cell in that row
    // sheet.getRange(rowIndex, sheet.getLastColumn() + 1).setValue('ERROR: ' + error.message);
  }
}

// =================================================================
// HELPER FUNCTIONS (The rest of your script, mostly unchanged)
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

  const preferredSkillsBoilerplateEN = `Experience in AI development and/or experience in using AI tools to improve development processes...`; // Full text
  const preferredSkillsBoilerplateJP = `AIの開発経験もしくはAIツールを使用した開発経験...`; // Full text
  const englishRequirementsBoilerplateEN = `(Note: If you have other qualifications or experiences demonstrating English proficiency...)`; // Full text
  const englishRequirementsBoilerplateJP = `※TOEIC以外にも英語力がわかる資格や経験をお持ちの方はご相談ください...`; // Full text
  
  const prompt = `I would like help creating a job description...`; // The full, long prompt text goes here
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