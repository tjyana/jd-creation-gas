// =================================================================
// CONFIGURATION & CONSTANTS
// =================================================================

const FOLDER_ID = '1_M44p2Cqv-1tdBJLsOS_3MtaRZ98IqhL'; // IMPORTANT: Make sure this is your correct folder ID
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

// --- List of columns to include in the Internal Notes section ---
const INTERNAL_NOTES_COLUMN_HEADERS = [
  "担当ハイヤリングマネージャー名 /  Name of the Responsible Hiring Manager", "担当リクルーター名 /  Name of the Responsible Recruiter",
  "職種 /  Job Title", "勤務形態 /  Employment Type", "勤務地 /  Work Location", "配属先部署名 /  Assigned Department Name",
  "募集の理由 /  Reason for Recruitment", "具体的な募集背景 /  Specific Background of the Recruitment", "採用納期 /  Hiring Deadline",
  "上記の理由/背景をご教示下さい /  Please provide the reason/background for the above.", "採用温度感 /  Urgency of Hiring",
  "想定グレード /  Expected Grade", "オファー年収のイメージ /  Estimated Annual Salary Offer", "年齢 / Ageお任せしたい業務 /  Tasks to be Assigned",
  "上記をお任せすることにあたって必要な経験・スキル /  Experience and Skills Required for the Above-mentioned Tasks",
  "ターゲット企業や業界 /  Target Companies and Industries", "技術課題の有無 /  Presence of Technical Challenges",
  "＜上記質問で「あり_track」を選択した方＞該当課題のURLを展開して下さい /  <For those who selected \"Present (Track)\" in the above question> Please provide the URL for the relevant challenge.",
  "＜上記質問で「あり_track以外」を選択した方＞課題を展開して下さい /  <For those who selected \"Present (Other than Track)\" in the above question> Please outline the challenge.",
  "技術課題レビュー担当者 /  Reviewer for Technical Challenges", "カジュアル面談 担当者 /  Casual Interview Representative",
  "カジュアル面談担当者の英語対応可否 /  English Proficiency of the Casual Interview Representative", "一次面接 担当者 /  First Interview Representative",
  "一次面接担当者の英語対応可否 /  English Proficiency of the First Interview Representative", "二次面接 担当者 /  Second Interview Representative",
  "二次面接担当者の英語対応可否 /  English Proficiency of the Second Interview Representative", "最終面接 担当者 /  Final Interview Representative",
  "最終面接担当者の英語対応可否 /  English Proficiency of the Final Interview Representative", "その他 / Otherオファー面談 担当者 /  Offer Meeting Representative",
  "エージェント利用可否 /  Availability of Agent Usage", "ビザサポートが必要な海外在住者に対してオープンしますか？ /  Are you open to candidates residing overseas who require visa support?"
];


// =================================================================
// MAIN TRIGGERED FUNCTION
// =================================================================

/**
 * Main function triggered by a form submission. Orchestrates the entire process.
 */
function processFormSubmission(e) {
  try {
    // --- 1. GET DATA FROM THE EVENT ---
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const submittedRange = e.range;
    const submittedRowIndex = submittedRange.getRowIndex();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowData = sheet.getRange(submittedRowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];

    // Get info for doc titles
    const jdTitle = getSheetData(headers, rowData, "職種タイトル / Job Title", '不明な求人');
    const jdDeptName = getSheetData(headers, rowData, "配属先部署名 /  Assigned Department Name", '不明な部署');
    const jdWorkLocation = getSheetData(headers, rowData, "勤務地 /  Work Location", '不明な勤務地');

    // --- 2. BUILD PROMPT & GET AI RESPONSE ---
    const jdPrompt = buildJdPrompt(headers, rowData);
    if (!jdPrompt) {
      Logger.log("Row " + submittedRowIndex + " appears to be empty. Skipping.");
      return;
    }
    const fullAiResponse = getAiSummary(jdPrompt);

    // --- 3. PARSE AND BUILD FINAL CONTENT ---
    const parsedJds = parseAiResponse(fullAiResponse);

    const japaneseInternalNotes = buildInternalNotesContent('jp', headers, rowData);
    const englishInternalNotes = buildInternalNotesContent('en', headers, rowData);

    const finalJapaneseContent = parsedJds.japanese + "\n\n---\n\n## Internal Recruiter Notes\n\n" + japaneseInternalNotes;
    const finalEnglishContent = parsedJds.english + "\n\n---\n\n## Internal Recruiter Notes\n\n" + englishInternalNotes;

    // --- 4. CREATE THE TWO GOOGLE DOCS ---
    let japaneseDocUrl = "";
    if (parsedJds.japanese) {
      const docTitle = `(JP) 【${jdTitle}】${jdDeptName}_${jdWorkLocation}`;
      japaneseDocUrl = createGoogleDoc(docTitle, finalJapaneseContent);
    }

    let englishDocUrl = "";
    if (parsedJds.english) {
      const docTitle = `(EN) - 【${jdTitle}】${jdDeptName}_${jdWorkLocation}`;
      englishDocUrl = createGoogleDoc(docTitle, finalEnglishContent);
    }
    
    // --- 5. WRITE LINKS BACK TO SHEET ---
    updateSheetWithLinks(sheet, submittedRowIndex, headers, japaneseDocUrl, englishDocUrl);

  } catch (error) {
    Logger.log("Error in processFormSubmission: " + error.toString());
    Logger.log("Stack: " + error.stack);
  }
}


// =================================================================
// LOGIC & CONTENT-BUILDING FUNCTIONS
// =================================================================

/**
 * Gets data from a specific column in the row safely.
 */
function getSheetData(headers, rowData, headerName, defaultValue = '') {
  const index = headers.indexOf(headerName);
  return index !== -1 && rowData[index] ? rowData[index] : defaultValue;
}


/**
 * Builds the AI prompt.
 */
function buildJdPrompt(headers, rowData) {
  const rawInformationBlock = JD_COLUMN_HEADERS.map(header => {
      const value = getSheetData(headers, rowData, header, '記載なし');
      return `- ${header}: ${value}`;
    }).join('\n');

  if (!rawInformationBlock || rawInformationBlock.trim() === "") {
    return null; 
  }

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

/**
 * MODIFIED: This function now iterates through ALL columns in the sheet to build the notes,
 * instead of a predefined list. It still creates language-specific versions.
 */
function buildInternalNotesContent(language, headers, rowData) {
  let notesContent = "";

  // Loop through every header in the sheet. 'colIndex' is the position (0, 1, 2...).
  headers.forEach((bilingualHeader, colIndex) => {
    
    // Get the data from the cell in the same position.
    const cellData = rowData[colIndex];

    // Check if the cell actually has data in it before adding it to the notes.
    if (cellData && cellData.toString().trim() !== '') {
      
      // Split the bilingual header like "日本語ヘッダー / English Header"
      const headerParts = bilingualHeader.split(' / ');
      const jpHeader = headerParts[0];
      // If there is an English part, use it; otherwise, just use the Japanese part.
      const enHeader = headerParts.length > 1 ? headerParts[1].trim() : jpHeader;

      // Pick the correct header based on the 'language' parameter ('jp' or 'en').
      const chosenHeader = (language === 'en') ? enHeader : jpHeader;
      
      // Add the formatted line to our notes string.
      notesContent += `${chosenHeader}:\n${cellData}\n\n`;
    }
  });

  return notesContent.trim() === "" ? "No internal notes provided." : notesContent;
}

/**
 * Parses the AI's response into separate Japanese and English strings.
 */
function parseAiResponse(fullText) {
  const englishMarker = "### ENGLISH OUTPUT TEMPLATE";
  const japaneseMarker = "### JAPANESE OUTPUT TEMPLATE";
  
  let englishContent = "";
  let japaneseContent = "";
  
  const japaneseStartIndex = fullText.indexOf(japaneseMarker);
  const englishStartIndex = fullText.indexOf(englishMarker);

  if (englishStartIndex !== -1) {
    const endOfEnglishIndex = (japaneseStartIndex > englishStartIndex) ? japaneseStartIndex : fullText.length;
    englishContent = fullText.substring(englishStartIndex + englishMarker.length, endOfEnglishIndex).trim();
  }

  if (japaneseStartIndex !== -1) {
    const endOfJapaneseIndex = (englishStartIndex > japaneseStartIndex) ? englishStartIndex : fullText.length;
    japaneseContent = fullText.substring(japaneseStartIndex + japaneseMarker.length, endOfJapaneseIndex).trim();
  }
  
  if (!englishContent && !japaneseContent) {
     const fallbackSplit = fullText.split('### JAPANESE OUTPUT');
     if (fallbackSplit.length > 1) {
       englishContent = fallbackSplit[0].replace('### ENGLISH OUTPUT','').trim();
       japaneseContent = fallbackSplit[1].trim();
     }
  }

  return {
    english: englishContent,
    japanese: japaneseContent
  };
}


// =================================================================
// GOOGLE & AZURE SERVICE FUNCTIONS
// =================================================================

/**
 * Creates a Google Doc, moves it to the correct folder, and returns the URL.
 */
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


/**
 * Updates the sheet with two separate document URLs.
 */
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


/**
 * Calls the Azure OpenAI API to get an AI-generated summary.
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
      throw new Error('AI response is in an invalid format. Response: ' + responseBody);
    }
  } else {
    throw new Error(`Error calling Azure OpenAI API. Status: ${responseCode}, Response: ${responseBody}`);
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
}