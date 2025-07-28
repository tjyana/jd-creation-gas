// =================================================================
// CONFIGURATION & CONSTANTS
// =================================================================

const FOLDER_ID = '1_M44p2Cqv-1tdBJLsOS_3MtaRZ98IqhL'; // IMPORTANT: Make sure this is your correct folder ID
const AZURE_OPENAI_API_KEY = PropertiesService.getScriptProperties().getProperty('AZURE_OPENAI_API_KEY');
const AZURE_API_VERSION = '2024-07-01-preview';
const AZURE_DEPLOYMENT_MODEL = 'gpt-4o';
const AZURE_APIM_BASE_URL = "https://mf-genai-poc-apim.azure-api.net/esotad";

// --- Headers for the final documents ---
const DOC_LINK_HEADER_JP = "求人内容";
const DOC_LINK_HEADER_EN = "Document Link";
const INTERNAL_NOTES_LINK_HEADER = "Internal Notes Link";

// --- List of columns to include in the Job Description ---
const JD_COLUMN_HEADERS = [
  "職種タイトル / Job Title",
  "募集背景 /  Background of the Recruitment",
  "主な業務内容 /  Main Responsibilities",
  "仕事のやりがい・得られる経験 /  Job Satisfaction and Experience Gained",
  "期待する役割 /  Expected Role",
  "期待するマインド /  Expected Mindset",
  "求めるスキル・経験 /  Desired Skills and Experience",
  "あると望ましいスキル・経験 /  Preferred Skills and Experience",
  "日本語要件 /  Japanese Language Requirements",
  "英語要件 /  English Language Requirements",
  "こんな方に仲間になってほしい /  We are looking for someone like this to join our team.",
  "技術スタック /  Technology Stack",
  "使用ツール /  Tools Used",
  "参考URL /  Reference URL"
];

// --- List of columns to include in the Internal Notes document ---
const INTERNAL_NOTES_COLUMN_HEADERS = [
  "担当ハイヤリングマネージャー名 /  Name of the Responsible Hiring Manager",
  "担当リクルーター名 /  Name of the Responsible Recruiter",
  "職種 /  Job Title",
  "勤務形態 /  Employment Type",
  "勤務地 /  Work Location",
  "配属先部署名 /  Assigned Department Name",
  "募集の理由 /  Reason for Recruitment",
  "具体的な募集背景 /  Specific Background of the Recruitment",
  "採用納期 /  Hiring Deadline",
  "上記の理由/背景をご教示下さい /  Please provide the reason/background for the above.",
  "採用温度感 /  Urgency of Hiring",
  "想定グレード /  Expected Grade",
  "オファー年収のイメージ /  Estimated Annual Salary Offer",
  "年齢 / Ageお任せしたい業務 /  Tasks to be Assigned",
  "上記をお任せすることにあたって必要な経験・スキル /  Experience and Skills Required for the Above-mentioned Tasks",
  "ターゲット企業や業界 /  Target Companies and Industries",
  "技術課題の有無 /  Presence of Technical Challenges",
  "＜上記質問で「あり_track」を選択した方＞該当課題のURLを展開して下さい /  <For those who selected \"Present (Track)\" in the above question> Please provide the URL for the relevant challenge.",
  "＜上記質問で「あり_track以外」を選択した方＞課題を展開して下さい /  <For those who selected \"Present (Other than Track)\" in the above question> Please outline the challenge.",
  "技術課題レビュー担当者 /  Reviewer for Technical Challenges",
  "カジュアル面談 担当者 /  Casual Interview Representative",
  "カジュアル面談担当者の英語対応可否 /  English Proficiency of the Casual Interview Representative",
  "一次面接 担当者 /  First Interview Representative",
  "一次面接担当者の英語対応可否 /  English Proficiency of the First Interview Representative",
  "二次面接 担当者 /  Second Interview Representative",
  "二次面接担当者の英語対応可否 /  English Proficiency of the Second Interview Representative",
  "最終面接 担当者 /  Final Interview Representative",
  "最終面接担当者の英語対応可否 /  English Proficiency of the Final Interview Representative",
  "その他 / Otherオファー面談 担当者 /  Offer Meeting Representative",
  "エージェント利用可否 /  Availability of Agent Usage",
  "ビザサポートが必要な海外在住者に対してオープンしますか？ /  Are you open to candidates residing overseas who require visa support?"
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
  // const ui = SpreadsheetApp.getUi();
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

    const jdDeptNameHeader = "配属先部署名 /  Assigned Department Name";
    const jdDeptNameIndex = headers.indexOf(jdDeptNameHeader);
    const jdDeptName = jdDeptNameIndex !== -1 && rowData[jdDeptNameIndex] ? rowData[jdDeptNameIndex] : '不明な求人票';

    const jdWorkLocationHeader = "勤務地 /  Work Location";
    const jdWorkLocationIndex = headers.indexOf(jdWorkLocationHeader);
    const jdWorkLocation = jdWorkLocationIndex !== -1 && rowData[jdWorkLocationIndex] ? rowData[jdWorkLocationIndex] : '不明な勤務地';
    
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
      const docTitle = `【${jdTitle}】${jdDeptName}_${jdWorkLocation}`;
      jdDocUrl = createGoogleDoc(docTitle, summaryText);
    }

    // --- 4. CREATE INTERNAL NOTES DOC (if content exists) ---
    let internalNotesDocUrl = "";
    if (internalNotesContent) {
      const docTitle = `[Internal Notes] - 【${jdTitle}】${jdDeptName}_${jdWorkLocation}`;
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
  
  // --- Set variables ---
  const jobTitleHeader = "職種タイトル / Job Title";
  const jobTitleIndex = headers.indexOf(jobTitleHeader);
  const jobTitle = jobTitleIndex !== -1 && rowData[jobTitleIndex] ? rowData[jobTitleIndex] : '不明な求人票';

  const recruitmentBackgroundHeader = "募集背景 /  Background of the Recruitment";
  const recruitmentBackgroundIndex = headers.indexOf(recruitmentBackgroundHeader);
  const recruitmentBackground = rowData[recruitmentBackgroundIndex] || '記載なし';

  const mainResponsibilitiesHeader = "主な業務内容 /  Main Responsibilities";
  const mainResponsibilitiesIndex = headers.indexOf(mainResponsibilitiesHeader);
  const mainResponsibilities = rowData[mainResponsibilitiesIndex] || '記載なし';

  const experienceGainedHeader = "仕事のやりがい・得られる経験 /  Job Satisfaction and Experience Gained";
  const experienceGainedIndex = headers.indexOf(experienceGainedHeader);
  const experienceGained = rowData[experienceGainedIndex] || '記載なし';

  const expectedRoleHeader = "期待する役割 /  Expected Role";
  const expectedRoleIndex = headers.indexOf(expectedRoleHeader);
  const expectedRole = rowData[expectedRoleIndex] || '記載なし';

  const expectedMindsetHeader = "期待するマインド /  Expected Mindset";
  const expectedMindsetIndex = headers.indexOf(expectedMindsetHeader);
  const expectedMindset = rowData[expectedMindsetIndex] || '記載なし';

  const idealCandidateHeader = "こんな方に仲間になってほしい /  We are looking for someone like this to join our team.";
  const idealCandidateIndex = headers.indexOf(idealCandidateHeader);
  const idealCandidateProfile = rowData[idealCandidateIndex] || '記載なし';

  const requiredSkillsHeader = "求めるスキル・経験 /  Desired Skills and Experience";
  const requiredSkillsIndex = headers.indexOf(requiredSkillsHeader);
  const requiredSkills = rowData[requiredSkillsIndex] || '記載なし';

  const preferredSkillsHeader = "あると望ましいスキル・経験 /  Preferred Skills and Experience";
  const preferredSkillsIndex = headers.indexOf(preferredSkillsHeader);
  const preferredSkills = rowData[preferredSkillsIndex] || '記載なし';

  const japaneseRequirementsHeader = "日本語要件 /  Japanese Language Requirements";
  const japaneseRequirementsIndex = headers.indexOf(japaneseRequirementsHeader);
  const japaneseRequirements = rowData[japaneseRequirementsIndex] || '記載なし';

  const englishRequirementsHeader = "英語要件 /  English Language Requirements";
  const englishRequirementsIndex = headers.indexOf(englishRequirementsHeader);
  const englishRequirements = rowData[englishRequirementsIndex] || '記載なし';

  const techStackHeader = "技術スタック /  Technology Stack";
  const techStackIndex = headers.indexOf(techStackHeader);
  const techStack = rowData[techStackIndex] || '記載なし';

  const toolsUsedHeader = "使用ツール /  Tools Used";
  const toolsUsedIndex = headers.indexOf(toolsUsedHeader);
  const toolsUsed = rowData[toolsUsedIndex] || '記載なし';

  const referenceUrlHeader = "参考URL /  Reference URL";
  const referenceUrlIndex = headers.indexOf(referenceUrlHeader);
  const referenceUrl = rowData[referenceUrlIndex] || '記載なし';



  // --Define the raw info to be included in the JD--

  const rawInformationBlock = `
- 職種タイトル / Job Title: ${jobTitle}
- 募集背景 / Background of the Recruitment: ${recruitmentBackground}
- 主な業務内容 / Main Responsibilities: ${mainResponsibilities}
- 仕事のやりがい・得られる経験 / Job Satisfaction and Experience Gained: ${experienceGained}
- 期待する役割 / Expected Role: ${expectedRole}
- 期待するマインド / Expected Mindset: ${expectedMindset}
- 求めるスキル・経験 / Desired Skills and Experience: ${requiredSkills}
- あると望ましいスキル・経験 / Preferred Skills and Experience: ${preferredSkills}
- 日本語要件 / Japanese Language Requirements: ${japaneseRequirements}
- 英語要件 / English Language Requirements: ${englishRequirements}
- こんな方に仲間になってほしい / We are looking for someone like this to join our team.: ${idealCandidateProfile}
- 技術スタック / Technology Stack: ${techStack}
- 使用ツール / Tools Used: ${toolsUsed}
- 参考URL / Reference URL: ${referenceUrl}
`;


// Set the boilerplate text for preferred skills and English requirements
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
  
  // let interviewNotes = "";
  
  // Logger.log("--- Starting buildJdPrompt ---");
  // Logger.log(`Total headers found in sheet: ${headers.length}`);

  // // Instead of looping through all headers, we loop through our specific list
  // JD_COLUMN_HEADERS.forEach(headerName => {
  //   // LOG 1: What header are we looking for from our constant list?
  //   Logger.log(`Checking for header: '${headerName}'`);

  //   const colIndex = headers.indexOf(headerName);

  //   // LOG 2: Did we find it? Where?
  //   Logger.log(`Found at index: ${colIndex}`);

  //   // If we found it (index is not -1) and the cell has data...
  //   if (colIndex !== -1 && rowData[colIndex] && rowData[colIndex].toString().trim() !== '') {
  //     interviewNotes += `${headerName}:\n${rowData[colIndex]}\n\n`;
  //     // LOG 3: Log that we are adding this content.
  //     Logger.log(`SUCCESS: Found data for '${headerName}' and added it.`);
  //   }
  // });

  // Logger.log(`Final JD prompt content length: ${interviewNotes.length}`);
  // Logger.log("--- Finished buildJdPrompt ---");

  // if (interviewNotes.trim() === "") {
  //   return null;
  // }

  
  // Assume 'headers' and 'rowData' are variables available in your script.

  

  // The prompt template remains the same
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
  
  // SpreadsheetApp.getUi().alert('The trigger has been created! The script will now run automatically on new form submissions.');
}