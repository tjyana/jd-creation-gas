// docs.test.js

// --- MOCK SETUP ---
// Create a fake version of the PropertiesService for our test environment.
const mockPropertiesService = {
  getScriptProperties: () => ({
    // This fake function will be called by your script.
    // It just needs to return a dummy value so the script doesn't crash.
    getProperty: () => 'fake_api_key_for_testing'
  }),
};

// Assign the fake service to the global scope, so it's available
// when Jest runs the script file.
global.PropertiesService = mockPropertiesService;
// --- END MOCK SETUP ---


// 1. Import the function you want to test from your main script file.
// Make sure './Code.js' matches the name of your script file.
const { buildJdPrompt } = require('./docs.js'); 

// 2. The 'describe' block groups related tests together. It's for organization.
describe('buildJdPrompt', () => {

  // 3. The 'test' or 'it' block is an individual test case.
  // Describe what this specific test should do.
  test('should create a complete prompt from header and row data', () => {
    
    // ARRANGE: Set up your test data.
    const mockHeaders = ['職種タイトル / Job Title', '主な業務内容 /  Main Responsibilities', 'ドキュメントリンク'];
    const mockRowData = ['AI Prompt Engineer', 'Write excellent prompts for an AI model.', 'http://some-link.com'];

    // ACT: Call the function with your test data.
    const resultPrompt = buildJdPrompt(mockHeaders, mockRowData);

    // ASSERT: Check if the result is what you expect.
    // We expect it to include the title and responsibilities.
    expect(resultPrompt).toContain('職種タイトル / Job Title:\nAI Prompt Engineer');
    expect(resultPrompt).toContain('主な業務内容 /  Main Responsibilities:\nWrite excellent prompts for an AI model.');
  });
  
  test('should ignore columns like "ドキュメントリンク"', () => {
    const mockHeaders = ['職種タイトル / Job Title', 'ドキュメントリンク'];
    const mockRowData = ['AI Prompt Engineer', 'http://some-link.com'];
    
    const resultPrompt = buildJdPrompt(mockHeaders, mockRowData);

    // We expect the result to NOT include the Document Link text.
    expect(resultPrompt).not.toContain('ドキュメントリンク');
  });

  test('should return null if all input data is empty or ignored', () => {
    const mockHeaders = ['職種タイトル / Job Title', 'ドキュメントリンク'];
    const mockRowData = ['', '']; // All data is blank
    
    const resultPrompt = buildJdPrompt(mockHeaders, mockRowData);
    
    // We expect the function to return null for empty input.
    expect(resultPrompt).toBeNull();
  });

});