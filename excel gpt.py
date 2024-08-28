
// Your OpenAI API key
const OPENAI_API_KEY = 'YOUR API KEY HERE'; // Replace with your actual API key


/**
* Custom function to call GPT from a cell.
* @param {string} prompt The input prompt for GPT, can include cell references.
* @param {number} maxTokens The maximum number of tokens for the response. Optional, default is 450.
* @return The generated text from GPT.
* @customfunction
*/
function GPT(prompt, maxTokens = 450) {
 // Get the active spreadsheet and the cell that called this function
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
 var cell = sheet.getActiveCell();
  // Process the prompt to replace cell references with their values
 prompt = processCellReferences(prompt, sheet, cell);
  if (!prompt) {
   return "Error: Please provide a prompt.";
 }
  const apiUrl = 'https://api.openai.com/v1/chat/completions';
  const payload = {
   'model': 'gpt-4o-mini',  // or whichever model you're using
   'messages': [
     {'role': 'system', 'content': 'You are a helpful assistant.'},
     {'role': 'user', 'content': prompt}
   ],
   'max_tokens': maxTokens
 };
  const options = {
   'method': 'post',
   'contentType': 'application/json',
   'headers': {
     'Authorization': 'Bearer ' + OPENAI_API_KEY
   },
   'payload': JSON.stringify(payload)
 };
  try {
   const response = UrlFetchApp.fetch(apiUrl, options);
   const json = JSON.parse(response.getContentText());
   return json.choices[0].message.content.trim();
 } catch (error) {
   return "Error: " + error.toString();
 }
}


/**
* Process cell references in the prompt and replace them with their values.
* @param {string} prompt The original prompt.
* @param {Sheet} sheet The active sheet.
* @param {Range} cell The cell that called the function.
* @return {string} The processed prompt with cell values.
*/
function processCellReferences(prompt, sheet, cell) {
 // Regular expression to match cell references like A1, B2, etc.
 var cellRefRegex = /\b[A-Z]+\d+\b/g;
  return prompt.replace(cellRefRegex, function(match) {
   try {
     var value = sheet.getRange(match).getValue();
     return value.toString();
   } catch (e) {
     // If the cell reference is invalid, return the original match
     return match;
   }
 });
}


