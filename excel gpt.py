
// Twój klucz API OpenAI
const OPENAI_API_KEY = 'TU WPROWADŹ SWÓJ KLUCZ API'; // Zamień na swój rzeczywisty klucz API


/**
* Niestandardowa funkcja wywołująca GPT z komórki.
* @param {string} prompt Wprowadzenie dla GPT, może zawierać odniesienia do komórek.
* @param {number} maxTokens Maksymalna liczba tokenów dla odpowiedzi. Opcjonalne, domyślnie 450.
* @return Tekst wygenerowany przez GPT.
* @customfunction
*/
function GPT(prompt, maxTokens = 450) {
 // Pobierz aktywny arkusz kalkulacyjny i komórkę, która wywołała tę funkcję
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
 var cell = sheet.getActiveCell();
 // Przetwórz wprowadzenie, aby zastąpić odniesienia do komórek ich wartościami
 prompt = processCellReferences(prompt, sheet, cell);
 if (!prompt) {
   return "Błąd: Proszę podać wprowadzenie.";
 }
 const apiUrl = 'https://api.openai.com/v1/chat/completions';
 const payload = {
   'model': 'gpt-4o-mini',  // lub jakikolwiek model, którego używasz
   'messages': [
     {'role': 'system', 'content': 'Jesteś pomocnym asystentem.'},
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
   return "Błąd: " + error.toString();
 }
}


/**
* Przetwórz odniesienia do komórek w wprowadzeniu i zastąp je ich wartościami.
* @param {string} prompt Oryginalne wprowadzenie.
* @param {Sheet} sheet Aktywny arkusz.
* @param {Range} cell Komórka, która wywołała funkcję.
* @return {string} Przetworzone wprowadzenie z wartościami komórek.
*/
function processCellReferences(prompt, sheet, cell) {
 // Wyrażenie regularne do dopasowywania odniesień do komórek, takich jak A1, B2 itp.
 var cellRefRegex = /\b[A-Z]+\d+\b/g;
 return prompt.replace(cellRefRegex, function(match) {
   try {
     var value = sheet.getRange(match).getValue();
     return value.toString();
   } catch (e) {
     // Jeśli odniesienie do komórki jest nieprawidłowe, zwróć oryginalne dopasowanie
     return match;
   }
 });
}