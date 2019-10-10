/*
Grab all negative keywords within multiple negative keyword lists in Google Ads then
output all negative keywords to a column in a Google Sheet
*/

function main() {
  const SPREADSHEET_URL = 'YOUR SPREADSHEET URL' // input your google spreadsheet URL
  const SHEET_NAME = 'YOUR SHEET NAME' // the name of the sheet/tab within the google spreadsheet that you want to output all keywords to
  const sheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL).getSheetByName(SHEET_NAME); // access the google sheet

  const NEGATIVE_KEYWORD_LISTS = ['NEGATIVE_KEYWORD_LIST_1', 'NEGATIVE_KEYWORD_LIST_2', 'NEGATIVE_KEYWORD_LIST_3']; // input your negative keyword list names
  var allKeywords = getListKeys(NEGATIVE_KEYWORD_LISTS)

  addArrayToSheetColumn(sheet, 'A', allKeywords)  
}

// function below takes a an array of negative keyword lists as input (used in main) and gets all negative keywords within the supplied lists
function getListKeys(list) {
  var sharedNegativeKeywords = [];

  for (var i = 0; i < list.length; i++) {
    var negativeKeywordListIterator =
      AdsApp.negativeKeywordLists()
      .withCondition('Name = "' + list[i] + '"')
      .get();

    if (negativeKeywordListIterator.totalNumEntities() == 1) {
      var negativeKeywordList = negativeKeywordListIterator.next();
      var sharedNegativeKeywordIterator =
        negativeKeywordList.negativeKeywords().get();

      while (sharedNegativeKeywordIterator.hasNext()) {
        sharedNegativeKeywords.push(sharedNegativeKeywordIterator.next().getText());
      }
    }
  }
  return sharedNegativeKeywords
}

// transform the array of arrays outputted by getListKeys into a flattened array, and output to your Google Sheet column
// takes a sheet object, a column (in letter format: i.e. "A") to output to, and the values you want to output (the list of all negative keywords)
function addArrayToSheetColumn(sheet, column, values) {
  const range = [column, "1:", column, values.length].join("");
  const fn = function(value) {
    return [ value ];
  };
  sheet.getRange(range).setValues(values.map(fn));
}