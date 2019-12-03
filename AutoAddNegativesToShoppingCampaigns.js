/**
 *
 * This script will automatically add negative keywords to Google Ads Shopping Campaigns with specified match type (at the campaign level)
 * Positive keywords are used to assess which keywords we should NOT negate and live within a google spreadsheet (template provided)
 * Please adjust the positives list to your liking 
 *
 **/

// google spreadsheet details
// template spreadsheet for the positive keywords can be found here: 
// https://docs.google.com/spreadsheets/d/1JxnJEYA6mhQmpBCYoK-cqgbq2yAxMs1Ym6i0kxq0gio/edit?usp=sharing
var SPREADSHEET_URL = "YOUR_SPREADSHEET_URL"
var SS = SpreadsheetApp.openByUrl(SPREADSHEET_URL)
var sheet_name = 'YOUR_SHEET_NAME' || 'AutoNegs'

// who to send the email to for confirmation of the negatives implemented; emails are separated by commas
var EMAILS = 'YOUR_FIRST_EMAIL, YOUR_SECOND_EMAIL'

function main() {
  // change the sheet name to your google sheet's sheet name
  var sheet = SS.getSheetByName(sheet_name)
  var SETTINGS = {}

  // grab the setting in google sheet for the negative match type 
  SETTINGS["NEGATIVE_MATCH_TYPE"] = sheet.getRange("E2").getValue()

  // get positive keys from spreadsheet
  var keyword_col = 2
  var numMatchesRow = 4
  var firstKeywordRow = 6
  var rangeData = sheet.getDataRange()
  var lastRow = rangeData.getLastRow()
  var pos_key_range = sheet.getRange(firstKeywordRow, keyword_col, lastRow - firstKeywordRow + 1)
  var pos_keys = pos_key_range.getValues()
  var pos_key_rows = pos_key_range.getLastRow()
  var campMins = 1

  // get the search queries from the campaign; we will check the positive keywords against the search queries
  // queries live in a hidden sheet called 'queries'
  // to get search queries, we will have to get the data via a google script or other means (this script is not included here)
  var query_sheet_name = 'queries'
  var query_sheet = SS.getSheetByName(query_sheet_name)
  var query_sheet_lastRow = query_sheet.getLastRow()
  var query_range = query_sheet.getRange(1, 1, query_sheet_lastRow).getValues()

  var negs = [];
  var queries = []
  var pos_keys_list = []
  var matches_list = []

  for (var i = 1; i < query_sheet_lastRow; i++) {
    var q = query_range[i][0]
    queries.push(q)
    var matches = 0
    var count = 0

    // loop through the positive keywords
    for (r = 0; r < pos_key_rows - firstKeywordRow; r++) {
      var key = pos_keys[r][0]
      pos_keys.push(key)
      count++

      // if the keyword is in the query, we have a match. match++
      if (q.indexOf(key) > -1) {
        matches++;
      }

      // if we have reached the end of the positive keywords i.e. checked them all
      // and if the number of matches is less than the minimum number of matches for the campaign (specified on the sheet)
      // then add the query to the negatives array
      if (matches < 1 && count == pos_key_range.getNumRows() - 1) {
        negs.push(q);
        break;
      }
    }
  }

  log("Found a total of " + negs.length + " negative keywords to add")
  log("Adding the negative keywords...")

  // we have the negs. Now add them to the campaign...
  // please adjust you campaign naming
  // the campaignIterator can be generalized to non-shopping campaigns by using 'AdsApp.Campaigns()'
  var campaignIterator = AdsApp.shoppingCampaigns()
  var campaignIterator = campaignIterator
    .withCondition("Name CONTAINS 'YOUR_CAMPAIN_NAMING'")
    .withCondition("Impressions > 100")
    .forDateRange("LAST_7_DAYS")
    .get();

  while (campaignIterator.hasNext()) {
    var campaign = campaignIterator.next();

    for (var neg in negs) {
      var neg = addMatchType(negs[neg], SETTINGS)
      campaign.createNegativeKeyword(neg);
    }
    log("Campaign Name: " + campaign.getName() + " added " + negs.length + " negative keywords.")
  }
  log("Finished")
  var negs_output = "Negatives Added to Shopping Campaigns:\n" + negs.join('\n')
  log(negs_output)

  // send an email with the output
  sendEmail(negs_output)
}
//end main

function addMatchType(word, SETTINGS) {
  if (SETTINGS["NEGATIVE_MATCH_TYPE"].toLowerCase() == "broad") {
    word = word.trim();
  } else if (SETTINGS["NEGATIVE_MATCH_TYPE"].toLowerCase() == "bmm") {
    word = word.split(" ").map(function(x) {
      return "+" + x
    }).join(" ").trim()
  } else if (SETTINGS["NEGATIVE_MATCH_TYPE"].toLowerCase() == "phrase") {
    word = '"' + word.trim() + '"'
  } else if (SETTINGS["NEGATIVE_MATCH_TYPE"].toLowerCase() == "exact") {
    word = '[' + word.trim() + ']'
  } else {
    throw ("Error: Match type not recognised. Please provide one of Broad, BMM, Exact or Phrase")
  }
  return word;
}

function log(x) {
  Logger.log(x)
}

////////////////////
// send email with script log information
///////////////////

function sendEmail(body) {
  // get today's date
  var today = new Date()
  var date = today.getFullYear() + '-' + (today.getMonth() + 1) + '-' + today.getDate()
  
  var scriptName = 'Script: Auto Add Negatives for Shopping Campaigns'
  subject = scriptName + ' on ' + date.toString()

  MailApp.sendEmail(EMAILS,
    subject,
    body
  )
}