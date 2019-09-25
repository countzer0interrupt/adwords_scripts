//-----------------------------------------------------------------------------------------------------------------------------------------------
// BRAND SUITABILITY KPI SCRIPT v1.3.2
// Instructions
// If you have a time out (> 30 mins) then please re-run the script, it will skip over any campaigns remediated in the previous 3 days. If you 
// re-run the script after this, it will continue to loop through all campaigns.
// Please look at the main() function, and verify that the URL variable is set to your client (STEP ONE)
// If you are running the script at MCC level, ensure that 'LEVEL = "Account"' is commented out with two forward slashes, or is deleted.
// If you are running the script at Account level, ensure that 'LEVEL = Account' is present, otherwise the script will fail (STEP TWO)
// ----------------------------------------------------------------------------------------------------------------------------------------------

// CLIENT URLs: Please do not edit these 
var COKE_URL = 'https://docs.google.com/spreadsheets/d/1HNBVG7IlKVLvKtSa7C7sfeohTgkO5DFE6aBfNfEQzY4/';
var DIAGEO_URL = 'https://docs.google.com/spreadsheets/d/1XYtlSXnHIlECgGkmaUlC3EZsKGt3j46n-Sh5VJzoFnY/';
var PROCTOR_URL = 'https://docs.google.com/spreadsheets/d/1o_H5MeJfsnxuxiSsdgIzDo5aPxRX2QnJ5APBeYRrQr8/';
var HNK_URL = 'https://docs.google.com/spreadsheets/d/1UzHLyCRtPhat462C8BZolLi2zoJYJwxKID9LsR4F5B4/';
var MICROSOFT_URL = 'https://docs.google.com/spreadsheets/d/1mhPy_EK89uevQ-oHEAvK88vBuNy_62_xQAO8wg9O9YU/';
var MONDELEZ_URL = 'https://docs.google.com/spreadsheets/d/1YNqG9e1X73ufj2QMT2BnyoQA0zJdpOdua0LKxxOIx8k/';
var BMW_URL = 'https://docs.google.com/spreadsheets/d/1vDYE8TqKG7z-jPuS9Eiu-RTQMJJyOPntw19g2mY4_K0/';
var LVMH_URL = 'https://docs.google.com/spreadsheets/d/1_fgEVphua2nQLAScOkoQ2_48LGF2RU43u2VGHvFMEqE/';
var LOREAL_URL = 'https://docs.google.com/spreadsheets/d/10wZ_K5mrMMtInrILFviQVsM1kUIyix6lk2QzLRm3ZhA/';
var LEGO_URL = 'https://docs.google.com/spreadsheets/d/1nKd46O-nCYxIozF9u3LGvGBR7HsCKcwNYnMZF9hSwJA/'; 


// LEAVE: Please do not edit these URLs
var TOPICLIST_URL = 'https://docs.google.com/spreadsheets/d/1zLKscdaHQTtRINjUvbinTxMKs2LfthBl_jpBjfoLiZo/';
var CONTENTTYPE_URL = 'https://docs.google.com/spreadsheets/d/1ap27whVCfA638S858XqXqfHD2rKvK93LEf4F-z6sn58/';
var LOG_URL = 'https://docs.google.com/spreadsheets/d/1x7D-WjdSd4klm_eJzs4Y7eX_npK51A6n5ZUqJ8VGHUc/';

// other constants, do not edit this
var LEVEL = "MCC";
var location;
var log;

function main() {

  // STEP ONE - SELECT YOUR CLIENT URL HERE: It is important you select your client here (eg. "COKE_URL, or DIAGEO_URL")
  var URL = #INPUT URL HERE#
 
  log = openSpreadSheet(LOG_URL);
  // STEP TWO - BY DEFAULT THIS SCRIPT ASSUMES YOU ARE RUNNING AT MCC LEVEL, IF IT IS AT ACOCUNT LEVEL, PLEASE UNCOMMENT THE LINK BELOW
  LEVEL = "Account"
  

  if(LEVEL=="Account") { account = AdsApp.currentAccount(); run(URL, account ); appendRow(log, account.getCustomerId(), account.getCustomerId(), d.toString(), "ACCTCOMPLETED") }
  else {
    
    var accountIterator = AdsManagerApp.accounts().get();
    
    while (accountIterator.hasNext()) {
      var account = accountIterator.next();
      
      a_id = rvlookup(log.getSheets()[0], 1, 4, account.getCustomerId());
      if(String(a_id.length)>0) {
      if((a_id)==account.getCustomerId()) { continue; } else { a_id = ""; }}
      
      // Select the client account.
      AdsManagerApp.select(account);
      run(URL, account);
      appendRow(log, account.getCustomerId(), account.getCustomerId(), d.toString(), "ACCTCOMPLETED", account.getCustomerId());
    }
  }
}

function run(url, current) {
  var ss = openSpreadSheet(url); 
  // set up spreadsheets
  d = new Date();
  
  
  // iterate through video + standard campaigns
  iterators = [AdsApp.videoCampaigns().get(), AdsApp.campaigns().get()];

    c_id = rvlookup(log.getSheets()[0], 1, 4, current.getCustomerId());
    
  for(t=0; t<iterators.length; t++) {
    iterator = iterators[t];
    while (iterator.hasNext()) {
    var campaign = iterator.next();
    
    if(String(c_id).length > 0) {
    if((+c_id)!==campaign.getId()) { continue; } else { c_id = ""; }}
    
    if(hasActiveAds(campaign)) {
      appendRow(log, current.getCustomerId(), current.getName(), d.toString(), "CAMPAIGNSTART", campaign.getId());
      var locations = campaign.targeting().targetedLocations().get();
      if(locations.totalNumEntities() > 1 || locations.totalNumEntities() == 0) { 
        Logger.Log("Please note that campaign " + campaign.getName() + " has " + locations.totalNumEntities() + " locations. This will be ineligible for the Brand Suitability KPI, which requires exactly 1.");
        continue;
      } else {
        location = locations.next().getCountryCode();
      }
    var excludedTopicsArr = findExcludedTopics(ss);
    var contentTypesArr = findContentTypes(ss);
    var contentLabelsArr = findContentLabels(ss);
    var miscArr = findMisc(ss);
      excl_array = findTopicByNameArray(excludedTopicsArr); // need to find a way to get exclusions
      addExcludedTopicsToCampaign(campaign, excl_array); // if we do at camp level, consider ad group level
      excl_array = findExcludedContentLabelByNameArray(contentTypesArr, campaign);
      addExcludedContentLabelsToCampaign(campaign, excl_array); // if we do at camp level, consider ad group level
      checkKeywordExclusions(campaign);
      checkVideoChannelExclusions(campaign);
      checkSiteExclusions(campaign);
    }
      appendRow(log, current.getCustomerId(), current.getName(), d.toString(), "CAMPAIGNCOMPLETE", campaign.getId());
  }
  }
    
}

function findTopicByNameArray(arr) {
  retVal = [];
  ss = openSpreadSheet(TOPICLIST_URL)

for(k=0; k<arr.length; k++) {
  val = vlookup(ss.getActiveSheet(), 1, 1, arr[k][0]);
  if(val!==null && typeof val !== 'undefined' && arr[k][1] == true) { retVal.push(val); } //wtf is the === negation?
  else { Logger.log("Error: The topic " + arr[k][0] + " is not a valid setting / doesn't exist in the Topics list."); }
}
  return retVal;
}

function findExcludedContentLabelByNameArray(arr, campaign) {
  retVal = [];
  var campaignType = campaign.getEntityType();
  var offset = 0;

  if(campaignType == "Campaign") {
  targeting = campaign.display();
} else if (campaignType == "VideoCampaign") {
  targeting = campaign.videoTargeting();
  offset = 2;
}
  ss = openSpreadSheet(CONTENTTYPE_URL)

for(k=0; k<arr.length; k++) {
  val = vlookup(ss.getActiveSheet(), 1+offset, 1, arr[k][0]);
  if(val!==null && typeof val !== 'undefined') { retVal.push(val); } //wtf is the === negation?
  else { Logger.log("Error: The topic " + arr[k][0] + " is not a valid setting / doesn't exist in the Topics list."); }
}
  
  return retVal;
}

function addExcludedTopicsToCampaign(campaign, excl_array) {
var targeting;
var campaignType = campaign.getEntityType();
  
  if(campaignType == "Campaign") {
  targeting = campaign.display();
} else if (campaignType == "VideoCampaign") {
  targeting = campaign.videoTargeting();
}
  
for(i=0; i<excl_array.length;i++) {
        try {
        var topicBuilder = targeting.newTopicBuilder();
        var topic = topicBuilder.withTopicId(excl_array[i]).exclude();
        Logger.log("Exclusion Topic Added: " + excl_array[i]);
        }
        catch(e) {
        Logger.log("Error adding topic to campaign. It may be that it does not accept Topic Exclusions.")
        }
      }
}

function addExcludedContentLabelsToCampaign(campaign, excl_array) {
  
for(i=0; i<excl_array.length;i++) {
        try {
        var topicBuilder = campaign.excludeContentLabel(excl_array[i]);
        Logger.log("Content Label Exclusion Added: " + excl_array[i]);
        }
        catch (e) {
        Logger.log("Unspecified error with content label exclusion. It may be : " + e);
          
        }
      }
}

function checkKeywordExclusions(campaign) {
    var campaignType = campaign.getEntityType();
  return; // below commented out for time being, until further direction from TT/TG
  if(campaignType == "Campaign") {
  targeting = campaign.display();
    if(campaign.negativeKeywordLists().get().totalNumEntities() <= 0 || campaign.negativeKeywords().get().totalNumEntities() <= 0) {
    // add a random keyword - CONFIRM WITH TAMMY AND TAZIO THAT THIS IS OK
      campaign.createNegativeKeyword("porn");
    }
} else if (campaignType == "VideoCampaign") {
  targeting = campaign.videoTargeting();
  videoKeywordBuilder = targeting.newKeywordBuilder();
  if(targeting.keywords().get().totalNumEntities() <= 0 || campaign.negativeKeywordsLists().get().totalNumEntities() <= 0) {
  var videoKeywordOperation = videoKeywordBuilder
   .withText('porn')     // required
   .exclude();                // create the keyword
}       
}
}

function checkVideoChannelExclusions(campaign) {
    // check with tammy and tazio
      var campaignType = campaign.getEntityType();
      return; // below commented out for time being, until further direction from TT/TG
  if(campaignType == "Campaign") {
  targeting = campaign.display();
} else if (campaignType == "VideoCampaign") {
  targeting = campaign.videoTargeting();
}  
 
  if(targeting.excludedYouTubeVideos().get().totalNumEntities() <= 0 && targeting.excludedYouTubeChannels().get().totalNumEntities() <= 0) {
     var youTubeVideoBuilder = targeting.newYouTubeVideoBuilder();
     var youTubeVideoOperation = youTubeVideoBuilder
     .withVideoId('r5DYWxehYLc')      // required
     .exclude(); 
    }
      

}


function checkSiteExclusions(campaign) {
    // check with tazio and tammy
    return; // below commented out for time being, until further direction from TT/TG
    var campaignType = campaign.getEntityType();
      if(campaignType == "Campaign") {
  targeting = campaign.display();
} else if (campaignType == "VideoCampaign") {
  targeting = campaign.videoTargeting();
}  
  
    if(targeting.excludedPlacements().get().totalNumEntities() <= 0) {
      var placementBuilder = targeting.newPlacementBuilder()
     .withUrl("http://www.kpex.co.nz") 
     .exclude()    
    }
      
}

// verify that content types are the exposed via the display() interface
function addContentTypesToDisplay(display, excl_array) {
for(i=0; i<excl_array.length;i++) {
        var topicBuilder = display.newTopicBuilder();
        var topic = topicBuilder.withTopicId(excl_array[i]).exclude();
        Logger.log("Exclusion Topic Added: " + topic.getResult());
      }
}

function getAllCampaigns() {
  // AdsApp.campaigns() will return all campaigns that are not removed by
  // default.
  return AdsApp.campaigns().get();
}

function hasActiveAds(campaign) {
  var retVal = false;
  var adGroupIterator;
  if(campaign.getEntityType() == "Campaign") {
  adGroupIterator = campaign.ads().get();
  } else if(campaign.getEntityType() == "VideoCampaign") { 
    adGroupIterator = campaign.videoAds().get() };
  
  
  while (adGroupIterator.hasNext()) {
    var adGroup = adGroupIterator.next();
    if(adGroup.isEnabled()) { retVal = true; break; }
  }
  return retVal;
}

function doTopicExclusionsAtCampaign(campaign) {
   var displaySettings = campaign.display();
   topicIterator = displaySettings.excludedTopics().get();
   while (topicIterator.hasNext()) {
    var topic = topicIterator.next();
    //Logger.log(topic.getTopicId() + " " + topic.isEnabled());
  }
}

function doTopicExclusions(displaySelector) {
  var topicBuilder = displaySelector.newTopicBuilder();
  var topic = topicBuilder.withTopicId(50).exclude();
  //Logger.log(topic.getResult());
}


//////////// SPREADSHEET 
function findExcludedTopics(ss, index) {
return findOptions(ss, "Topic Exclusions", index);
}
function findContentTypes(ss, index) {
a = findOptions(ss, "Content Types", index);
b = findOptions(ss, "Sensitive Subjects", index);
return a.concat(b)
}
function findContentLabels(ss, index) {
return findOptions(ss, "Content Labels", index);
}
function findMisc(ss, index) {
return findOptions(ss, "Misc", index);
}
function findLineups(ss, index) {
return findOptions(ss, "Lineups", index);
}


function findOptions(ss, text, sheet_index) {
    if(sheet_index==null) sheet_index = 0;
    retArr = [];
    start = findValueInRow(ss.getSheets()[0], text);
    if(retVal !== -1) {
      rng = ss.getSheets()[0].getRange(2, start);
      if(rng.isPartOfMerge()) {
      var merger = rng.getMergedRanges()[0];
      var last = merger.getLastColumn();
      }
      for(i=start; i <= last; i++) {
      var val = ss.getSheets()[0].getRange(3, i).getValue();
      var index = findValueInColumn(ss.getSheets()[0],location);
      var flag = ss.getSheets()[0].getRange(index, i).getValue();
      retArr.push([val,flag]);
      }
    }
    return retArr;
  }

  function vlookup(sheet, column, index, value) {
  var retVal; 
  var lastRow=sheet.getLastRow();
  var data=sheet.getRange(1,column,lastRow,column+index).getValues();
  
  for(i=0;i<data.length;++i){
 
    if (data[i][0]==value){
      retVal = data[i][index];
      break;
    }
  }
    return retVal;
}

function rvlookup(sheet, column, index, value) {
  var retVal = "";
  var lastRow=sheet.getLastRow();
  var data=sheet.getRange(1,column,lastRow,column+index).getValues();
  
  for(i=data.length-1;i>=0;i--){
    
    if (data[i][0]==value){
      retVal = data[i][index];
      break;
    }
  }
  
    return retVal;
}

  function findValueInRow(sheet, value) {
    retVal = -1;
    var values = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues();
    for(var i=0, iLen=sheet.getLastColumn(); i<iLen; i++) {
      if(values[0][i] == value) {
        retVal = i+1;
        break;
      }
    }
    return retVal;
  }
  
function findValueInColumn(sheet, value) {
  retVal = -1;
  var values = sheet.getRange(1, 3, sheet.getLastRow(), 1).getValues();
  
  for(var i=0, iLen=sheet.getLastRow(); i<iLen; i++) {
    if(values[i][0] == value) {
      retVal = i+1;
      break;
    }
  }
  return retVal;
}

function openSpreadSheet(url) {
  // The code below opens a spreadsheet using its URL and logs the name for it.
  // Note that the spreadsheet is NOT physically opened on the client side.
  // It is opened on the server only (for modification by the script).
  var ss = SpreadsheetApp.openByUrl(url);
  return ss;
}


function appendRow(ss) {
  var sheet = ss.getSheets()[0];
  var current = AdsApp.currentAccount();
  var arr = [];
  if(arguments.length>1) {
    for(i=1; i<arguments.length; i++) {
   arr.push(arguments[i]);
    }
  }
  // Appends a new row with 3 columns to the bottom of the
  // spreadsheet containing the values in the array.
  
  sheet.appendRow(arr);
}