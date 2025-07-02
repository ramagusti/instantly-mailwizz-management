function getInstantlyWarmupStats() {
  try {
    fetchEmailAccounts();
    fetchWarmupAnalytics();
  } catch (e) {
    Logger.log('Error in getInstantlyWarmupStats(): ' + e.toString());
    
    // Remove any existing retry triggers for getInstantlyWarmupStats
    ScriptApp.getProjectTriggers().forEach(function(trig) {
      if (trig.getHandlerFunction() === 'getInstantlyWarmupStats') {
        ScriptApp.deleteTrigger(trig);
      }
    });
    
    // Schedule a one-time retry in 1 minute
    ScriptApp.newTrigger('getInstantlyWarmupStats')
      .timeBased()
      .after(1 * 60 * 1000)   // 1 minute in milliseconds
      .create();
  }
}

function fetchEmailAccounts() {
  // Getting email list, status, and health score
  var apiUrl = "https://api.instantly.ai/api/v2/accounts?limit=100";
  
  var headers = {
    "Authorization": "Bearer " + instantlyApiKey
  };
  
  var options = {
    "method": "get",
    "headers": headers,
    "muteHttpExceptions": true
  };
  
  var statusDict = {
    "1": "Active",
    "2": "Paused",
    "-1": "Connection Error",
    "-2": "Soft Bounce Error",
    "-3": "Sending Error"
  };
  
  var existingData = emailSheet.getDataRange().getValues();
  for (var i = 2; i < existingData.length; i++) {
    emailIndex[existingData[i][0]] = i + 1;
  }
  
  var nextStartingAfter = null;
  do {
    var url = apiUrl;
    if (nextStartingAfter) {
      url += "&starting_after=" + nextStartingAfter;
    }
    
    var response = UrlFetchApp.fetch(url, options);
    var json = JSON.parse(response.getContentText());
    
    if (json.items && json.items.length > 0) {
      json.items.forEach(function(item) {
        var email = item.email;
        var status = statusDict[item.status] || "Unknown";
        var warmupScore = item.stat_warmup_score / 100;
        
        if (emailIndex[email] !== undefined) {
          var row = emailIndex[email];
          emailSheet.getRange(row, 6).setValue(status);
          // emailSheet.getRange(row, 3).setValue(warmupScore);
        } else {
          var newRow = emailSheet.getLastRow() + 1;
          emailSheet.appendRow([email, "", "", "", "", status]);
          emailIndex[email] = newRow;
        }
      });
    }
    
    nextStartingAfter = json.next_starting_after || null;
  } while (nextStartingAfter);
}

function fetchWarmupAnalytics() {
  var apiUrl = "https://api.instantly.ai/api/v2/accounts/warmup-analytics";

  var existingData = emailSheet.getDataRange().getValues();
  var emails = [];
  for (var i = 1; i < existingData.length; i++) {
    if (existingData[i][0]) {
      emails.push(existingData[i][0]);
    }
  }

  function chunkArray(array, size) {
    var results = [];
    for (var i = 0; i < array.length; i += size) {
      results.push(array.slice(i, i + size));
    }
    return results;
  }

  var emailChunks = chunkArray(emails, 100);

  for (var c = 0; c < emailChunks.length; c++) {
    var chunk = emailChunks[c];
    var payload = JSON.stringify({ "emails": chunk });
    var headers = {
      "Authorization": "Bearer " + instantlyApiKey,
      "Content-Type": "application/json"
    };
    var options = {
      "method": "post",
      "headers": headers,
      "payload": payload,
      "muteHttpExceptions": true
    };

    try {
      var response = UrlFetchApp.fetch(apiUrl, options);
      var json = JSON.parse(response.getContentText());

      var emailData = json.email_date_data || {};
      var aggregateData = json.aggregate_data || {};

      for (var i = 1; i < existingData.length; i++) {
        var email = existingData[i][0];
        if (!email || !chunk.includes(email)) continue;

        if (!emailData[email]) {
          Logger.log("No data for email: " + email);
          continue;
        }

        var sortedDates = Object.keys(emailData[email]).sort().reverse();
        var scores = [];

        for (var j = 0; j < sortedDates.length; j++) {
          var date = sortedDates[j];
          var data = emailData[email][date];
          if (data.sent && data.landed_inbox) {
            scores.push({
              date: date,
              score: data.landed_inbox / data.sent
            });
          }
        }

        var last3Avg = scores.slice(0, 3).reduce((sum, item) => sum + item.score, 0) / Math.min(scores.length, 3) || "";

        emailSheet.getRange(i + 1, 8).setValue(last3Avg);
        for (var k = 0; k < 7; k++) {
          if (scores[k]) {
            emailSheet.getRange(i + 1, 9 + k * 2).setValue(scores[k].date);
            emailSheet.getRange(i + 1, 10 + k * 2).setValue(scores[k].score);
          } else {
            emailSheet.getRange(i + 1, 9 + k * 2).setValue("");
            emailSheet.getRange(i + 1, 10 + k * 2).setValue("");
          }
        }
      }
    } catch (error) {
      Logger.log("API call failed for batch " + (c + 1) + ": " + error);
    }
  }
}

function postEmailManagementData() {
  updateInstantlyWarmupQuota();
  updateCoSendingQuota();
}

function updateInstantlyWarmupQuota() {
  try {
    const headers = {
      "Authorization": "Bearer " + instantlyApiKey,
      "Content-Type": "application/json"
    };

    const data = emailSheet.getDataRange().getValues();

    for (let i = 2; i < data.length; i++) {
      const email = data[i][0];
      const newQuota = data[i][6]; // Column G (index 6)

      if (!email || !newQuota) continue;

      const url = `https://api.instantly.ai/api/v2/accounts/${email}`;

      const payload = JSON.stringify({
        "warmup": {
          "limit": parseInt(newQuota)
        }
      });

      const options = {
        method: "patch",
        headers,
        payload,
        muteHttpExceptions: true
      };

      try {
        const res = UrlFetchApp.fetch(url, options);
        Logger.log(`Updated warmup quota for ${email}: ${JSON.stringify(res.getContentText(), null, '\t')}`);
      } catch (e) {
        Logger.log(`Failed to update warmup quota for ${email}: ${e.message}`);
      }
    }
  } catch (e) {
    Logger.log('Error in updateInstantlyWarmupQuota(): ' + e.toString());

    // Remove existing retry triggers
    ScriptApp.getProjectTriggers().forEach(function(trig) {
      if (trig.getHandlerFunction() === 'updateInstantlyWarmupQuota') {
        ScriptApp.deleteTrigger(trig);
      }
    });

    // Schedule a retry in 1 minute
    ScriptApp.newTrigger('updateInstantlyWarmupQuota')
      .timeBased()
      .after(1 * 60 * 1000)
      .create();
  }
}
  
function processRow(row, rowIndex) {
  const fromEmail = row[0];                   // Column A
  const customerName = row[2];                // Column C
  const status = row[3].toLowerCase();        // Column D
  const dailyQuota = row[4];                  // Column E
  
  if (!fromEmail) return; // Skip empty rows
  
  const url = 'https://app.companyname.com/api/company/processgooglesheetrow';
  
  const payload = {
    'from_email': fromEmail
  };
  
  // Only add parameters if they have values
  if (customerName) payload.customer_name = customerName;
  if (status) payload.status = status;
  if (dailyQuota) payload.daily_quota = dailyQuota;
  
  const options = {
    'method': 'post',
    'contentType': 'application/x-www-form-urlencoded',
    'payload': payload,
    'muteHttpExceptions': true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());
    
    // // Update status column (F) with result
    // SpreadsheetApp.getActiveSheet().getRange(rowIndex+1, 6).setValue(
    //   response.getResponseCode() === 200 ? 
    //   (result.status === 'success' ? 'Updated ✓' : 'Error: ' + result.message) :
    //   'API Error: ' + response.getResponseCode()
    // );
    
    // // Add details in column G
    // if (result.status === 'success' && result.updates) {
    //   SpreadsheetApp.getActiveSheet().getRange(rowIndex+1, 7).setValue(
    //     result.updates.join(', ')
    //   );
    // }
    Logger.log("Processing row: " + (rowIndex+1));
    Logger.log(result);
    
    return result;
  } catch (e) {
    SpreadsheetApp.getActiveSheet().getRange(rowIndex+1, 6).setValue('Error: ' + e.toString());
    return null;
  }
}

function updateCoSendingQuota() {
  try {
    const dataRange = emailSheet.getDataRange();
    const values = dataRange.getValues();
    
    // Process each row starting from row 3 (skip header)
    for (let i = 2; i < values.length; i++) {
      if (values[i][0] && (values[i][2] || values[i][3] || values[i][4])) {
        processRow(values[i], i);
        // Add a small delay to avoid hitting rate limits
        Utilities.sleep(500);
      }
    }
  } catch (e) {
    Logger.log('Error in updateCoSendingQuota(): ' + e.toString());

    // Remove existing retry triggers
    ScriptApp.getProjectTriggers().forEach(function(trig) {
      if (trig.getHandlerFunction() === 'updateCoSendingQuota') {
        ScriptApp.deleteTrigger(trig);
      }
    });

    // Schedule a retry in 1 minute
    ScriptApp.newTrigger('updateCoSendingQuota')
      .timeBased()
      .after(1 * 60 * 1000)
      .create();
  }
}

function getMailwizzDeliveryServers() {
  const mwzApiUrl = "https://app.companyname.com/api/company/getdeliveryservers";

  const options = {
    method: "get",
    headers: {
      "X-MW-PUBLIC-KEY": mailwizzApiKey,
      "X-MW-PRIVATE-KEY": mailwizzApiKey,
      "Accept": "application/json"
    },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(mwzApiUrl, options);
  console.log(response.getContentText());

  // const json = JSON.parse(response.getContentText());
  // console.log(json);
}

function getInstantlyAccountDetailsByEmail(email = "example@mail.com") {
  const apiUrl = "https://api.instantly.ai/api/v2/accounts?limit=100";
  const headers = {
    "Authorization": "Bearer " + instantlyApiKey
  };
  let nextStartingAfter = null;

  do {
    let url = apiUrl;
    if (nextStartingAfter) {
      url += "&starting_after=" + nextStartingAfter;
    }

    const response = UrlFetchApp.fetch(url, {
      method: "get",
      headers,
      muteHttpExceptions: true
    });

    const json = JSON.parse(response.getContentText());
    const items = json.items || [];
    for (let item of items) {
      if (item.email === email) {
        console.log("Account found: " + JSON.stringify(item, null, 2));
        return item; // return full account object
      }
    }

    nextStartingAfter = json.next_starting_after || null;
  } while (nextStartingAfter);

  Logger.log("No account found for email: " + email);
  return null;
}
