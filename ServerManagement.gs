function quotaCalculations() {
  calculateInstantlyQuota(true);
  calculateCoQuota(true);
  postEmailManagementData();
}

function quotaCalculationsManual() {
  calculateInstantlyQuota(false);
  calculateCoQuota(false);
  postEmailManagementData();
}

function calculateInstantlyQuotaTrue() { calculateInstantlyQuota(true); }
function calculateInstantlyQuotaFalse() { calculateInstantlyQuota(false); }

function calculateCoQuotaTrue() { calculateCoQuota(true);  }
function calculateCoQuotaFalse() { calculateCoQuota(false); }

function calculateInstantlyQuota(updateDays = false) {
  try {
    const dataRange = serverSheet.getDataRange();
    const values = dataRange.getValues();

    // Process each row starting from row 3 (skip header)
    for (let i = 2; i < values.length; i++) {
      if (values[i][4] && values[i][6] && values[i][6] > 0 && values[i][17]) {
        let additionInstantly = (values[i][4] - values[i][3]) / values[i][6];
        instantlyCalculation(i + 1, additionInstantly, updateDays);

        let emailData = emailSheet.getRange("A3:B").getValues();
        let eligibleEmails = emailData.filter(val => val[0].includes(values[i][17]));
        let prioritizedEmails = eligibleEmails.filter(val => val[1] == "Bulk Sending (CO)");
        let otherEmails = eligibleEmails.filter(val => val[1] != "Bulk Sending (CO)");
        var remappedEmailData = prioritizedEmails.concat(otherEmails);
        var newEmails = [];

        if (remappedEmailData.length < additionInstantly / 200) {
          var emailsNeeded = Math.ceil(additionInstantly / 200) - remappedEmailData.length;
          do {
            let request = UrlFetchApp.fetch(`https://randomuser.me/api/?results=${emailsNeeded + 2}&nat=gb`);
            let response = JSON.parse(request.getContentText());
            let formattedEmails = response.results.map(val => `${val.name.first.toLowerCase()}.${val.name.last.toLowerCase()}@${values[i][17]}`);

            for (email of formattedEmails) {
              if (!remappedEmailData.map(val => val[0]).includes(email) && emailsNeeded > 0) {
                newEmails.push([email, "Instantly Warmup Only"]);
                remappedEmailData.push([email, "Instantly Warmup Only"]);
                emailsNeeded -= 1;
              }
            }
          } while (emailsNeeded > 0);
        }

        let emailsLength = remappedEmailData.length;
        let distribution = new Array(emailsLength).fill(0);
        let remaining = additionInstantly;

        let bulkIndices = [];
        let otherIndices = [];
        for (let i = 0; i < emailsLength; i++) {
          if (remappedEmailData[i][1] === "Bulk Sending (CO)") {
            bulkIndices.push(i);
          } else {
            otherIndices.push(i);
          }
        }

        if (additionInstantly < emailsLength) {
          for (let i of bulkIndices) {
            if (remaining <= 0) break;
            distribution[i] = 1;
            remaining--;
          }
          for (let i of otherIndices) {
            if (remaining <= 0) break;
            distribution[i] = 1;
            remaining--;
          }
        } else {
          for (let i of bulkIndices) {
            if (remaining <= 0) break;
            let give = Math.min(200, remaining);
            distribution[i] = give;
            remaining -= give;
          }
          for (let i of otherIndices) {
            if (remaining <= 0) break;
            let give = Math.min(200, remaining);
            distribution[i] = give;
            remaining -= give;
          }
        }

        if (newEmails.length > 0) {
          emailSheet.getRange(emailData.filter(val => val[0]).length + 3, 1, newEmails.length, 2).setValues(newEmails);
          SpreadsheetApp.flush();

          try {
            let endpoint = `https://api.instantly.ai/api/v2/accounts`;

            // extract just the address strings
            let emails = newEmails.map(row => row[0]);

            for (email of emails) {
              let options = {
                method: 'post',
                contentType: 'application/json',
                headers: {
                  Authorization: `Bearer ${instantlyApiKey}`
                },
                payload: JSON.stringify({
                  "email": email,
                  "first_name": email.split(".")[0],
                  "last_name": email.split(".")[1].split("@")[0],
                  "provider_code": 2,
                  "imap_username": "imapuser",
                  "imap_password": "imappass",
                  "imap_host": "111.111.111.111",
                  "imap_port": 1111,
                  "smtp_username": "user@user.com",
                  "smtp_password": "paswword",
                  "smtp_host": "222.222.222.222",
                  "smtp_port": 2222
                }),
                muteHttpExceptions: true
              };

              let resp = UrlFetchApp.fetch(endpoint, options);
              let code = resp.getResponseCode();
              let body = JSON.parse(resp.getContentText());
              if (code < 200 || code >= 300) {
                console.error('Instantly API error:', code, body);
              } else {
                console.log(`Added ${emails.length} new emails to Instantly list`);
              }
            }
          } catch(e) {
            console.log(e);
          }
        }

        emailData = emailSheet.getRange("A3:B").getValues();
        for (k in remappedEmailData) {
          let idx = emailData.findIndex(row => row[0] === remappedEmailData[k][0]);

          if (idx === -1) {
            console.warn(`Couldn’t find ${remappedEmailData[k][0]} in your sheet!`);
            continue;
          }

          let row = idx + 3;  
          emailSheet.getRange(row, 7).setValue(distribution[k]);
        }
      }
    }
  } catch (e) {
    Logger.log(`Error in calculateInstantlyQuota(${updateDays}): ${e}`);

    // remove any existing retry trigger for this mode
    const handler = updateDays
      ? 'calculateInstantlyQuotaTrue'
      : 'calculateInstantlyQuotaFalse';
    ScriptApp.getProjectTriggers().forEach(trig => {
      if (trig.getHandlerFunction() === handler) {
        ScriptApp.deleteTrigger(trig);
      }
    });

    // schedule a one-time retry in 1 minute
    ScriptApp.newTrigger(handler)
      .timeBased()
      .after(1 * 60 * 1000)
      .create();
  }
}

function calculateCoQuota(updateDays = true) {
  try {
    const dataRange = serverSheet.getDataRange();
    const values = dataRange.getValues();

    // Process each row starting from row 3 (skip header)
    for (let i = 2; i < values.length; i++) {
      if (values[i][11] && values[i][13] && values[i][13] > 0 && values[i][17]) {
        let additionCo = (values[i][11] - values[i][10]) / values[i][13];
        coCalculation(i + 1, additionCo, updateDays);
        
        let emailData = emailSheet.getRange("A3:D").getValues();
        let remappedEmailData = emailData.filter(val => val[0].includes(values[i][17]) && val[1] == "Bulk Sending (CO)");
        let newEmails = [];

        if (remappedEmailData.length < 1) {
          newEmails = relationshipMailSheet.getRange("A2:F").getValues().filter(val => val[4].includes(values[i][17])).map(val => [val[4], "Bulk Sending (CO)", val[0], "Active"]);
          remappedEmailData = relationshipMailSheet.getRange("A2:F").getValues().filter(val => val[4].includes(values[i][17])).map(val => [val[4], "Bulk Sending (CO)", val[0], "Active"]);
        }

        let emailsLength = remappedEmailData.length;
        let distribution = new Array(emailsLength).fill(0);
        let remaining = additionCo;
        let base  = Math.floor(remaining / emailsLength);
        let extra = remaining % emailsLength;

        remappedEmailData.forEach((_, idx) => {
          distribution[idx] = base + (idx < extra ? 1 : 0);
        });

        if (newEmails.length > 0) {
          emailSheet.getRange(emailData.filter(val => val[0]).length + 3, 1, newEmails.length, 4).setValues(newEmails);
          SpreadsheetApp.flush();
        }

        emailData = emailSheet.getRange("A3:A").getValues();
        for (k in remappedEmailData) {
          let idx = emailData.findIndex(row => row[0] === remappedEmailData[k][0]);

          if (idx === -1) {
            console.warn(`Couldn’t find ${remappedEmailData[k][0]} in your sheet!`);
            continue;
          }

          let row = idx + 3;  
          emailSheet.getRange(row, 5).setValue(distribution[k]);
        }
      }
    }
  } catch (e) {
    Logger.log(`Error in calculateCoQuota(${updateDays}): ${e}`);

    // clean up old retry trigger
    const handler = updateDays
      ? 'calculateCoQuotaTrue'
      : 'calculateCoQuotaFalse';
    ScriptApp.getProjectTriggers().forEach(trig => {
      if (trig.getHandlerFunction() === handler) {
        ScriptApp.deleteTrigger(trig);
      }
    });

    // schedule the retry
    ScriptApp.newTrigger(handler)
      .timeBased()
      .after(1 * 60 * 1000)
      .create();
  }
}

function instantlyCalculation(row, addition, updateDays) {
  serverSheet.getRange(row, 4).setValue((serverSheet.getRange(row, 4).getValue() ?? 0) + addition);
  serverSheet.getRange(row, 6).setValue(addition);
  
  if (updateDays) {
    serverSheet.getRange(row, 7).setValue(serverSheet.getRange(row, 7).getValue() - 1);
  }
}

function coCalculation(row, addition, updateDays) {
  serverSheet.getRange(row, 11).setValue((serverSheet.getRange(row, 11).getValue() ?? 0) + addition);
  serverSheet.getRange(row, 13).setValue(addition);

  if (updateDays) {
    serverSheet.getRange(row, 14).setValue(serverSheet.getRange(row, 14).getValue() - 1);
  }
}
