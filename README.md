# AdminDay2020
Stream of thing you need:


Macro Edit Snips:
var d = new Date();
var months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

.setValue(months[d.getMonth()] + " Blocked Issue Safety Report");

selectedSheet.getTables()[0];



Get Issues Macro:
function main(workbook: ExcelScript.Workbook): Employee[] {
    // Here we go!!
    var d = new Date();
    var months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
    let sheet = workbook.getWorksheet(months[d.getMonth()]);
    let table = sheet.getTables()[0];
    const PLANT_INDEX = 0;
    const NAME_INDEX = 1;
    const EMAIL_INDEX = 2;
    const ISSUE_INDEX = 3;
    const STATUS_REPORT_INDEX = 4;

    let bodyRangeValues = table.getRangeBetweenHeaderAndTotal().getValues();

    let people: Employee[] = [];
    for (let i = 0; i < bodyRangeValues.length; i++) {
        let row = bodyRangeValues[i];
        if (row[STATUS_REPORT_INDEX] === "Blocked" ){
            // save email to return
            people.push({plant: row[PLANT_INDEX], name: row[NAME_INDEX], email: row[EMAIL_INDEX], issue: row[ISSUE_INDEX]});
        }
    }
//log the array just to check
console.log(people);
return people;
}

/**
 * An interface to represent the employee
 * array returned from Script
 * 
 */
interface Employee {
    plant: string,
    name: string,
    email: string,
    issue: string
}


Update Status Macro:
function main(workbook: ExcelScript.Workbook,
senderEmail: string,
plant: string,
statusReportResponse: string) 
{
    // code here
  var d = new Date();
  var months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  let sheet = workbook.getWorksheet(months[d.getMonth()]);
  let table = sheet.getTables()[0];
  const PLANT_INDEX = 0;
  const NAME_INDEX = 1;
  const EMAIL_INDEX = 2;
  const ISSUE_INDEX = 3;
  const STATUS_REPORT_INDEX = 4;

  let bodyRange = table.getRangeBetweenHeaderAndTotal();
  let tableRowCount = bodyRange.getRowCount();
  let bodyRangeValues = bodyRange.getValues();
  
  // tracking flag
  let statusAdded = false;
let rowID = 0;

  //loop time
  for (let i = 0; i<tableRowCount && !statusAdded; i++){
    let row = bodyRangeValues[i];
    //check for match 
    if (row[EMAIL_INDEX] === senderEmail && row[PLANT_INDEX] == plant){
      // add card response
      bodyRange.getCell(i, STATUS_REPORT_INDEX).setValues([
        [statusReportResponse]
      ]);
        statusAdded = true;
        rowID = i;
    }
  }
  // log the statusAdded
  if (statusAdded){
    console.log(
      "Succesfully updated status for " + senderEmail +"containing:" + statusReportResponse + " on row " + rowID
    );
  }

}



Output Sample from Adaptive Card:
outputs('CardWaitAction').body.data.response


Adaptive Card JSON
{
  "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
  "type": "AdaptiveCard",
  "version": "1.0",
  "body": [
    {
      "type": "Container",
      "items": [
        {
          "type": "TextBlock",
          "text": "Safety Update Needed: PLANTNAME",
          "weight": "bolder",
          "size": "medium"
        },
        {
          "type": "ColumnSet",
          "columns": [
            
            {
              "type": "Column",
              "width": "stretch",
              "items": [
                {
                  "type": "TextBlock",
                  "text": "Sent From: Reminder Service",
                  "weight": "bolder",
                  "wrap": true
                }
              ]
            }
          ]
        }
      ]
    },
    {
      "type": "Container",
      "items": [
        {
          "type": "ColumnSet",
          "columns": [
            
            {
              "type": "Column",
              "width": "auto",
              "items": [
                {
                  "type": "TextBlock",
                  "text": "Issue: SAFETYISSUE",
                  "wrap": true
                }
              ]
            }
          ]
        }
      ]
        
    },
    {
      "type": "Container",
      "items": [
        {
          "type": "FactSet",
          "facts": [
            
            {
              "title": "Note:",
              "value": "This reminder was sent to you via the automated service for action prior to month-end reprting.  Replies to this thread in Teams are not responded to by the bot.  "
            }
          ]
        }
      ]
    },
    {
      "type": "Input.ChoiceSet",
      "choices":[
        {
          "title":"Blocked",
          "value":"Blocked"
        },
        {
          "title":"Open",
          "value":"Open"
        }
        ,
        {
          "title":"Closed",
          "value":"Closed"
        }
      ],
      "placeholder": "Pick an issue status",
      "id": "response"
    }
  ]
  ,
  "actions": [
    {
      "type":"Action.Submit",
      "title": "Submit",
      "id": "submit"
    }
  ]
}


