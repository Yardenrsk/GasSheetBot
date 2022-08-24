//Yarden Rotem - 05/07/2022

//setting the connection between Google Sheets and Telegram API

let token = "************";

function getME() {
  let response =   UrlFetchApp.fetch("https://api.telegram.org/bot"+ token+ "/getMe");
  console.log(response.getContentText());
}

function setWebhook() {
  let webAppUrl = "************";

  let response =   UrlFetchApp.fetch("https://api.telegram.org/bot"+ token+ "/setWebhook?url=" + webAppUrl);
  console.log(response.getContentText());
}


//http request for sending message on telegram
function sendText(chat_id, text){
  let data = {
    method: "post",
    payload: {
      method: "sendMessage",
      chat_id: String(chat_id),
      text: text,
      parse_mode: "HTML"
    }
  };
   UrlFetchApp.fetch("https://api.telegram.org/bot"+ token+ "/" , data);
}

//the whole proccess of sending a chart (need to be formated as image and then it can be sent)
function sendChart(chat_id)
{
  let sheet = SpreadsheetApp.getActive().getSheetByName("GRAPH");
  let chart = sheet.getCharts()[0];
  let chartImage = formatChartAsImage(chart);
  const file = DriveApp.createFile(chartImage);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const outputUrl = "https://drive.google.com/uc?export=download&id=" + file.getId();
  console.log(outputUrl);
  sendImage(chat_id,outputUrl);

}

//http request for sending image to telegram chat
function sendImage(chat_id, text) {

   UrlFetchApp.fetch("https://api.telegram.org/bot" + token + "/sendPhoto?caption=" + encodeURIComponent("") + "&photo=" + encodeURIComponent(text) + "&chat_id=" 
   + chat_id + "&parse_mode=HTML");
}

//charts from a Sheet are formatted unreliably. Better to channel them through a Slide presentation.
function formatChartAsImage(chart) {

    let proxySaveSlide = SlidesApp.openById("1OndvGO-2kKfnyxyXip_uB8NAwdfAsOp8qRwkhM4bdHY").getSlides()[0];
    let chartImage = proxySaveSlide.insertSheetsChartAsImage(chart);

    // Get image from slides
    let myimage = chartImage.getBlob();

    // Delete image in presentation slide
    chartImage.remove();

    // Return the image
    return myimage;
}


//catching incoming message and adding the values to the sheet in the right places
function doPost(e){
  let contents = JSON.parse(e.postData.contents);
  let chat_id = contents.message.chat.id;
  let text = contents.message.text;
  if(!isNumeric(text))
  {
    sendText(chat_id,"Please enter numbers only.\nyou can use floating point.\nEnter 0 to only get info.")
  }
  else if(text == "0")
  {
    let row = findRow(chat_id);
    let lastMonthVal = SpreadsheetApp.getActive().getSheetByName("DB").getRange(row-1,26).getValue();
    let monthTotal = SpreadsheetApp.getActive().getSheetByName("DB").getRange(row,24).getValue();
    let lastMonthTotal = SpreadsheetApp.getActive().getSheetByName("DB").getRange(row-1,24).getValue();
    let leftTanks = Number.parseFloat((500-monthTotal)/50).toFixed(2);
    let msg = "info:\nThis month (" + getMonthAndYear() +") total:\n"+ monthTotal  +
                " Liters.\nTanks left for this month:\n" + leftTanks +" tanks.\nLast month (" + lastMonthVal + ") total:\n" + lastMonthTotal + " Liters.\nHave a good day :)";

    sendText(chat_id,msg);
    sendChart(chat_id);

  }
  else
  {
  let row = findRow(chat_id);
  let col = findCol(row,chat_id);
  let lastMonthVal = SpreadsheetApp.getActive().getSheetByName("DB").getRange(row-1,26).getValue();
  SpreadsheetApp.getActive().getSheetByName("DB").getRange(row,col).setValue(text);
  SpreadsheetApp.getActive().getSheetByName("DB").getRange(row,col+1).setValue(getToday());
  let monthTotal = SpreadsheetApp.getActive().getSheetByName("DB").getRange(row,24).getValue();
  let lastMonthTotal = SpreadsheetApp.getActive().getSheetByName("DB").getRange(row-1,24).getValue();
  let leftTanks = Number.parseFloat((500-monthTotal)/50).toFixed(2);
  let msg = "Value was added to the sheet:\n" + text + " Liters.\nThis month (" + getMonthAndYear() +") total:\n"+ monthTotal  +
             " Liters.\nTanks left for this month: " + leftTanks +" tanks.\nLast month (" + lastMonthVal + ") total:\n" + lastMonthTotal + " Liters.\nHave a good day :)";

  sendText(chat_id,msg);
  sendChart(chat_id);
  }
}

//return today's date in "dd/mm/yyyy" format
function getToday()
{
  let today = new Date();
  var dd = String(today.getDate()).padStart(2, '0');
  var mm = String(today.getMonth() + 1).padStart(2, '0'); //January is 0!
  var yyyy = today.getFullYear();

  today = dd + '/' + mm + '/' + yyyy;
  console.log(today);
  return today;
}

//returns today's date in "mm/yyyy" format
function getMonthAndYear()
{
  let today = new Date();
  var dd = String(today.getDate()).padStart(2, '0');
  var mm = String(today.getMonth() + 1).padStart(2, '0'); //January is 0!
  var yyyy = today.getFullYear();

  res = mm + '/' + yyyy;
  console.log(res);
  return res;
}

//finding the suitable row to the current month
function findRow(chat_id)
{
  let res = 2;
  for (res = 2; res <= 32 ; res++) {
       if (SpreadsheetApp.getActive().getSheetByName("DB").getRange(res,26).getValue() == getMonthAndYear()){
      return res;
       }
  }
  sendText(chat_id,"Could not found month: " + getMonthAndYear() + ", speak with Yarden.");
}

//finding the 1st empty coloum to add in the gas value
function findCol(row,chat_id)
{
  let res = 2;
  for (res = 2; res <= 22 ; res++) {
       if (SpreadsheetApp.getActive().getSheetByName("DB").getRange(row,res).isBlank()){
      return res;
       }
  }
  sendText(chat_id,"more than expected times than expected, speak with Yarden.");
}

//checking if string is numeric or not
function isNumeric(str) {
  if (typeof str != "string") return false // we process only strings  
  return !isNaN(str) && 
         !isNaN(parseFloat(str)) 
}


//the whole proccess of sending a chart (need to be formated as image and then it can be sent)
function sendChart(chat_id)
{
  let sheet = SpreadsheetApp.getActive().getSheetByName("GRAPH");
  let chart = sheet.getCharts()[0];
  let chartImage = formatChartAsImage(chart);
  const file = DriveApp.createFile(chartImage);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const outputUrl = "https://drive.google.com/uc?export=download&id=" + file.getId();
  console.log(outputUrl);
  sendImage(chat_id,outputUrl);

}

//http request for sending image to telegram chat
function sendImage(chat_id, text) {

   UrlFetchApp.fetch("https://api.telegram.org/bot" + token + "/sendPhoto?caption=" + encodeURIComponent("") + "&photo=" + encodeURIComponent(text) + "&chat_id=" 
   + chat_id + "&parse_mode=HTML");
}

//charts from a Sheet are formatted unreliably. Better to channel them through a Slide presentation.
function formatChartAsImage(chart) {

    let proxySaveSlide = SlidesApp.openById("1OndvGO-2kKfnyxyXip_uB8NAwdfAsOp8qRwkhM4bdHY").getSlides()[0];
    let chartImage = proxySaveSlide.insertSheetsChartAsImage(chart);

    // Get image from slides
    let myimage = chartImage.getBlob();

    // Delete image in presentation slide
    chartImage.remove();

    // Return the image
    return myimage;
}





