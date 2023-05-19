var ss = SpreadsheetApp.getActiveSpreadsheet();

// time related
var date = new Date();
var theDate = date.toLocaleDateString();
var time = date.toLocaleTimeString();
var mySchedule;

function makeAgenda() {
  try {
    // sets date and time on sheet
    ss.getSheetByName("Agenda").getRange('A1').setValue(theDate);    
    ss.getSheetByName("Agenda").getRange('D3').setValue("[" + time.substring(0,time.length-6) + " " + time.substring(time.length-2) + "]");

    // check buttons and sets schedule
    checkButtons();

    // changes class period of the day
    changePeriod(mySchedule);

    // true when time = 11:59PM
    var nextDay = time.substring(0,5) == "11:59" && time.substring(time.length-2) == "PM";
    // true when it is not Sunday or Saturday
    var notWeekend = date.getDay() != "6" && date.getDay() != "0";

    // resets agenda every day
    if (nextDay && notWeekend){

      // makes a copy of Agenda sheet
      ss.getSheetByName('Agenda').copyTo(ss);
      var duplicateSheet = ss.getSheetByName("Copy of Agenda")
      duplicateSheet.setName("Agenda " + date.toLocaleDateString());

      clearAgenda();
      midtermCount();
    }
  } catch (err){
    console.log(err);
  }
}

function checkButtons(){
  var forum = ss.getSheetByName("Agenda").getRange("G2").getValue();
  var halfDay = ss.getSheetByName("Agenda").getRange("H2").getValue();
  var twoHour = ss.getSheetByName("Agenda").getRange("I2").getValue();
  var skipLetterDay = ss.getSheetByName("Agenda").getRange("J2").getValue();

  if (forum){
    getSchedule("forum");
  } else if (halfDay){
    getSchedule("halfday");
  } else if (twoHour){
    getSchedule("2hr");
  } else if (skipLetterDay){    
        ss.getSheetByName("Agenda").getRange("J2").setValue(false);
    
    for (var i = 0; i < 5; i++){
        var letterDay = ss.getSheetByName("Agenda").getRange('D1').getValue();
        changeLetterDay(letterDay);
    }
  } else {
    getSchedule("normal");
  }
}

function clearAgenda(){
  ss.getSheetByName("Agenda").getRange("G2").setValue(false);
  ss.getSheetByName("Agenda").getRange("H2").setValue(false);
  ss.getSheetByName("Agenda").getRange("I2").setValue(false);

  var startingRow = 2; // row where agenda starts
  var endingRow = 7; // row where agenda ends
  for (var i = startingRow; i <= endingRow; i++){
    ss.getSheetByName("Agenda").getRange(i,2).setValue("");
  }
  var letterDay = ss.getSheetByName("Agenda").getRange('D1').getValue();
  changeLetterDay(letterDay);
}
function changeLetterDay(letterDay){
  if (letterDay == "Day A"){
    ss.getSheetByName("Agenda").getRange('D1').setValue("Day B");
  } else if (letterDay == "Day B"){
    ss.getSheetByName("Agenda").getRange('D1').setValue("Day C");
  } else if (letterDay == "Day C"){
    ss.getSheetByName("Agenda").getRange('D1').setValue("Day D");   
  } else if (letterDay == "Day D"){
    ss.getSheetByName("Agenda").getRange('D1').setValue("Day E");
  } else if (letterDay == "Day E"){
    ss.getSheetByName("Agenda").getRange('D1').setValue("Day F");
  } else if (letterDay == "Day F"){
    ss.getSheetByName("Agenda").getRange('D1').setValue("Day A");
  }
}

function getSchedule(typeOfSchedule){
  var letterOne = "A";
  var letterTwo = "B";
  var letterThree ="C";

  if (typeOfSchedule =="forum"){
    letterOne = "E";
    letterTwo = "F";
    letterThree = "G";
  } else if (typeOfSchedule =="halfday") {
    letterOne = "I";
    letterTwo = "J";
    letterThree = "K";
  } else if (typeOfSchedule =="2hr"){
    letterOne = "M";
    letterTwo = "N";
    letterThree = "O";
  } 

  mySchedule =[[],[],[],[],[],[],[],[],[],[],[],[]];
  for (var i = 3; i < 15; i++){
      var tempPeriod = ss.getSheetByName("Schedule").getRange(letterOne + i).getValue();
      var tempTime = ss.getSheetByName("Schedule").getRange(letterTwo + i).getValue().toLocaleString().substring(12);
      var tempClass = ss.getSheetByName("Schedule").getRange(letterThree + i).getValue();
      mySchedule[i-3].push(tempPeriod, tempTime, tempClass);
  }

  return mySchedule;
}

function changePeriod(mySchedule){
  //console.log(time.substring(0,time.length-6));

  for (var i = 0; i < mySchedule.length; i++){
    // console.log(time)
    // console.log(mySchedule[i][1])
    if (time.substring(0,time.length-6) == mySchedule[i][1].substring(0, mySchedule[i][1].length-6)){
      ss.getSheetByName("Agenda").getRange('D2').setValue("Period: " + mySchedule[i][0]);
    }
  }
}

function midtermCount(){
  var count = parseInt(ss.getSheetByName("Agenda").getRange('D5').getValue());
  if (count > 0){
    if (count <=5){
      ss.getSheetByName("Agenda").getRange('D5').setFontColor("red");    
    }
    ss.getSheetByName("Agenda").getRange('D5').setValue(count-1);
  }
}
