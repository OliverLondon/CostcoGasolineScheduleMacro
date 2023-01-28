//config for features
const giveWarning = false     //set to true or false if you want/don't want the red boxes saying "check this row" to appear on partial shifts @ gas
const makeCoverageView = true //creates a second spreadsheet showing how many people are out at the station at any given time. Uses times filled out by this macro.
const changeNames = true      //T/F if you want names in the "namesToBeFixed" to be updated with what name people actually call them by. 
const dontDeleteRows = true   //when attempting to fit to page, tall schedules might have up to 2 lines from the Hours Cut/added section removed. Set to True to ensure all 6 lines remain.
const sortingFormat = "hybrid" //Options: "start" "end" and "hybrid", sorts people by their start or end times. Hybrid does the top half with start, bottom with end.
const hybridBias = 1          //A constant that determines the portion of the sheet that is sorted by end times in hybrid sorting format. 
                              //0 = split down middle. Bigger is more of the sheet sorted by end time. negative (ex: -2) is more of the sheet sorted by start time
                              //Only input whole numbers for hybridBias, as it ignores all decimals.
const tryToFixSentFromFE = true  //true OR false accepted. Attempts to fix the 15 or 30 minute discrepency on the schedule when employees are sent from the front end.  
                                  //i.e. Lunch is at 5:00 pm, but our board has them start at gas at 5:00 pm as well. Should be 5:30 pm on our board.
const tellChanges = true      //true OR false accepted. Announces when the above setting changes the start time of a shift

//To add a new name: 
//In the first list put a comma after the current bottom name, then on a new line add the new name in quotes.
//In the second list put a comma after the current bottom name, the on a new line add the new name that should replace the corrisponding entry in the first list, in quotes.
const namesToBeFixedA = [
  "WICKIZER, JONATHAN",
  "COLE, JOSHUA",
  "PETTIT, TIMOTHY",
  "ALLEN, TERRANCE",
  "DYER, NICHOLAS",
  "ADAMS, RONALD",
  "TICHENOR, MATTHEW",
  "LAKEBERG, ZACHARY",
  "LOGAN, CINDY",
  "GIGLIOTTI, JOSEPHINE",
  "STORM-CARROLL, MICHELLE"
]
const namesToBeFixedB = [
  "WICKIZER, JAY",
  "COLE, JOSH",
  "PETTIT, TIM",
  "ALLEN, TERRY",
  "DYER, NICK",
  "ADAMS, RON",
  "TICHENOR, MATT",
  "LAKEBERG, ZACK",
  "LOGAN, LOU",
  "GIGLIOTTI, JOJO",
  "STORM-CARROLL, MICAH"
]

//global variables ++DO NOT EDIT++
const hourmillis = 1000 * 60 * 60
const halfhourmillis = 1000 * 60 * 30
const fifteenminmillis = 1000 * 60 * 15
const maxpixelheight = 745
//end globals

function TimeTableFill() {
  //Get the sheet being operated on
  const base = SpreadsheetApp.getActiveSpreadsheet()
  const s1 = base.getSheetByName("Sheet1")
  
  //break every cell at the top apart for simple formatting down the line
  s1.getRange("A1:R16").breakApart().clearFormat()
  s1.getRange("F19:H"+s1.getLastRow()).clearContent()

  //get all the data inside
  var rawdata = s1.getDataRange().getValues()

  var numsups = 0
  var toprow
  var bottomrow
  //find the first row with a person's hours (0 index)
  rawdata.forEach(function(value,index){
    if(value[2] =="Name"){ //Find when names start
      toprow = index + 1
    }
    bottomrow = index  //keep reassigning the bottom until we run out of rows with people
  })

  //count number of sups today (just top 4 rows)
  for (var j = toprow; j < toprow + 4; j = j+1){
    if (s1.getRange(j,2).getValue()== "Sup Hourly"){numsups = numsups+1}
  }
  //sorts people by selected format, followed by  filling in the times in the middle
  SortPeople(s1,toprow,bottomrow, numsups, sortingFormat)
  var timeRange = s1.getRange("D"+(toprow+1)+":I"+s1.getLastRow())
  var valRange = timeRange.getValues()
  FillBreakAid(s1,timeRange,valRange)

  ///Sheet cleanup:
  s1.deleteRows(4,14) //starting with 4, delete 14 rows
  s1.getRangeList(['A3:M3','A4:B60','J4:J60','K4:K60']).clearContent()
  FormatTopRows(s1,numsups)
  
  //write in the remaining cells that need text
  bottomrow = s1.getLastRow()
  s1.getRange('A'+(bottomrow+1)+':K'+(bottomrow+1)).setValues([['Hours Cut:','','','','Hours Added:','','','','Misc.','','']]).setHorizontalAlignment("left")
  var extrarows = 17 - (bottomrow-5)//ensure enough space for all hourly security checks
  for(var i = 0; i < extrarows; i = i+1){
    s1.insertRowAfter(bottomrow)
    s1.getRange('A'+(bottomrow+1)+':L'+(bottomrow+1)).setBackground('white')
  }
  var newbottom = bottomrow + extrarows
  s1.getRange('A'+(bottomrow+1)+':K'+newbottom).setBorder(true,true,true,true,true,true,"black",SpreadsheetApp.BorderStyle.SOLID)

  Resizer(s1) //lots of formatting
  FitToPageAttempt(s1,bottomrow) //adds or removes lines based on how many people are working that day. Can toggle removing lines off at top.
  
  var datename = s1.getRange("C1").getValue().toString()
  var today = ""
  if (datename[0] == "S"){//check if it's weekend ("Sat" or "Sun" in the date's format)
    if (datename[2] == "t"){//Saturday
      today = "Saturday"
    }
    else{//Sunday
      today = "Sunday"
    }
  }
  SecurityChecks(s1,today) //fills the hourly security checks.
  if(changeNames){SetPreferredNames(s1)}
  if (makeCoverageView){CoverageView()}
}

//=====================================================================================================================================================================================
function CoverageView(){
  SpreadsheetApp.getActiveSpreadsheet().insertSheet('Coverage');
  const s2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Coverage")
  const s0 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1")
  
  s2.insertColumns(1,180)
  s2.setColumnWidths(1,206,37)
  s2.getRange("A1:GX1").setBackground("lightgray").setValues([["Name","5:00",	"5:05",	"5:10",	"5:15",	"5:20",	"5:25",	"5:30",	"5:35",	"5:40",	"5:45",	"5:50",	"5:55",	"6:00",	
  "6:05",	"6:10",	"6:15",	"6:20",	"6:25",	"6:30",	"6:35",	"6:40",	"6:45",	"6:50",	"6:55",	
  "7:00", "7:05",	"7:10",	"7:15",	"7:20",	"7:25",	"7:30",	"7:35",	"7:40",	"7:45",	"7:50",	"7:55",	
  "8:00",	"8:05",	"8:10",	"8:15",	"8:20",	"8:25",	"8:30",	"8:35",	"8:40",	"8:45",	"8:50",	"8:55",	
  "9:00",	"9:05",	"9:10",	"9:15",	"9:20",	"9:25", "9:30", "9:35",	"9:40",	"9:45",	"9:50",	"9:55",	
  "10:00", "10:05", "10:10", "10:15", "10:20", "10:25", "10:30", "10:35", "10:40", "10:45", "10:50", "10:55", 
  "11:00", "11:05", "11:10", "11:15", "11:20", "11:25", "11:30", "11:35", "11:40", "11:45", "11:50", "11:55", 
  "12:00", "12:05", "12:10", "12:15", "12:20", "12:25", "12:30", "12:35", "12:40", "12:45", "12:50", "12:55", 
  "1:00", "1:05", "1:10", "1:15", "1:20", "1:25", "1:30", "1:35", "1:40", "1:45", "1:50", "1:55", 
  "2:00", "2:05", "2:10", "2:15", "2:20", "2:25", "2:30", "2:35", "2:40", "2:45", "2:50", "2:55", 
  "3:00", "3:05", "3:10", "3:15", "3:20", "3:25", "3:30", "3:35", "3:40", "3:45", "3:50", "3:55", 
  "4:00", "4:05", "4:10", "4:15", "4:20", "4:25", "4:30", "4:35", "4:40", "4:45", "4:50", "4:55", 
  "5:00", "5:05", "5:10", "5:15", "5:20", "5:25", "5:30", "5:35", "5:40", "5:45", "5:50", "5:55", 
  "6:00", "6:05", "6:10", "6:15", "6:20", "6:25", "6:30", "6:35", "6:40", "6:45", "6:50", "6:55", 
  "7:00", "7:05", "7:10", "7:15", "7:20", "7:25", "7:30", "7:35", "7:40", "7:45", "7:50", "7:55", 
  "8:00", "8:05", "8:10", "8:15", "8:20", "8:25", "8:30", "8:35", "8:40", "8:45", "8:50", "8:55", 
  "9:00", "9:05", "9:10", "9:15", "9:20", "9:25", "9:30", "9:35", "9:40", "9:45", "9:50", "9:55","10:00"]]).setVerticalAlignment("middle").setHorizontalAlignment("center")
  s2.setColumnWidth(1,200)
  //grab the names of the first 40 people working, should be less people than that on any board, but just going overkill to cove extreme cases.
  s2.getRange("A2:A42").setValues(s0.getRange("B5:B45").getValues()).setVerticalAlignment("middle").setHorizontalAlignment("center")

  var lastrow = s2.getLastRow() - 1
  var breakAidTimes = s0.getRange("E5:I"+(s0.getLastRow()-1)).getDisplayValues()
  var coverageTimesAbove = s2.getRange("B1:GX1").getDisplayValues() // a 1 x 205 sized array 
  var coverageFillRange = s2.getRange(2,2,(s2.getLastRow()-1),205) // a (num_persons) x 205 sized Range
  var coverageBackgroundColors = coverageFillRange.getBackgrounds() // a (num_persons) x 205 sized array with values being the background colors (all white = #ffffff)

  for (var i = 0; i < lastrow; i = i+1){//for every person
    var currentCol = 0 //max col is 206 (205 0 based index)
    var caseNum = 0
    //the following variables will be set to a column number corrisponding to the time they start, end, have a lunch, etc.
    var st = 0
    var b1 = 0
    var lun = 0
    var b2 = 0
    var et = 0
    for (var j = 0; j < 5; j = j+1){//for each time (start, break1, lunch, break2, end)
      //Find the rough column that matches the time, skipping ahead if possible
      caseNum = 0
      var str1 = breakAidTimes[i][j]
      if ((str1 == null) || (str1 == "")){ //if no time, go to next loop
        continue
      }
      str1 = str1.toString()
      if (str1[0] == "x"){//if the time slot is "xxx" or similar, skip. Shouldn't run as they are cleared out in formatting, but left in in case.
        continue
      }
      if(str1[5] == "P"){//Start at 1:00 PM
        currentCol = 98
        caseNum = 1
      }
      else{
        if(str1[6] == "P"){ //Either 12:00 PM or 10:00 PM and after, start at 12PM
          currentCol = 86
          caseNum = 2
        }
        else{
          if(str1[5] == "A"){//5AM to 9AM, start at 5AM
            currentCol = 2
            caseNum = 3
          }
          else{
            if(str1[6] == "A"){//10AM to 11AM, start at 10AM
              currentCol = 62
              caseNum = 4
            }
          }
        }
      }
      while(true){//find the start column that matches the times for their shift start, any breaks and lunches, and end time.
        var str2 = coverageTimesAbove[0][currentCol-2].toString()
        if (str1[0] == str2[0]){//if first number matches, check if following characters match, else skip to next hour
          if (str1[1] == str2[1]){//if second char matches continue, else in caseNum 4: skip 1 hour.
            if(str1[2] == str2[2]){// ":" if 2 digit hour, otherwise compares 10 min increment. If not a match, skip 10 mins
              if(str1[3] == str2[3]){//10 min increment if 2 digit hour, 5 min increment if 1 digit hour. If match, continue, else based on case skip 5 or 10 mins
                if((str1[4] == " ") || (str1[4] == str2[4])){//5 min increment if 2 digit hour, a space if 1 digit hour, if latter, match found!
                  break //starting column found
                }
                else{currentCol = currentCol + 1}                
              }
              else{
                if ((caseNum == 2) || (caseNum == 4)){
                  currentCol = currentCol + 2
                }
                else{currentCol = currentCol + 1}
              }
            }
            else{currentCol = currentCol + 2}
          }
          else{currentCol = currentCol + 12}
        }
        else{currentCol = currentCol + 12}
        if (currentCol > 206){
          Logger.log("Could not find matching time, Working after 10PM?")
          Logger.log("Info dump: col num: "+currentCol+"  case: "+caseNum+"  Searching for time: "+str1+"  person number "+i)
          currentCol = 205
          break
        }
      }
      switch (j) {//set variables created above with their column that they start at
        case 0:
          st = currentCol
          break
        case 1:
          b1 = currentCol
          break
        case 2:
          lun = currentCol
          break
        case 3:
          b2 = currentCol
          break
        case 4:
          et = currentCol
          break
        default:
          Logger.log("++++++Error encountered") //J should never be any number besides 0 through 4, thus this should never run
          break
      }
    }
    //now that we have the start columns for the person's start, b1/2, lunch, and end times, color the cells to represent that.
    //if the column value remains 0, then they don't have that thing, be it a break, lunch, or both.
    for (j = st-1; j < et; j=j+1){//color whole shift green = #90ee90
      coverageBackgroundColors[i][j-1] = "#90ee90"
    }

    if (b1 != 0){
      for (j = b1-1; j < b1+3; j=j+1){ //color breaks and lunches blue = #add8e6
        coverageBackgroundColors[i][j-1] = "#add8e6"
      }
    }
    if (lun != 0){
      for (j = lun-1; j < lun+6; j=j+1){ //color breaks and lunches blue 
        coverageBackgroundColors[i][j-1] = "#add8e6"
      }
    }
    if (b2 != 0){
      for (j = b2-1; j < b2+3; j=j+1){ //color breaks and lunches blue
        coverageBackgroundColors[i][j-1] = "#add8e6"
      }
    }
  }
  coverageFillRange.setBackgrounds(coverageBackgroundColors)
  s2.setFrozenColumns(1)
  var botrow = s2.getLastRow() + 1
  
  //count up all the green squares in each column
  s2.getRange("A"+botrow).setValue("Coverage at current time:")
  var bgColor = "#90ee90"
  var vertical = coverageBackgroundColors.length //smaller number
  var horizontal = coverageBackgroundColors[0].length //205

  for (var i = 0; i < horizontal; i = i+1){
    var greencols = 0
    for (var j = 0; j < vertical; j = j+1){
      if(coverageBackgroundColors[j][i] == bgColor){greencols = greencols + 1}
    }
    s2.getRange(botrow,(i+2)).setValue(greencols)
  }
  
  //For different parts of the day, use different color schemes to account for smaller opening an closing staff counts
  //5AM to 6AM,     2 people normal, index 0 to 11
  //6AM to 7AM,     3 people normal, index 12 to 23
  //7AM to 7:30AM,  4 people normal, index 24 to 29
  //7:30AM to 9AM,  5 people normal, index 32 to 47
  //9AM to 8PM,     6+ people normal, index 48 to 179
  //8PM to 8:30PM,  5 people normal, index 180 to 185
  //8:30PM to 9PM,  4 people normal, index 186 to 191
  //9PM to closing, 2+ people normal, index 192 to 204

  var colorList = [[]]

  // #of people:         -5    -4    -3     -2        -1       good
  var coverageColors = ["red","red","red","orange","#85d370","green"] //6 items
  var gCount = s2.getRange(botrow,2,1,205).getValues()

  //look through number of people on staff at that time, and based on the expected normal number of people, color the cells. 
  //Weekend's early/end times are likely more red because of different store horus
  for (var i = 0; i < gCount[0].length; i = i+1){
    var idx = 0
    
    if (i < 12){//range 1 less than block max for 0 index
      //index for color = (coverageColors.length - normal people @ time block, minus 1 for 0 index) plus people currently out there
      // example here:      6 - 2 - 1 + x, = 3 + x, if 2 people, then it's green, if 1 person it's lighter green, if 0, orange. However 0 is always red.
      idx = 3 + gCount[0][i]
    }
    else{
      if((i => 12) && (i <= 23)){
        idx = 2 + gCount[0][i]
      }
      else{
        if((i > 23) && (i <= 29)){
          idx = 1 + gCount[0][i]
        }
        else{
          if((i > 29) && (i <= 47)){
            idx = gCount[0][i]
          }
          else{
            if((i > 47) && (i <= 179)){
              idx = gCount[0][i] -1
            }
            else{
              if((i > 179) && (i <= 185)){
                idx = gCount[0][i]
              }
              else{
                if((i > 185) && (i <= 191)){
                  idx = 1 + gCount[0][i]
                }
                else{
                  idx = 2 + gCount[0][i]
                }
              }
            }
          }
        }
      }
    }
    if (idx > 5 ){idx = 5} //if overstaffed, keep color as green
    if (gCount[0][i] <= 0){idx = 0} //if no one there, make it red
    colorList[0].push(coverageColors[idx])
  }
  s2.getRange(botrow,2,1,205).setBackgrounds(colorList)
}

//=======================================================================================================================================================================================
function SetPreferredNames(s1){
  var bottomrow = s1.getLastRow()-1
  var dategrab = new Date (s1.getRange("C1").getValue())
  var theDay = dategrab.getDate()

  var nameSpace = s1.getRange("B5:B"+bottomrow)
  var nameSpaceValues = nameSpace.getValues()
  for(var i = 0; i < nameSpaceValues.length; i = i+1){//for every name
    nameSpaceValues[i][0] = nameSpaceValues[i][0].toString().toUpperCase() //set it to uppercase

    var found = namesToBeFixedA.indexOf(nameSpaceValues[i][0])//if not -1, name is in the list, returns the index of it
    if (found != -1){
      nameSpaceValues[i][0] = namesToBeFixedB[found]
    }
    
    if ((theDay == 18) && (nameSpaceValues[i][0] == "LONDON, OLIVER")){
      nameSpaceValues[i][0] = "LONDON, OLUVER"
    }
  }
  nameSpace.setValues(nameSpaceValues)
}

function FitToPageAttempt(s1, bottomrow){//the sheet, last row with text
  //find out how much space is left to add empty but formatted rows
  var usedSpace = 124 //45+45+17+17, row 1/2/3/4
  var bottom = bottomrow + 2 
  var extra = 0
  for (var i = 0; i < 10; i=i+1){//find actual last row 
    if (s1.getRange('A'+(bottom+i)).isPartOfMerge()){ 
      extra = extra + 1
    }
  }

  //total up the px in rows 4 through formatted bottom 
  usedSpace = usedSpace + (22 * (bottom + extra - 4)) - 4 //the second -4 is because hours cut row is 18px, not 22px, first is to exclude the top 4 rows already counted
  var spaceRemaining = maxpixelheight - usedSpace

  if (spaceRemaining >= 0){//add rows if space allows
    var rowsToBeAdded = Math.trunc(spaceRemaining/22)
    for (var i = 0; i < rowsToBeAdded; i=i+1){
      s1.insertRowAfter(bottomrow)
      s1.getRange('A'+(bottomrow+1)+':K'+(bottomrow+1)).setBorder(true,true,true,true,true,true,"black",SpreadsheetApp.BorderStyle.SOLID)
      s1.getRange('B'+(bottomrow+1)+':C'+(bottomrow+1)).mergeAcross()
      s1.getRange('A'+(bottomrow+1)+':K'+(bottomrow+1)).setBackground("white")
      s1.getRange('L'+(bottomrow+1)).clearFormat()
    }
  }
  else{//if the spaceRemaining is negative, remove at least 1 row
    if (!dontDeleteRows){
      var rowsToBeRemoved = 1 + (Math.trunc(spaceRemaining/22) * -1)
      if (rowsToBeRemoved > 2){ //don't haphazardly remove rows if there are just an insane amount, only remove rows from hrs cut/added, which is only 6 lines long 
        rowsToBeRemoved = 2
      }
      var bottomrow = s1.getLastRow()
      for (var i = 0; i < rowsToBeRemoved; i=i+1){
        s1.deleteRow(bottomrow+1)
      }
    }
  }
  var bottomrow = s1.getLastRow()
  s1.getRange('A'+(bottomrow)+':K'+(bottomrow)).setBorder(true,true,true,true,true,true,"black",SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
}

function SecurityChecks(s1,today){//the sheet
  //add the numbers for security checks
  if(today != ""){
    if(today.length == 8){//saturday
      s1.getRange('J4:J21').setValues([['Security Checks'],['Open'],["6"],["7"],["8"],["9"],["10"],["11"],["12"],["1"],["2"],["3"],["4"],["5"],["6"],["7"],["8"],['Close']])
      s1.getRange('J4:K21').setBorder(true,true,true,true,null,null,"black",SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    }
    else{//sunday
      s1.getRange('J4:J20').setValues([['Security Checks'],['Open'],["6"],["7"],["8"],["9"],["10"],["11"],["12"],["1"],["2"],["3"],["4"],["5"],["6"],["7"],['Close']])
      s1.getRange('J4:K20').setBorder(true,true,true,true,null,null,"black",SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    }
  }
  else{//weekdays
    s1.getRange('J4:J22').setValues([['Security Checks'],['Open'],["6"],["7"],["8"],["9"],["10"],["11"],["12"],["1"],["2"],["3"],["4"],["5"],["6"],["7"],["8"],["9"],['Close']])
    s1.getRange('J4:K22').setBorder(true,true,true,true,null,null,"black",SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
  }
}

function FormatTopRows(s1,ns){//the sheet, number of sups
  //place boxes around cells that are to be used
  s1.getRange("A1:K"+s1.getLastRow()).setBorder(true,true,true,true,true,true,"black",SpreadsheetApp.BorderStyle.SOLID)
  s1.getRange('J1:K2').setBorder(true,true,true,true,true,true,"black",SpreadsheetApp.BorderStyle.SOLID)
  s1.getRange('A3:K3').clearFormat() //create blank row
  var last = s1.getLastRow()
  //box the hours cut/added                            T    L    B    R    V    H    Color   Style
  s1.getRange('A'+(last+1)+':K'+(last + 7)).setBorder(true,true,true,true,true,true,"black",SpreadsheetApp.BorderStyle.SOLID)
  
  //write text to cells that were added/cleared as needed
  var todaydate = s1.getRange("E1").getValue()
  s1.getRange('A2:K2').clearContent()
  s1.getRange('A1:K2').setValues([["","Tigard #111","",Utilities.formatDate(todaydate,Session.getScriptTimeZone(),"EEEEEEEEEEEE MM/d/yyyy"),"Daily Inspection:","Price Pro:","Door/Pad Lock Label:","Card Reader Cleaners:","","Diesel Cradles Cleaned:",""],["Open Time:","","","","","","","","","",""]])
  s1.getRange('A4').setValue('Override')
  
  //thick border on some cells
  bottomrow = s1.getLastRow()
  s1.getRange('A'+(bottomrow+2)+':K'+(bottomrow+7)).setBorder(true,true,true,true,true,null,"black",SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
  s1.getRange('B1:K1').setBorder(true,true,true,true,true,true,"black",SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
  s1.getRange('A2').setBorder(true,true,true,true,null,null,"black",SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
  if(ns > 0){
    s1.getRange('A5:I'+(4+ns)).setBorder(null,null,true,null,null,null,"black",SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
  }
}

function FillBreakAid(s1,timeArray,valArray){ //the sheet, the Range from D19 to I(last row), the values in said Range
  //clear out the middle of whatever workforce placed in there, it's trash =), and gets in the way of other functions.
  var clearing = s1.getRange("F19:H"+s1.getLastRow())
  clearing.clearContent().clearFormat()
  clearing.setBorder(true,true,true,true,true,true,"black",SpreadsheetApp.BorderStyle.SOLID)
  var arrlen = valArray.length
  for (var i = 0; i < arrlen; i = i+1){ //populate shifts with times based on their length
    var starttime = new Date(valArray[i][1])
    var endtime = new Date(valArray[i][5])
    var shifthours = (endtime.getTime()-starttime.getTime()) / hourmillis
    var hoursrounded = Math.trunc(shifthours)
    var flg = false
    var segHrs = valArray[i][0].toString()
    if ((segHrs.length > 11) || (segHrs[0] == "9")){
      Logger.log("9+ hour shift? Fill in manually.")
      continue
    }
    if (segHrs[0] != segHrs[5]){ //check if the whole number hours are not the same i.e. the 8 and 4 here: 8.00/4.00
      flg = true
    }
    if (shifthours >= 5){
      shifthours = shifthours - 0.5 //remove the lunch from the time
    }
    switch (hoursrounded){
      //Over 10 hours = 3B 2L(2nd lunch waivable), over 12 hours 3B 2L. These are almost never scheduled, so not including. Same with 9hr shifts.  
      case 8: //2B 1L
        EightHours(starttime,i,shifthours,valArray)
        break
      case 7: //2B 1L
        SevenHours(starttime,i,shifthours,valArray,flg,s1,segHrs)
        break
      case 6: //1B 1L if full shift(clock in to out) is 6.0, any more: 2B 1L
        SixHours(starttime,i,shifthours,valArray,flg,s1,segHrs)
        break
      case 5: //1B if full shift(clock in to out) is 5.0h, any more: 1B 1L 
        FiveHours(starttime,i,shifthours,valArray,flg,s1,segHrs)
        break
      case 4: //1B
        FourHours(starttime,i,shifthours,valArray,s1,segHrs)
        break
      case 3: //1B
        ThreeHours(starttime,i,shifthours,valArray,s1)
        break
      default: //2h shifts get 1B, but 99.9% of the time they're coming from a diff department, so break before or after transfer, thus none @ gas
        BlankOut(i,valArray,shifthours,s1,segHrs)
        break
    }
  }
  timeArray.setValues(valArray)
}

function SortPeople(s1,tr,br,ns,stl){ //the sheet, top row, bottom row, number of sups, format style
  //sorting by end time (excluding sups)
  var rowtobemoved = 0
  var rowtobeplacedafter = 0
  
  //Sort by start time
  if (stl == "start"){
    for (var i = tr+1+ns; i < br+1; i = i+1){//check every time pair, see if latter is before former
      var v1 = s1.getRange('E'+i).getValue()
      var v2 = s1.getRange('E'+(i+1)).getValue()
      if (v1 > v2){ //if so, move latter up until it isn't placed after a later time
        rowtobemoved = i+1
        rowtobeplacedafter = i-1
        for (var j = i-1; j > tr+ns; j = j-1){
          var v3 = s1.getRange('E'+j).getValue()
          if (v3 > v2){
            rowtobeplacedafter = j-1
          }
          else{
            break
          }
        }
        s1.moveRows(s1.getRange('A'+rowtobemoved),rowtobeplacedafter+1)
        //after swap:
        //if start times are the same, ensure the one who ends earlier is on top.
        v1 = s1.getRange("E"+rowtobeplacedafter).getValue() //one above
        v2 = s1.getRange("E"+(rowtobeplacedafter+1)).getValue() //self
        if (v1.valueOf() == v2.valueOf()){ //if start times are the same, ensure the one who ends earlier is on top.
          var vx = s1.getRange('I'+rowtobeplacedafter).getValue() //end of one above
          var vy = s1.getRange('I'+(rowtobeplacedafter+1)).getValue() //end of self
          if (vx > vy){ //if end time of above is larger(later), swap
            s1.moveRows(s1.getRange('A'+(rowtobeplacedafter+1)),rowtobeplacedafter)
          }
        }
      }
      else{
        if (v1.valueOf() == v2.valueOf()){ //if start times are the same, ensure the one who ends earlier is on top.
          var vx = s1.getRange('I'+i).getValue()
          var vy = s1.getRange('I'+(i+1)).getValue()
          if (vx > vy){
            s1.moveRows(s1.getRange('A'+(i+1)),i)
          }
        }
      }
    }
  }

  //Sort by end time
  if (stl == "end"){
    for (var i = tr+1+ns; i < br+1; i = i+1){//check every time pair, see if latter is before former
      var v1 = s1.getRange('I'+i).getValue()
      var v2 = s1.getRange('I'+(i+1)).getValue()
      if (v1 > v2){ //if so, move latter up until it isn't placed after a later time
        rowtobemoved = i+1
        rowtobeplacedafter = i-1
        for (var j = i-1; j > tr+ns; j = j-1){
          var v3 = s1.getRange('I'+j).getValue()
          if (v3 > v2){
            rowtobeplacedafter = j-1
          }
          else{
            break
          }
        }
        s1.moveRows(s1.getRange('A'+rowtobemoved),rowtobeplacedafter+1)
        //after swap:
        //if end times are the same, ensure the one who starts earlier is on top.
        v1 = s1.getRange("I"+rowtobeplacedafter).getValue() //one above
        v2 = s1.getRange("I"+(rowtobeplacedafter+1)).getValue() //self
        if (v1.valueOf() == v2.valueOf()){ //if end times are the same, ensure the one who starts earlier is on top.
          var vx = s1.getRange('E'+rowtobeplacedafter).getValue() //start of one above
          var vy = s1.getRange('E'+(rowtobeplacedafter+1)).getValue() //start of self
          if (vx > vy){ //if start time of above is larger(later), swap
            s1.moveRows(s1.getRange('A'+(rowtobeplacedafter+1)),rowtobeplacedafter)
          }
        }
      }
      else{
        if (v1.valueOf() == v2.valueOf()){ //if end times are the same, ensure the one who starts earlier is on top.
          var vx = s1.getRange('E'+i).getValue()
          var vy = s1.getRange('E'+(i+1)).getValue()
          if (vx > vy){
            s1.moveRows(s1.getRange('A'+(i+1)),i)
          }
        }
      }
    }
  }
  
  //Sort Hybrid
  if (stl == "hybrid"){
    var peopleCount = br - (tr+ns) + 1 //number of non-sups working that day
    var halfPeople = Math.trunc((peopleCount / 2)) - Math.trunc(hybridBias) // rounds down, the more negative the bias, the more of the rows are sorted by end time. bias = 0 to split down middle
    
    //first, sort the whole thing by start time to find the half that arrives earliest
    for (var i = tr+1+ns; i < br+1; i = i+1){//check every time pair, see if latter is before former
      var v1 = s1.getRange('E'+i).getValue()
      var v2 = s1.getRange('E'+(i+1)).getValue()
      if (v1 > v2){ //if so, move latter up until it isn't placed after a later time
        rowtobemoved = i+1
        rowtobeplacedafter = i-1
        for (var j = i-1; j > tr+ns; j = j-1){
          var v3 = s1.getRange('E'+j).getValue()
          if (v3 > v2){
            rowtobeplacedafter = j-1
          }
          else{
            break
          }
        }
        s1.moveRows(s1.getRange('A'+rowtobemoved),rowtobeplacedafter+1)
        //after swap:
        //if start times are the same, ensure the one who ends earlier is on top.
        v1 = s1.getRange("E"+rowtobeplacedafter).getValue() //one above
        v2 = s1.getRange("E"+(rowtobeplacedafter+1)).getValue() //self
        if (v1.valueOf() == v2.valueOf()){ //if start times are the same, ensure the one who ends earlier is on top.
          var vx = s1.getRange('I'+rowtobeplacedafter).getValue() //end of one above
          var vy = s1.getRange('I'+(rowtobeplacedafter+1)).getValue() //end of self
          if (vx > vy){ //if end time of above is larger(later), swap
            s1.moveRows(s1.getRange('A'+(rowtobeplacedafter+1)),rowtobeplacedafter)
          }
        }
      }
      else{
        if (v1.valueOf() == v2.valueOf()){ //if start times are the same, ensure the one who ends earlier is on top.
          var vx = s1.getRange('I'+i).getValue()
          var vy = s1.getRange('I'+(i+1)).getValue()
          if (vx > vy){
            s1.moveRows(s1.getRange('A'+(i+1)),i)
          }
        }
      }
    }
    //second, starting from half way down, sort by end time
    var lowestIndexRow = br + 2
    for (var i = br+1; i > tr+1+ns+halfPeople; i = i-1){
      //Look at the bottom end time, search up for any later end times, move said later end time to the bottom after making sure it's the latest of all end times from midpoint down
      //then move up one row as bottom is now latest end time, and repeat.
      var heldLatestEndTime = s1.getRange("I"+i).getValue() //grab bottom row's time
      var heldLatestIndex = i

      for (var j = i-1; j > tr+1+ns+halfPeople; j = j-1){//look through all above times
        var timeAbove = s1.getRange("I"+j).getValue()
        if (timeAbove > heldLatestEndTime){//if found new latest time, put that in hand.
          heldLatestEndTime = timeAbove
          heldLatestIndex = j
        }
      }
      //have latest time in hand, move it to bottom unless it's already there
      if (!((lowestIndexRow - 1) == heldLatestIndex)){ //if [NOT](swap_row == held_row)
        s1.moveRows(s1.getRange("A"+heldLatestIndex),lowestIndexRow)
      }
      
      lowestIndexRow = lowestIndexRow - 1 //shift lowest row up 1 for next lowest row

      //Don't need to swap matching end times so that start time is earlier, because it was sorted by start times before, the later start will always be on the bottom for matching end times.
    }
  }
}

function Resizer(page){ //the sheet
  //set row/column heights/widths, also does fonts.
  
  bottomrow = page.getLastRow()-1
  page.getRange('B4:C'+bottomrow).mergeAcross()
  //make all names all uppercase
  var nameSpace = page.getRange("B5:B"+bottomrow)
  var nameSpaceValues = nameSpace.getValues()
  for(var i = 0; i < nameSpaceValues.length; i = i+1){
    nameSpaceValues[i][0] = nameSpaceValues[i][0].toString().toUpperCase()
  }
  nameSpace.setValues(nameSpaceValues)
  
  //make text a little bigger
  page.getRange('A5:I'+bottomrow).setFontSize(11)
  page.getRange('B5:I'+bottomrow).setHorizontalAlignment('center').setVerticalAlignment('middle')
  
  //merge some cells 
  bottomrow = page.getLastRow()
  page.getRange('A'+bottomrow+':D'+(bottomrow+6)).mergeAcross()
  page.getRange('E'+bottomrow+':H'+(bottomrow+6)).mergeAcross()
  page.getRange('I'+bottomrow+':K'+(bottomrow+6)).mergeAcross()
  page.getRange('H1:I2').mergeAcross()
  page.getRange('J1:K2').mergeAcross()
  page.getRange('B2:C2').mergeAcross()
  page.getRange('J4:K4').mergeAcross()
  page.getRange('D1').moveTo(page.getRange('C1'))
  page.getRange('C1:D1').mergeAcross()
  page.getRange('B2:D2').mergeAcross()
  
  //format the top rows
  page.getRange('A1:K2').setVerticalAlignment('middle').setHorizontalAlignment('center').setWrap(true)
  page.setRowHeights(1,2,45)
  page.setColumnWidth(10,50)
  
  //shrink rows until it all fits on a standard page
  page.setColumnWidth(1,70)
  page.setColumnWidth(2,90)
  page.setColumnWidth(4,90)
  page.setColumnWidths(5,9,90)
  page.setColumnWidth(10,70)
  page.setRowHeights(4,50,22)
  
  //Set font and sizes
  var rawdata2 = page.getDataRange()
  rawdata2.setFontFamily("Calibri")
  page.getRange("A4:K4").setFontSize(10)
  page.getRange("A1:K2").setFontSize(11)
  page.getRange("A"+bottomrow+":K"+bottomrow).setFontSize(11)
  page.setRowHeights(3,2,17)
  page.setRowHeights(page.getLastRow(),1,18)
}

function EightHours(st,currentrow,sh,timeArray){//start time, current row being worked on, shift hours (at gas, minus lunch if exists), the sheet 
  var startTimeMS = st.getTime()
  var addedMS = 0
  
  //from (Col F; through Col H; move over column){
  for (var i = 2; i <= 4; i = i+1){
    
    addedMS = addedMS + (2 * hourmillis)    //push time 2hr forward for next break/lunch
    if (i == 4){//last break      //add 15/30 mins to place break 2 correctly
      addedMS = addedMS + halfhourmillis //account for lunch
      if (sh < 8){ //8.25 or 8.0 hours from clock in to clock out opposed to the usual 8.5
        addedMS = addedMS - fifteenminmillis
      }
    }
    timeArray[currentrow][i] = Utilities.formatDate(new Date(startTimeMS + addedMS),"GMT-8:00","''h:mm a")
  }

  FixStartEndTimes(currentrow,timeArray,0)
}

function SevenHours(st,currentrow,sh,timeArray,flag,page,totalhours){//start time, current row being worked on, shift hours (at gas, minus lunch), Array of vals to fill, flag for not all at gas, the sheet,seghours:8.00/7.00
  var cell
  var startTimeMS = st.getTime()
  var addedMS = 0
  
  if (flag){  //if not whole shift is at gas, then treat like a normal 8.0 hour shift, without first break
    page.getRange(currentrow+19,6).setBackground('#d0cece') //blank out first break

    for (var i = 3; i <= 4; i = i+1){ //lunch and break
      addedMS = addedMS + (2 * hourmillis) //default break gap of 2 hours

      if (i == 4){ //last break
        addedMS = addedMS + halfhourmillis //account for lunch
      }
      timeArray[currentrow][i] = Utilities.formatDate(new Date(startTimeMS + addedMS),"GMT-8:00","''h:mm a")
    }
    if (giveWarning){
      cell = page.getRange(currentrow+19,12)
      cell.setValue("Check this row").setBackground("Red")
    }
  }
  else{
    for (var i = 2; i <= 4; i = i+1){
      addedMS = addedMS + (2 * hourmillis)
      if (i == 3){ //lunch
        if(sh == 6.5){addedMS = addedMS - hourmillis} //7.0h @ work
        else{addedMS = addedMS - fifteenminmillis} //7.25/7.5/7.75h @ work
      }
      if (i == 4){
        addedMS = addedMS + halfhourmillis //account for lunch
        if (sh == 6.5){ //earlier break for shorter shift
          addedMS = addedMS - fifteenminmillis
        }
      }
      timeArray[currentrow][i] = Utilities.formatDate(new Date(startTimeMS + addedMS),"GMT-8:00","''h:mm a")
    }
  }
  if((tryToFixSentFromFE)&&(totalhours[0] != 7)){
    FixStartEndTimes(currentrow,timeArray,fifteenminmillis)
    if(tellChanges){page.getRange("M"+(currentrow+19)).setValue("Pushed Start time 15")}
  }
  else{FixStartEndTimes(currentrow,timeArray,0)}
}

function SixHours(st,currentrow,sh,timeArray,flag,page,totalhours){//start time, current row being worked on, shift hours (at gas, minus lunch), Array of vals to fill, flag for not all at gas, the sheet,seg hours=8.00/6.00
  var cell
  var startTimeMS = st.getTime()
  var addedMS = 0

  if (flag){  //if not whole shift is at gas, but at least 6.5 at gas
              //then treat like a normal 8.0 hour shift, without first break
    page.getRange(currentrow+19,6).setBackground('#d0cece') //blank out first break

    for (var i = 3; i <= 4; i = i+1){ //lunch and break
      addedMS = addedMS + (2 * hourmillis) //default break gap of 2 hours

      if (i == 4){ //last break
        addedMS = addedMS + halfhourmillis //account for lunch
      }
      timeArray[currentrow][i] = Utilities.formatDate(new Date(startTimeMS + addedMS),"GMT-8:00","''h:mm a")
    }
    if (giveWarning){
      cell = page.getRange(currentrow+19,12)
      cell.setValue("Check this row").setBackground("Red")
    }
  }
  else{ //whole shift at gas
    for (var i = 2; i <= 4; i = i+1){
      addedMS = addedMS + (2 * hourmillis) - halfhourmillis //default break gap of 2 hours, minus 30 becaues of shift size

      if (sh == 5.5){ //6.0h @ work
        if (i == 2){
          page.getRange(currentrow+19,6).setBackground('#d0cece') //blank out first break, 6.0h shift only
          timeArray[currentrow][i] = ""
          continue //next loop
        }
        if (i == 3){addedMS = addedMS - hourmillis} //2 full hours
        if (i == 4){addedMS = addedMS + fifteenminmillis + halfhourmillis}

      }
      else{
        if (sh == 6){ //6.5h @ work
          if (i == 2){
            page.getRange(currentrow+19,6).setBackground('#d0cece') //blank out first break, 6.0h shift only
            timeArray[currentrow][i] = ""
            continue //next loop
          }
          if (i == 3){addedMS = addedMS - hourmillis} //2 full hours
          if (i == 4){addedMS = addedMS + hourmillis}
        }
        else{
          if (i == 4){//last break
            if (sh == 5.75){addedMS = addedMS - fifteenminmillis}//6.25h shift
            //if (sh == 6){addedMS = addedMS}                    //6.5h shift
            if (sh == 6.25){addedMS = addedMS + fifteenminmillis}//6.75h shift
          }
        }
        
      }
      timeArray[currentrow][i] = Utilities.formatDate(new Date(startTimeMS + addedMS),"GMT-8:00","''h:mm a")
    }
  }
  if((tryToFixSentFromFE)&&(totalhours[0] != 6)){
    FixStartEndTimes(currentrow,timeArray,fifteenminmillis)
    if(tellChanges){page.getRange("M"+(currentrow+19)).setValue("Pushed Start time 15")}
  }
  else{FixStartEndTimes(currentrow,timeArray,0)}
}

function FiveHours(st,currentrow,sh,timeArray,flag,page,totalhours){//start time, current row being worked on, shift hours (at gas, minus lunch), Array of vals to fill, flag for not all at gas, the sheet,seghours:8.00/5.00
  var cell
  var startTimeMS = st.getTime()
  var addedMS = 0

  page.getRange(currentrow+19,6).setBackground('#d0cece')
  if (flag){  //if not whole shift is at gas, then treat like a normal 8.0 hour shift, without first break
    for (var i = 3; i <= 4; i = i+1){ //lunch then break
      addedMS = addedMS + (2 * hourmillis) //default break gap of 2 hours

      if (i == 4){ //last break
        addedMS = addedMS + halfhourmillis //account for lunch
      }
      timeArray[currentrow][i] = Utilities.formatDate(new Date(startTimeMS + addedMS),"GMT-8:00","''h:mm a")
    }
    if (giveWarning){
      cell = page.getRange(currentrow+19,12)
      cell.setValue("Check this row").setBackground("Red")
    }
  }
  else{
    for (var i = 3; i <= 4; i = i+1){
      cell = page.getRange(currentrow+1,i)
      addedMS = addedMS + (hourmillis + halfhourmillis) //90 mins added
      if (sh == 4.5){ //5.0h shift, 1B 0L
        if (i == 3){
          page.getRange(currentrow+19,7).setBackground('#d0cece')
          continue //next loop
        }
        if (i == 4){
          addedMS = addedMS - halfhourmillis //2.5h segments (2 90m segments = 3h, -30m = 2.5h)
          timeArray[currentrow][i] = Utilities.formatDate(new Date(startTimeMS + addedMS),"GMT-8:00","''h:mm a")
          continue
        }
      }
      if (i == 4){addedMS = addedMS + halfhourmillis} //account for lunch

      if (sh == 5.25){addedMS = addedMS + fifteenminmillis} //5.75h shift
      if ((sh == 5) && (i == 3)){ //5.5h shift
        addedMS = addedMS + fifteenminmillis  //push lunch 15m
      } 
      if ((sh == 4.75) && (i == 4)){addedMS = addedMS + fifteenminmillis} //5.25h shift

      timeArray[currentrow][i] = Utilities.formatDate(new Date(startTimeMS + addedMS),"GMT-8:00","''h:mm a")
    }
  }
  if((tryToFixSentFromFE)&&(totalhours[0] == 6)){
    FixStartEndTimes(currentrow,timeArray,halfhourmillis)
    if(tellChanges){page.getRange("M"+(currentrow+19)).setValue("Pushed Start time 30")}
  }
  else{
    if((tryToFixSentFromFE)&&(totalhours[0] == 8)){
      FixStartEndTimes(currentrow,timeArray,fifteenminmillis)
      if(tellChanges){page.getRange("M"+(currentrow+19)).setValue("Pushed Start time 15")}
    }
    else{FixStartEndTimes(currentrow,timeArray,0)}
  }
}

function FourHours(st,currentrow,sh,timeArray,page,totalhours){//start time, current row being worked on, shift hours (at gas, minus lunch if exists), Array of vals to fill, the sheet, 8.00/4.00 box
  var startTimeMS = st.getTime()
  var addedMS = (2 * hourmillis)
  page.getRange(currentrow+19,6,1,2).setBackground('#d0cece') //blank out first break and lunch
  if ((sh == 4.5 || sh == 4.75) && (totalhours[0] != "8")){addedMS = addedMS + fifteenminmillis}
  timeArray[currentrow][4] = Utilities.formatDate(new Date(startTimeMS + addedMS),"GMT-8:00","''h:mm a")
  if((tryToFixSentFromFE)&&(totalhours[0] != 4)){
    FixStartEndTimes(currentrow,timeArray,halfhourmillis)
    if(tellChanges){page.getRange("M"+(currentrow+19)).setValue("Pushed Start time 30")}
  }
  else{FixStartEndTimes(currentrow,timeArray,0)}
}

function ThreeHours(st,currentrow,sh,timeArray,page){//start time, current row being worked on, shift hours (at gas, minus lunch if exists), Array of vals to fill, the sheet
  var startTimeMS = st.getTime()
  addedMS = (hourmillis + halfhourmillis) //90 mins added
  page.getRange(currentrow+19,6,1,2).setBackground('#d0cece')

  if (sh == 3.5 || 3.75){addedMS = addedMS + fifteenminmillis}
  timeArray[currentrow][4] = Utilities.formatDate(new Date(startTimeMS + addedMS),"GMT-8:00","''h:mm a")
  FixStartEndTimes(currentrow,timeArray,0)
}

function BlankOut(currentrow,timeArray,sh,page,segHrs){//current row being worked on, Array of vals to fill, end time, shift hours (at gas, minus lunch if exists), the sheet
  page.getRange(currentrow+19,6,1,3).setBackground('#d0cece')
  if (giveWarning){//this code shouldn't run as I completely skip 9/10 hour shifts, but just in case
    if (sh >= 8.5){//if for some reason someone is scheduled 9+ hours
      page.getRange(currentrow+19,12).setValue("Check this row").setBackground("Red")
    }
  }
  if((tryToFixSentFromFE) && (segHrs[0] != 2)){//if more than 2 hour shift. bump start time for break
    FixStartEndTimes(currentrow,timeArray,fifteenminmillis)
    if(tellChanges){page.getRange("M"+(currentrow+19)).setValue("Pushed Start time 15")}
  } 
  else{FixStartEndTimes(currentrow,timeArray,0)}
}

function FixStartEndTimes(cr,timeArray,pushingStartAmount){//current row, array with the cell values between D19 and I(last row)
  timeArray[cr][1] = Utilities.formatDate(new Date(timeArray[cr][1].getTime()+pushingStartAmount),"GMT-8:00","''h:mm a")
  timeArray[cr][5] = Utilities.formatDate(new Date(timeArray[cr][5].getTime()),"GMT-8:00","''h:mm a")
}
