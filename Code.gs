class mergePair{
  constructor(topLeft, bottomRight, length, width){
    this.topLeft = topLeft;
    this.bottomRight = bottomRight;
    this.length = length;
    this.width = width;
  }

}

class cellPosition{
  constructor(row, column){
    this.row = row;
    this.column = column;
  }
}


class mergeProperty{
  constructor(stateOfMerge, length, width){
    this.stateOfMerge = stateOfMerge;
    this.length = length;
    this.width = width;
  }
}

const mergeState ={
  NotMerged: 'notmerged',
  Head: 'head',
  Body: 'body'
}




function isInMerge(row, column, mergePairs){
  for(var i = 0; i < mergePairs.length; i++){
    if((row == mergePairs[i].topLeft.row) && (column == mergePairs[i].topLeft.column)){
      return new mergeProperty(mergeState.Head, mergePairs[i].length, mergePairs[i].width);
    }
    else if(row >= mergePairs[i].topLeft.row && row <= mergePairs[i].bottomRight.row && column >= mergePairs[i].topLeft.column && column <= mergePairs[i].bottomRight.column){
      return new mergeProperty(mergeState.Body, mergePairs[i].length, mergePairs[i].width);
    }
  }
    
      return new mergeProperty(mergeState.NotMerged, 1, 1);
    
  }

 


function preProcessMerges(dataRange) {
  var mergedRanges = dataRange.getMergedRanges();
  const mergePairs = [];
  for(var i = 0; i < mergedRanges.length; i++){
    mergePairs.push(new mergePair(new cellPosition(mergedRanges[i].getRow()-1, mergedRanges[i].getColumn()-1),
     new cellPosition(mergedRanges[i].getLastRow()-1,
      mergedRanges[i].getLastColumn()-1), mergedRanges[i].getLastRow()-mergedRanges[i].getRow(),
       mergedRanges[i].getLastColumn()-mergedRanges[i].getColumn()));
        }

      
return mergePairs;

    
    }


function getWinner(row, allValues, numberColumns){
var tempHolder = -1;
var tempColHolder = -1;
for(var i = 0; i < numberColumns.length; i++){
  if(allValues[row][numberColumns[i]] > tempHolder){
    tempHolder = allValues[row][numberColumns[i]];
    tempColHolder = numberColumns[i];
    //console.log(tempColHolder + " " + tempHolder)
  }
  else{
  }
}
var winnerName = allValues[0][tempColHolder];
const winnerArr = winnerName.split("\n")
var winnerParty = winnerArr[1];
return winnerParty;
}

function getNumberColumns(allValues, symb){
  const numberColumns= [];
  for(var i = 0; i < allValues[1].length; i++){
    if(allValues[1][i] == symb){
      numberColumns.push(i);
    }
  }
  return numberColumns;
}

function myFunction() {


    var isSortable = true;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Sheet1");
    var dataRange = sheet.getDataRange()
    var allValues = dataRange.getValues();
    const numOfRow = dataRange.getNumRows();
    const numOfCol = dataRange.getNumColumns();
    var mergedPairs = preProcessMerges(dataRange);
    var numberColumns = getNumberColumns(allValues, '#');
    var percentColumns = getNumberColumns(allValues, '%');

    var output = "";
    if(isSortable){
      output += '{| class="wikitable sortable"\n';
    }
    else{
      output += '{|\n';
    }
    output += '|+ County results\n'
    var r = 0;
    var c = 0;



    // Iterates through the rows
    while(r < numOfRow){

      // Iterates through the columns
      while(c < numOfCol){
        
        
          var propertyOfMerge = isInMerge(r,c,mergedPairs);
          if(propertyOfMerge.stateOfMerge != mergeState.NotMerged){
            if(allValues[r][c] != ""){
            if(propertyOfMerge.length > 0){
               output += '!rowspan="' + (propertyOfMerge.length + 1) + '" ';
            }
            else if(propertyOfMerge.width > 0){
              output += '!colspan="' + (propertyOfMerge.width + 1) + '" ';
            }
            if(r == 0){
              output += 'scope="col" '
            }
            output += "|" + allValues[r][c] + "\n";
            }

            c++;
          }
       
         
        
        else{
          if(r > 1 && r < numOfRow){
          for(var i = 0; i < percentColumns.length; i++){
            if(c == percentColumns[i]){
              allValues[r][c] = allValues[r][c]*100;
              allValues[r][c] = parseFloat(allValues[r][c].toString()).toFixed(2).toString() + '%';
              
            }
          }
          }
          if(allValues[r][c] != null){
          if(r == numOfRow - 1 || r == 1){
            output += '!'
          }
          else if(r != 0){
            output += '| '
            if(c == 0){
              output += '{{party shading/' + getWinner(r, allValues, numberColumns) + '}} '
            }
            output += 'align="center" |'
          }
          else{
            output += '|'
          }
          output += allValues[r][c] + '\n';
          }
        
        c++;
        }
      }
      if(r != numOfRow - 1){
        output += '|-\n';
      }
      
              c = 0;
              r++;
      

    }
    output += '|}'
    return(output);

    
      
    }

    function doGet(){
        return HtmlService.createHtmlOutputFromFile('ui.html').append('    <textarea id="w3review" name="w3review" rows="50" cols="50">' + myFunction() + '</textarea>');
    }

    
  
