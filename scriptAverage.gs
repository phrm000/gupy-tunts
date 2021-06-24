function getDataAndPrint() {

  var sheetActive = SpreadsheetApp.openById("16j8vOBVyqbYnnSTyo5AAlqRxeeQ0Dzxgvvz7anwbhQk");  // Get sheet as a intire document
  var sheet = sheetActive.getSheetByName("engenharia_de_software"); // Select only a page by the name
  sheet.getRange(4, 7,24,2).setValue(""); // Clear if any of the results cell have content

 

  // Defining values
  var counterRows = 4; // Starting counter row
  var counterColumns = 4; // Starting counter column
  var classesTotal = sheet.getRange(2, 1).getValue().replace("Total de aulas no semestre:",""); //Get numbers of classes


  
  while(!sheet.getRange(counterRows,1).isBlank()){ //A loop that stops when it finds a empty cell

    verifyAbsences(sheet,counterRows,classesTotal); // Calls a function to check absences from the student
    
    if(sheet.getRange(counterRows,7).isBlank()) // If there's no content on the result cell as Failed on absences, the program will verify the gradde from the student
      verifyAverage(sheet,counterRows,counterColumns);
    

    
    counterRows++;// When finished, the next row is analyzed
  
  }

}


function verifyAbsences(sheet,counterRows,classesTotal) { // The function verifies throught math the percentage of absences

  var absences = sheet.getRange(counterRows, 3).getValue(); // The number of absences are got

  var absencesRatio = 100 * absences /classesTotal;  // The percentage math operation is done
  absencesRatio = +absencesRatio.toFixed(0); // It's converted to a whole number

  if(absencesRatio > 25){ //Checks if the limits are exceded

    //If true, fill the cell with the information
    sheet.getRange(counterRows, 7).setValue("Reprovado por Falta");
    sheet.getRange(counterRows, 8).setValue(0);

  }

}

function verifyAverage(sheet,counterRows,counterColumns) { // The function verifies throught math the average of the student

  // Variables to the math operation
  var averageGrade = 0;
  var averageCounter = 0;
  var grade = 0;

  while(!sheet.getRange(counterRows, counterColumns).isBlank()){ // Counts how many columns are not empty

    grade = sheet.getRange(counterRows, counterColumns).getValue() // Get the value
    averageGrade = averageGrade + grade; // Adds the next to the existent one
    counterColumns++; // Increses the counter to get the next column value
      
      
  }

  averageCounter = counterColumns - 4; // The counter - 4 represents the divider of the average
  averageGrade = averageGrade / averageCounter; // The average of the sum of all the grades by the number of them 

  if(averageGrade < 50){ // If the average grade is under 50 the student disapproved with no chance to the final exam

    sheet.getRange(counterRows, 7).setValue("Reprovado por Nota");
    sheet.getRange(counterRows, 8).setValue(0);

  }

  else if(averageGrade >= 50 && averageGrade < 70){ // If the average grade is under 70 but greater or equal than 50 the student has the chance to take the final exam

    sheet.getRange(counterRows, 7).setValue("Exame Final");

    averageGrade = +averageGrade.toFixed(0); 
    var finalGrade =100 - averageGrade ;
    sheet.getRange(counterRows, 8).setValue(finalGrade);

  }

  else{ // If the average grade is greater than 70 the student approved 
    sheet.getRange(counterRows, 7).setValue("Aprovado");
    sheet.getRange(counterRows, 8).setValue(0);

  }


}
