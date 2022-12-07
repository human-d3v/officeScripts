function main(workbook: ExcelScript.Workbook) {
    //define the current sheet
      const ws = workbook.getActiveWorksheet();
      ws.setName("Full Data");
    //create an array of objects usedRange




    function csvJSON(rng: ExcelScript.Range){
      //get used range to loop
        const r = rng.getUsedRange();
        const ws = workbook.getActiveWorksheet()
        //define header row and parse the data from the top row
          const hdrs: string[] = []; //define header row
          const lCol = r.getLastColumn().getColumnIndex();//find last column
          const lRow = r.getLastRow().getRowIndex(); //get last row

          //loop through headers and add them to an array to get the unique values.
            for (let i = 1; i<=lCol;i++){
                const obj: object={};
                const h:string = ws.getCell(1,i).getText();
                hdrs.push(h);
            };
          
          //loop through the rows and add each value to the array
          


      
    }
}