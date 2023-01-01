//define table as variable
    const tbl:ExcelScript.Table = workbook.getWorksheet("Sheet1").getTables()[0];
//get full range
    const tblRng:ExcelScript.Range = tbl.getRange(); 

//find total number of columns in table
    const colCnt:number = tblRng.getColumnCount(); 

/*
resize range and assign it to a new variable.
getResizedRange() Expects positive integers to increase the size of range and negative numbers to decrease. Subtracting the total number of columns from 2 will allow you to resize your image range to only two columns. 
*/

    let newRng:ExcelScript.Range = tblRng.getResizedRange(0,(2-colCnt));
    const tblImg:string = newRng().getImage();
