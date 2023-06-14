/*

In Office typescript library, the range.setFormulasR1C1() method has laughably little documentation. The method requires an array of arrays, each sub-array with it's own copy of the formula to be used, or barring that, each sub-array requires a separate function. 

*/

//Here's a function that can accept an excel function string and create the desired array, it accepts two arguments: f <- the formula to be used; n <- the number of iterations needed

function frmlaArrConstructor(f:string,n:number):string[][]{
    const arr:string[][] = [];
    for (let i:number = 0; i<n; i++){
        arr[i] = [];
        for(let j:number = 0; j<1; j++){
            arr[i][j] = f;
        };
    };
    return arr;
};

//The initialization of the above function looks like this:

let sheet: ExcelScript.Worksheet = workbook.getActiveWorksheet();
let sheetName:string = sheet.getName()
let fArray:string[][] = frmlaArrConstructor(`='${sheetName}'!RC[-1]`,5); //using an object literal to capture a sheet name.
console.log(fArray);

/* The output looks like this:
    [
    ["='Sheet1'!RC[-1]"],
    ["='Sheet1'!RC[-1]"],
    ["='Sheet1'!RC[-1]"],
    ["='Sheet1'!RC[-1]"],
    ["='Sheet1'!RC[-1]"],
    ]
*/

//In the context of the .formulasR1C1() method, it can be used to add the formula to an entire range of cells:

let ws2: ExcelScript.Worksheet = workbook.getWorksheet("Data")
let lastRow:number = workbook.getActiveWorksheet().getUsedRange().getLastRow().getRowIndex();

workbook.getActiveWorksheet().getRange(`B2:B${lastRow}`).setFormulasR1C1(frmlaArrConstructor(`=COUNTIF('${ws2.getName()}'!C[2],RC[-1])`,lastRow - 1)); //offset the last row to account for starting in B2.

//here's another example:

function main(wb: ExcelScript.Workbook) {
    const ws:ExcelScript.Worksheet = wb.getWorksheet("Orders"); 

    const ws2:ExcelScript.Worksheet = wb.addWorksheet('Filtered_Values');

    ws2.getRange("A2").setFormula(`=UNIQUE(SORT('${ws.getName()}'!C:C))`); //uses excel's native UNIQUE() array function to populate column A from the 'Orders' sheet column C

    let ws2LRow:number = ws2.getUsedRange().getLastRow().getRowIndex(); //find the new last row of the unique spill range

    ws2.getRange(`B2:B${ws2LRow}`).setFormulasR1C1(frmlaArrConstructor(`=AVERAGEIF('${ws.getName()}'!C[1],RC[-1],'${ws.getName()}'!C[8])`,ws2LRow -1));
    
    //add headers to the sheet
        let hdrs: string[]= ['Country of Sale','Average Sale by Country'];
        let i:number = 0;
        for (let h of hdrs){
            ws2.getCell(0,i).setValue(h);
            i +=1;
        };
        

    function frmlaArrConstructor(f:string,startRowIdx:number,endRowIdx:number):string[][]{
        const arr:string[][]=[];
        for(let n:number = startRowIdx; n<=endRowIdx; n++){
            arr.push([f]);
        };
        return arr;
    };
    
};

/*
One of the benefits of using the .setFormulasR1C1() method over the .setFormulas() method is the benefit of relative references. If you were to use the frmlaArrConstructor function to create
an array for the .setFormulas() method, it's certainly possible, just a more complicated. It would require a sort of symbol not used in excel formulas that could be used with the indexOf() method.
Here's an example: 
*/

function frmlaArrConstructorV2(f:string,n:number,startRow:number):string[][]{
    const arr:string[][] = [];
    let r:number = startRow;
    for (let i:number = 0; i<n; i++){
        arr[i] = [];
        let s:string = f.replace('|',r.toString()); //replace the specified character (in this case '|') with the row number incremented every loop.
        for(let j:number = 0; j<1;j++){
            arr[i][j] = s;
        };
        r+=1;
    };
    return arr;
}

//So the following equation:

var formArray:string[][] = frmlaArrConstructorV2(`=POWER(($A$| * PI()),2)`,5,2);

//would have the following output:

/*
    [
        ["=POWER(($A$2 * PI()),2)"],
        ["=POWER(($A$3 * PI()),2)"],
        ["=POWER(($A$4 * PI()),2)"],
        ["=POWER(($A$5 * PI()),2)"],
        ["=POWER(($A$6 * PI()),2)"]
    ]
*/
