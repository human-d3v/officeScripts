/*

In Office typescript library, the range.setFormulasR1C1() method has laughably little documentation. The method requires an array of arrays, each sub-array with it's own copy 
of the formula to be used. Or, barring that, each sub-array requires a separate function. 

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
let sheetName = sheet.getName()
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

