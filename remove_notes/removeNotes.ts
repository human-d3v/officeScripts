/*
  Notes live on an abstraction layer on excel OTW. They are rendered differently than comments.
  so the getAllComments() method won't capture notes. The only solution that I've found for this is to 
  retain the values while removing the notes on a specified range. 
*/

function main(wb:ExcelScript.Workbook){
    const ws:ExcelScript.Worksheet = wb.getActiveWorksheet();
    const rng:ExcelScript.Range = ws.getUsedRange(); 
    /*
        alternatively, you can feed it a specific range like 
        ws.getRange("A1:Z99")
    */
    const vals = rng.getValues();
    rng.clear(ExcelScript.ClearApplyTo.all) //<-removes formatting, hyperlinks, notes, etc.
    
    rng.setValues(vals); //resets the values
}
