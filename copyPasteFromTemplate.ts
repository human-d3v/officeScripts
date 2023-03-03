function main(wb: ExcelScript.Workbook) {
  //define sheets
    const ws:ExcelScript.Worksheet = wb.getWorksheet('ImportTemplate');
    const avgWS:ExcelScript.Worksheet = wb.getWorksheet('MasterAvgs');
    const cntWS:ExcelScript.Worksheet = wb.getWorksheet('MasterCounts');
  //new Date
    let newDate:Date = new Date();
    let mstrDate:Date = new Date(newDate.getFullYear()-1,newDate.getMonth(),1);//offset year value
    cntWS.getRange("C1").setValue(formatDate(mstrDate));
  // //countsheet
    cpShift(cntWS);
    cpVals(cntWS,3);
  // //avgsheet
    cpShift(avgWS);
    cpVals(avgWS,2);
  //cleanup
    cleanup(ws);

    
  function cpVals(destWksht:ExcelScript.Worksheet,col:number){ //accepts a destination worksheet and a column to copy from (expressed as index starting from 0)
    let lRow:number = ws.getUsedRange().getLastRow().getRowIndex();
    let destLRow: number = destWksht.getUsedRange().getLastRow().getRowIndex();
    let destLCol: number = destWksht.getUsedRange().getLastColumn().getColumnIndex();
    for(let i:number = 1;i <= lRow; i++){
      let name:string = ws.getRangeByIndexes(i,0,1,1).getText();
      let val:string|number|boolean = ws.getRangeByIndexes(i,col,1,1).getValue();
      for(let j:number = 1;j<=destLRow;j++){
        if (destWksht.getRangeByIndexes(j,0,1,1).getText()==name){
          destWksht.getRangeByIndexes(j,destLCol,1,1).setValue(val);
          ws.getRangeByIndexes(i,4,1,1).setValue('Found')
          break;
        };
      };
    };
  };

  function cpShift(wksht:ExcelScript.Worksheet){ //shifts values over one column to the left
    let rng: ExcelScript.Range = wksht.getUsedRange();
    let lRow: number = rng.getLastRow().getRowIndex();
    let lCol: number = rng.getLastColumn().getColumnIndex();
    wksht.getRange(`C2:C${lRow + 1}`).setValue(''); //clear existing value
    let cpRng: ExcelScript.Range = wksht.getRangeByIndexes(1, 3, lRow + 1, lCol - 2); //plus 1 to offset last value for row count && minus 2 to both offset and compensate for starting in column 3
    wksht.getRanges('C2').copyFrom(cpRng);
    wksht.getRangeByIndexes(1,lCol,lRow + 1,1).setValue('');
  };

  function formatDate(d:Date):string{
    return (
      [
        pad2d(d.getMonth() + 1),
        pad2d(1),
        d.getFullYear()
      ].join('-')
    );
  };

  function pad2d(n:number):string{ //helper function for formatDate
    return n.toString().padStart(2,'0');
  };

  function cleanup(wksht:ExcelScript.Worksheet){
    wksht.getUsedRange().setValue(''); //clear entire range
    let c:number = 0;
    let titles:string[] = ['fullname','teamname','avgRating','countCases'];
    for (let t of titles){
      wksht.getRangeByIndexes(0,c,1,1).setValue(t);
      c += 1;
    };
    wksht.setVisibility(ExcelScript.SheetVisibility.hidden);
  };
};
