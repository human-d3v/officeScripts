function main(wb: ExcelScript.Workbook) {
  var ws:ExcelScript.Worksheet = wb.getActiveWorksheet();
  const lRow:number = (ws.getUsedRange().getLastRow().getRowIndex()+1) //offset 
  let arr: object[] = arrBuilder(ws.getRange(`C2:C${lRow + 1}`));
  const lMonth:number = getMonthFromStr(arr,'max'); //returns the index of the month starting from zero
  const fMonth: number = getMonthFromStr(arr,'min');
  resolveDate(ws); //creates the resolved date column in ws
  createSheets(fMonth,lMonth);

  
  //this is the master function that creates the individual sheets based on the number of months in the range. It then creates a caldendar based on a 31 day month and adds a function to pull the data from column B of the raw data sheet.

  function createSheets(startMonth: number, endMonth: number) {
    for (let i: number = startMonth; i <= endMonth; i++) {
      let name: string = getMonthName(i);
      wb.addWorksheet(name);
      let wksht:ExcelScript.Worksheet = wb.getWorksheet(name);
      wksht.getRange("B2").setValue(name); //allows us to use the Name Value in CalendarDay
      makeCalendar(i,wksht);
    };
  };
    
  function arrBuilder(rng: ExcelScript.Range):object[]{
      const arr:object[] = [];
      const stopRow:number = rng.getLastRow().getRowIndex();
      for(let i:number = 2;i<=stopRow;i++){
          let nRng:ExcelScript.Range = ws.getRange(`C${i}`);
          arr.push(new Date(nRng.getText()));
      };
    return arr;
  };

  function getMonthFromStr(arr:object[],minOrMax:string):number{
    if (minOrMax.toLowerCase() === 'min'){
      let m = new Date(Math.min.apply(null,arr));
      return m.getMonth();
    }
    else if (minOrMax.toLowerCase() === 'max'){
      let m = new Date(Math.max.apply(null,arr));
      return m.getMonth();
    }
    else {
      throw new Error(`${minOrMax} is not accepted syntax. Please enter 'min' or 'max'.`);
    };
  };

  function getMonthName(n:number):string{
    const months:string[] = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    return months[n];
  };

  function getLastDay(n:number):number{
    let arr:number[] = [31,28,31,30,31,30,31,31,30,31,30,31];
    let y:number = new Date().getFullYear();
    if (y % 4 === 0){
      arr[1] = 29;
    };
    return arr[n];
  };

  function weekDayHeader(wksht:ExcelScript.Worksheet){
    let a:string[] = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
    let c:number = 2;
    let n:number = 0;
    for (let s of a){
      wksht.getCell(2,c).setValue(a[n]);
      c += 2;
      n ++;
    };
  }

  function getWeekday(m:number):number{
    return new Date(`${m+1}/01/2023`).getDay();
  };

  /* this function creates the grid(s) for calendar and accepts the month number to use in the getLastDay function.*/
    function makeCalendar(m:number,wksht:ExcelScript.Worksheet){ 
      let rNum:number = 3;
      let cNum:number = (getWeekday(m)*2)+1; //(x*2)+1 -> the offset to apply the correct day to the correct column. 
      const lD:number = getLastDay(m);
      weekDayHeader(wksht); //adds weekday headers
      for (let d:number = 1;d<=lD;d++){
        let cell:ExcelScript.Range = wksht.getCell(rNum,cNum);
        allBorders(cell);
        cell.setValue(d);
        let dRng:ExcelScript.Range = wksht.getRangeByIndexes(rNum,cNum,10,2);
        allBorders(dRng); //creates the cell for the calendar
        d === lD ? lastCalendarDay(rNum,cNum,wksht,m) : calendarDay(rNum,cNum,wksht);
        if((cNum + 2)>14){
          cNum = 1;
          rNum += 10;
        }
        else{
          cNum +=2;
        };
      };
    };


  //accepts a range (one cell or multiple) and places a medium border around it in black. Used in makeCalendar to facilitate readability.
    function allBorders(rng:ExcelScript.Range){
      let format = rng.getFormat();
        format.getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setColor("000000");
        format.getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setWeight(ExcelScript.BorderWeight.medium);
        format.getRangeBorder(ExcelScript.BorderIndex.edgeTop).setColor("000000");
        format.getRangeBorder(ExcelScript.BorderIndex.edgeTop).setWeight(ExcelScript.BorderWeight.medium);
        format.getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setColor("000000");
        format.getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setWeight(ExcelScript.BorderWeight.medium);
        format.getRangeBorder(ExcelScript.BorderIndex.edgeRight).setColor("000000");
        format.getRangeBorder(ExcelScript.BorderIndex.edgeRight).setWeight(ExcelScript.BorderWeight.medium);
    };

  //receives a cell reference from createSheets in the loop, offsets those cell references, and adds a formula that captures the name of the original sheet, and what day of the month it is. 
    function calendarDay(r:number,c:number,wsName:ExcelScript.Worksheet){
      let name: string = ws.getName();
      wsName.getCell((r + 1), (c + 1)).setFormulaR1C1(`=FILTER('${name}'!C2, (('${name}'!C12>=DATEVALUE(R[-1]C[-1]&" "&R2C2&" 2023"))*('${name}'!C12<DATEVALUE((R[-1]C[-1]+1)&" "&R2C2&" 2023"))),"")`);
    };

    function lastCalendarDay(r: number, c: number, wsName: ExcelScript.Worksheet,month:number) {
      let name: string = ws.getName();
      let mName: string = getMonthName(month+1)
      wsName.getCell((r + 1), (c + 1)).setFormulaR1C1(`=FILTER('${name}'!C2, (('${name}'!C12>=DATEVALUE(R[-1]C[-1]&" "&R2C2&" 2023"))*('${name}'!C12<DATEVALUE("1 ${mName} 2023"))),"")`);
    };

    function formArrBuilder(formula:string,num:number):string[][]{
      let arr:string[][]=[];
      for (let i = 0;i<=num;i++){
        arr[i] = [];
        for(let j = 0;j<1;j++){
          arr[i][j] = formula;
        };
      };
      return arr;
    };

    function resolveDate(wksht:ExcelScript.Worksheet){
      let lCol:number = wksht.getUsedRange().getLastColumn().getColumnIndex() +1; //offset to create a new lastColumn
      let rng:ExcelScript.Range = wksht.getRangeByIndexes(1,lCol,lRow-1,1);
      rng.setFormulasR1C1(formArrBuilder(`=R[0]C3 + (NUMBERVALUE(LEFT(RC[-1],1))*IFS(RIGHT(RC[-1],1)="d",1,RIGHT(RC[-1],1)="w",7))`,lRow - 2));
    };
};
