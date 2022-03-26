/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    // document.getElementById("getRange").onclick = rangeForData([
    //   ["nguyen mau", "tuan2"],
    //   ["nguyen mau", "tuan"],
    // ]);
    document.getElementById("getRange").onclick = getRange;
  }
});

async function write(){
  const data =  getRange;
  // console.log(typeof data)
  // document.getElementById('test').innerHTML = data;
  console.log(data)
}
export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      // range.format.fill.color = "black";
      await context.sync();
      console.log(`${range.address}`);
    });
  } catch (error) {
    console.error(error);
  }
}


export async function  getRange(){
try{
    await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRanges();
    range.load("address");
    await context.sync();
    // console.log(`${range.address}`);
    const address=   getAddress(range.address);
     getTextFromRange(address);
    // console.log(typeof arrName);
    
  });
} catch(error){
  console.log(error);
}}
export function getAddress(range){
  const arrRange = range.split("!");
  return arrRange[1];
}

export async function getTextFromRange(address) {
  try{
    await Excel.run(async (context) =>{
      let sheet = context.workbook.worksheets.getItem('Sheet1');
      let range = sheet.getRange(address);
      range.load("text");
      await context.sync();
      let result = range.text
      handleTextName(result);
    })
  }catch(error) {
    console.log(error);
  }
}

const handleTextName = (arrayName) =>{
  let range = [];

 arrayName.forEach(element =>seperateFullName(element[0].trim(), range));
   
}

const seperateFullName = (fullName,range) =>{
  let data =[]
  let arrLastName = fullName.split(" ");
  let firstName = arrLastName.splice(-1);
  data.push(arrLastName.join(' '),firstName[0])
  
  range.push(data);
 console.log(range);
  rangeForData(range);
  // range.push([arrLastName.join(" "), firstName[0]]);
  // console.log(arrLastName.join(" "));

}




async function rangeForData(valuesRange) { 
  try {
    // console.log(valuesRange);
    await new Promise((resolve, reject) => {
      Office.context.document.bindings.addFromPromptAsync(
        Office.BindingType.Matrix,
        { id: "currencyRange", promptText: "Select where to display the data" },
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve();
          } else {
            reject();
          }
        }
      );
    });
    await Excel.run(async (context) => {
      let binding = context.workbook.bindings.getItem("currencyRange");
      let range = binding.getRange();
      range.load("address");
      let resizeRange = range.getResizedRange(valuesRange.length - 1, valuesRange[0].length - 1);
      resizeRange.getCell().format.horizontalAlignment = Excel.HorizontalAlignment.center;
      resizeRange.values = valuesRange;
      await context.sync();
    });
  } catch (error) {
    console.log(error);
  }}