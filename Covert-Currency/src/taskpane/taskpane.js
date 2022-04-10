// coding standard
// coding convention
// clean coding
// refactoring code legacy code
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */
// const fse = require("fs-extra");
const jsonData = require(".//uit_member.json");
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("getRange").onclick = getRange;
  }
});
export async function getRange() {
  try {
console.log(jsonData);

    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRanges();
      range.load("address");
      await context.sync();
      // console.log(`${range.address}`);
      const address = getAddress(range.address);
      getTextFromRange(address);
      // console.log(typeof arrName);
    });
  } catch (error) {
    console.log(error);
  }
}
export function getAddress(range) {
  const arrRange = range.split("!");
  return arrRange[1];
}

export async function getTextFromRange(address) {
  try {
    await Excel.run(async (context) => {
      let sheet = context.workbook.worksheets.getItem("Sheet1");
      let range = sheet.getRange(address);
      range.load("text");
      await context.sync();
      let result = range.text;
      handleTextName(result);
    });
  } catch (error) {
    console.log(error);
  }
}

const handleTextName = (arrayName) => {
  let range = [];

  arrayName.forEach((element) => seperateFullName(element[0].trim(), range));
};

const seperateFullName = (fullName, range) => {
  let data = [];
  const errorName = ["Họ và tên không hợp lệ", ""];
  let arrLastName = fullName.split(" ");
  if (arrLastName.length >= 3){
     for (let i = 0; i <= arrLastName.length; i++) {
       if (!isValid(arrLastName[i])) {
         range.push(errorName);
         break;
       } else {
         if (i == arrLastName.length - 1) {
           let firstName = arrLastName.splice(-1);
           data.push(arrLastName.join(" "), firstName[0]);
           range.push(data);
           console.log(range);
         }
       }
     }
  }else{
    range.push(errorName);
  }  
  rangeForData(range);
};

function removeAscent(str) {
  if (str === null || str === undefined) return str;
  str = str.toLowerCase();
  str = str.replace(/à|á|ạ|ả|ã|â|ầ|ấ|ậ|ẩ|ẫ|ă|ằ|ắ|ặ|ẳ|ẵ/g, "a");
  str = str.replace(/è|é|ẹ|ẻ|ẽ|ê|ề|ế|ệ|ể|ễ/g, "e");
  str = str.replace(/ì|í|ị|ỉ|ĩ/g, "i");
  str = str.replace(/ò|ó|ọ|ỏ|õ|ô|ồ|ố|ộ|ổ|ỗ|ơ|ờ|ớ|ợ|ở|ỡ/g, "o");
  str = str.replace(/ù|ú|ụ|ủ|ũ|ư|ừ|ứ|ự|ử|ữ/g, "u");
  str = str.replace(/ỳ|ý|ỵ|ỷ|ỹ/g, "y");
  str = str.replace(/đ/g, "d");
  return str;
}
function isValid(string) {
  var re = /^[a-zA-Z!@#\$%\^\&*\)\(+=._-]{2,}$/g; // regex here
  return re.test(removeAscent(string))
}
async function rangeForData(valuesRange) {
  try {
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
  }
}
