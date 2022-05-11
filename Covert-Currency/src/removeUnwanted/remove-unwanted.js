
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("btnRemove").onclick = removeCharacterUnwanted;
  }
});

async function removeCharacterUnwanted(){
let characterUnwanted = document.getElementById("inputCharacter").value;
let arrayDataSelecteds = await ReturnArrayDataFromCells();
let result= arrayDataSelecteds.map((element,i,arrParent) =>{
  var orgArr=element.map((strCharacter,i,arr) => {
    strCharacter.includes(characterUnwanted)
      ? (strCharacter = recursionString(characterUnwanted, strCharacter))
      : strCharacter;
    arr = [strCharacter];
    return arr;
  },);
  arrParent = [orgArr];
  console.log(arrParent);
  return arrParent;
})
rangeForData(result);
}

function characterWanted(characterUnwanted, strCharacter) {
  strCharacter= strCharacter.replace(characterUnwanted,'');
  strCharacter = recursionString(characterUnwanted, strCharacter)
  return strCharacter;
}
function recursionString(characterUnwanted, strCharacter){
  strCharacter= strCharacter.includes(characterUnwanted) 
  ? strCharacter= characterWanted(characterUnwanted, strCharacter) 
   : strCharacter;
    return strCharacter;  
}
export async function ReturnArrayDataFromCells() {
  try {
    const result= await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRanges();
      range.load("address");
      await context.sync();
    var address = range.address;
    var addressDetail = filterAddress(address);
    let  arrContentFromAddressDetail= await getContentInAddress(addressDetail);
    return arrContentFromAddressDetail;  
    });
    return result;
  } catch (error) {
    console.log(error);
  }
}


// range.address => sheet!address
// filterAdress fuction return address
function filterAddress(address){
 const arrRange = address.split("!");
 return arrRange[1];
}

export async function getContentInAddress(addressDetail) {
    try {
     const result= await Excel.run(async (context) => {
        let sheet = context.workbook.worksheets.getItem("Sheet1");
        let range = sheet.getRange(addressDetail);
        range.load("text");
        await context.sync();
        let result = range.text;
        return result 
      });
        return result;
    } catch (error) {
      console.log(error);
    }
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
      console.log(valuesRange);

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



