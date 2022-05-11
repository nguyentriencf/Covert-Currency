
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("btnRemove").onclick = removeCharacterUnwanted;
  }
});

async function removeCharacterUnwanted(){
let characterUnwanted = document.getElementById("inputCharacter").value;
let arrayDataSelecteds = await ReturnArrayDataFromCells();
console.log(arrayDataSelecteds[0]);
arrayDataSelecteds.map(arrays =>{
    arrays.filter(strCharacter =>{
            strCharacter.includes(characterUnwanted)
              ? console.log(`${strCharacter}đúng`)
              : console.log(`${strCharacter} Sai`);
    });
})
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



