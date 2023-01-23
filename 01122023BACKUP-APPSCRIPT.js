//Google sheet
let googleSheet = SpreadsheetApp.getActiveSpreadsheet();

// Sheets
let userFormSheet = googleSheet.getSheetByName("User Form");
let deploymentRecordSheet = googleSheet.getSheetByName("Deployment Records");
let supportDataSheet = googleSheet.getSheetByName("Support Data");
let itemsSheet = googleSheet.getSheetByName("Items");

let testSheet = googleSheet.getSheetByName("test");

;
let dateNow = new Date();
let hours = ("0" + dateNow.getHours()).slice(-2);
let minutes = ("0" + dateNow.getMinutes()).slice(-2);
let time = `${hours}:${minutes}`;

Logger = BetterLog.useSpreadsheet('1OK2oUMaPSDtHPWQzRO0cPnsMAs9T6awLcMSuHGOawpk'); // to log output

let actionString;
let status;


const date = userFormSheet.getRange('C6');
const item = userFormSheet.getRange('C8');
const user = userFormSheet.getRange('C10');
const deployment = userFormSheet.getRange('C12');
const item_status = userFormSheet.getRange('C14');
const remarks = userFormSheet.getRange('C16');
const qty = userFormSheet.getRange('F9:F10');

const testt = userFormSheet.getRange('L44');


const datetime = `${date.getValue()} ${time}`;

const dateStock = userFormSheet.getRange('C29');
const itemStock = userFormSheet.getRange('C31');
const userStock = userFormSheet.getRange('C33');
const actionStock = userFormSheet.getRange('C35');
const remarksStock = userFormSheet.getRange('C37');
const qtyStock = userFormSheet.getRange('F32:F33');

// Label
const item_id_label =  userFormSheet.getRange('C27');
const item_name_label = userFormSheet.getRange('C31');
const item_total_label = userFormSheet.getRange('F35:G37');

const errorText = userFormSheet.getRange('E26');
const remarksError = userFormSheet.getRange('C17');
const itemStatusError = userFormSheet.getRange('C15');


const searchField = userFormSheet.getRange('E27:F27');



function test() {
}


function validate() {
  if(date.isBlank() === true){
    SpreadsheetApp.getUi().alert('Please enter the date');
    errorBorder(date, "red");
    return false;
  }else if(item.isBlank() === true){
    SpreadsheetApp.getUi().alert('Please enter the item');
    errorBorder(item, "red");
    return false;
  }else if(user.isBlank() === true){
    SpreadsheetApp.getUi().alert('Please select the user');
    errorBorder(user, "red");
    return false;
  }else if(deployment.isBlank() === true){
    SpreadsheetApp.getUi().alert('Please enter the Deployment');
    errorBorder(deployment, "red");
    return false;
  }else if(item_status.isBlank() === true){
    SpreadsheetApp.getUi().alert('Please select the item status');
    errorBorder(item_status, "red");
    return false;
  }else if(remarks.isBlank() === true){
    SpreadsheetApp.getUi().alert('Please enter the remarks');
    errorBorder(remarks, "red");
    return false;
  }else if(qty.isBlank() === true){
    SpreadsheetApp.getUi().alert('Please enter quantity');
    errorBorder(qty, "red");
    return false;
  }

  return true;

}

function validateItemStock() {
  if(searchField.isBlank() === false){
    if(item_id_label.getValue() === ""){
      SpreadsheetApp.getUi().alert('Click the Search Button');
      return false;
    }else{
      if(dateStock.isBlank() === true){
        SpreadsheetApp.getUi().alert('Please enter the Date');
        errorBorder(dateStock, "red");
        return false;
      }else if(userStock.isBlank() === true){
        SpreadsheetApp.getUi().alert('Please select the user');
        errorBorder(userStock, "red");
        return false;
      }else if(actionStock.isBlank() === true){
        SpreadsheetApp.getUi().alert('Please select the action');
        errorBorder(actionStock, "red");
        return false;
      }else if(remarksStock.isBlank() === true){
        SpreadsheetApp.getUi().alert('Please enter the remarks');
        errorBorder(remarksStock, "red");
        return false;
      }else if(qtyStock.isBlank() === true){
        SpreadsheetApp.getUi().alert('Please enter the quantity');
        errorBorder(qtyStock, "red");
        return false;
      }
    }
  }else {
    SpreadsheetApp.getUi().alert('Please enter Item ID or Item Name');
    errorBorder(searchField);
    return;
  }

  return true;
}

function errorBorder(rangeField, color) {
  return rangeField.setBorder(true, true, true, true, true, true, color, SpreadsheetApp.BorderStyle.SOLID);

}

function deploySubmit() {
  if(validate() === true){
    deployClearBorder();

    remarksError.setFontColor('transparent');
    itemStatusError.setFontColor('transparent');

      if(deployment.getValue().toUpperCase() === 'DEPLOY'){
          deployItem(item.getValue(), qty.getValue());
          return;
      }else if(deployment.getValue().toUpperCase() === 'REPLACEMENT') {
        replacementItem(item.getValue(), qty.getValue());
        return;
      }else if(deployment.getValue().toUpperCase() === 'DISPATCH'){
        dispatchItem(item.getValue(), qty.getValue());
        return;
      }

  }


}

function deployItem(itemName, qtyNo) {
  const messageLog = `Deploy ${itemName} successfully`;
  const items = getItems();
  let totalQty;

  const result = items.filter((item) => {
    if(item.item_name.toLowerCase() == itemName.toLowerCase()){

      if(item_status.getValue().toUpperCase() == 'DEFECTIVE'){
        SpreadsheetApp.getUi().alert('Check the item status');
        itemStatusError.setValue('Check the item status').setFontColor('red');
        return;
      }

      if(item.quantity == 0) {
        SpreadsheetApp.getUi().alert("No available this item "+ itemName);
        return;
      }else if(qtyNo > item.quantity){
        SpreadsheetApp.getUi().alert("Error, Try again");
        return;
      }

      totalQty = item.quantity - qtyNo;
      item.quantity = totalQty;
        
      saveListData();
      SpreadsheetApp.getUi().alert(messageLog);
      Logger.log(messageLog);
      clearFields();

      // SpreadsheetApp.getUi().alert(`Deploy ${item.getValue()} successfully`);
      return item
    }
  });

  updateData(items);
}


function replacementItem(itemName, qtyNo) {
  const messageDefectiveLog = `Replacement item ${itemName} status is ${item_status.getValue()} was not added to the stock`;
  const messageWorkingLog = `Replacement item ${itemName} status is ${item_status.getValue()} was added to the stock`;

  const items = getItems();
  let totalQty;
  const result = items.filter((item) => {
    if(item.item_name.toLowerCase() == itemName.toLowerCase()){

      if(item.quantity == 0) {
        SpreadsheetApp.getUi().alert("No available this item "+ itemName);
        return;
      };

      if(item_status.getValue().toUpperCase() == 'DEFECTIVE'){
        totalQty = item.quantity;
        item.quantity = totalQty;
        SpreadsheetApp.getUi().alert(messageDefectiveLog);
        Logger.log(messageDefectiveLog);
      }else {
        totalQty = item.quantity + qtyNo;
        item.quantity = totalQty;
        SpreadsheetApp.getUi().alert(messageWorkingLog);
        Logger.log(messageWorkingLog);
      }
      saveListData();
      SpreadsheetApp.getUi().alert(`Replacement ${itemName} successfully`);
      clearFields();
      return item;
    }
  });


  updateData(items);
}

function dispatchItem(itemName, qtyNo) {
  const message = `Dispatch item  ${itemName} decrease by ${qtyNo}`;
  const items = getItems();
  let totalQty;

  const result = items.filter((item) => {
    if(item.item_name.toLowerCase() == itemName.toLowerCase()) {



     if(item.quantity == 0){
       SpreadsheetApp.getUi().alert('No available this item');
       return;
     } 

      if(item_status.getValue().toUpperCase() == 'DEFECTIVE'){
        totalQty = item.quantity - qtyNo;
        item.quantity = totalQty;
        SpreadsheetApp.getUi().alert(message);
        Logger.log(message);
      }else {
        SpreadsheetApp.getUi().alert('Item status must be Defective');
        itemStatusError.setValue('Check the item status').setFontColor('red');
        return;      
      }

      saveListData();
      SpreadsheetApp.getUi().alert(`Dispatch ${itemName} successfully`);
      clearFields();
      return item;
    }
  });

  updateData(items);
}

function stockSubmit() {
  // const confirmation = ui.alert("Submit", "Do you want to submit the data?", ui.ButtonSet.YES_NO);

  // if(confirmation == ui.Button.NO){
  //   return;
  // }

  if(validateItemStock() === true){
      stockClearBorder();
      errorText.setFontColor('transparent');
    

    if(actionStock.getValue() === 'ADD'){
      itemStockAdd(itemStock.getValue(), qtyStock.getValue());
      return;
    }else {
      itemStockDecrease(itemStock.getValue(), qtyStock.getValue());
      return;
    }

  }
}




function stockClearBorder() {
  errorBorder(dateStock, "black");
  errorBorder(searchField, "black");
  errorBorder(userStock, "black");
  errorBorder(actionStock, "black");
  errorBorder(remarksStock, "black");
  errorBorder(qtyStock, "black");
}

function deployClearBorder() {
  errorBorder(date, "black");
  errorBorder(user, "black");
  errorBorder(item, "black");
  errorBorder(deployment, "black");
  errorBorder(item_status, "black");
  errorBorder(remarks, "black");
  errorBorder(qty, "black");
}

function clearFields() {
  deployClearBorder();

  remarksError.setFontColor('transparent');
  itemStatusError.setFontColor('transparent');
 
  date.clearContent();
  user.clearContent();
  item.clearContent();
  deployment.clearContent();
  item_status.clearContent();
  remarks.clearContent();
  qty.clearContent();
}

function clearFieldsStock() {
  stockClearBorder();

  errorText.setFontColor('transparent');

  dateStock.clearContent();
  itemStock.clearContent();
  userStock.clearContent();
  remarksStock.clearContent();
  qtyStock.clearContent();
  item_id_label.clearContent();
  item_name_label.clearContent();
  item_total_label.clearContent();
  searchField.clearContent();
  actionStock.clearContent();
}

function getItems() {
  let data = itemsSheet.getDataRange().getValues();
  let arrayOfobj = convertToArrayOfObjects(data);

  return arrayOfobj;
}

function updateData(data) {
  const array = data.map((obj) => {
    return [obj.item_id, obj.item_name, obj.quantity];
  });
  let len = array.length + 2;

  for(let i = 2, e = 0, s = 2; i < len, e < len, s < len; i++, e++, s++){
    itemsSheet.getRange(i, 1).setValue(array[e][0])
    itemsSheet.getRange(i, 2).setValue(array[e][1]);
    itemsSheet.getRange(i, 3).setValue(array[e][2]);
  }
}


function convertToArrayOfObjects(data) {
    var keys = data.shift(),
        i = 0, k = 0,
        obj = null,
        output = [];
    for (i = 0; i < data.length; i++) {
        obj = {};
        for (k = 0; k < keys.length; k++) {
            obj[keys[k]] = data[i][k];
        }
        output.push(obj);
    }
    return output;
}

function saveListData() {

  let blankRow = deploymentRecordSheet.getLastRow()+1;
  const datetime = `${date.getValue()} ${time}`;

  deploymentRecordSheet.getRange(blankRow, 1).setValue(datetime);
  deploymentRecordSheet.getRange(blankRow, 2).setValue(item.getValue());
  deploymentRecordSheet.getRange(blankRow,3).setValue(user.getValue());
  deploymentRecordSheet.getRange(blankRow, 4).setValue(deployment.getValue());
  deploymentRecordSheet.getRange(blankRow, 5).setValue(item_status.getValue());
  deploymentRecordSheet.getRange(blankRow, 6).setValue(qty.getValue());
  deploymentRecordSheet.getRange(blankRow,7).setValue(remarks.getValue());


  // statusColumn.isBlank() ? list.getRange(blankRow, 3).setValue('Working') : deploymentRecordSheet.getRange(blankRow, 3).setValue(statusColumn.getValue());
  
  // date.isBlank() && deploymentRecordSheet.getRange(blankRow, 1).setValue(datetime);

}


function decreaseStock() {
  const confirmation = SpreadsheetApp.getUi().alert("Submit", 'Do you want to submit the data?', SpreadsheetApp.getUi().ButtonSet.YES_NO);

  if(searchField.isBlank() === true){
    SpreadsheetApp.getUi().alert('Please enter Item name or Item ID');
    return;
  }
  if(confirmation === SpreadsheetApp.getUi().Button.NO) return;

  if(validateItemStock() === true){
    itemStockDecrease(itemStock.getValue(), qtyStock.getValue());
  }

}

function itemStockAdd(itemName, stock) {
  const items = getItems();
  const itemStatus = "Add Stock";

  const result = items.filter((item) => {
    if(item.item_name == itemName){

      totalQty = item.quantity + stock;
      item.quantity = totalQty;
      
      saveStockData(itemStatus, totalQty);
      SpreadsheetApp.getUi().alert("Add stock "+itemName +" by "+stock+" successfully");
      clearFieldsStock();
      return item
    }  
  });


  updateData(items);
}

function itemStockDecrease(itemName, stock) {
  const items = getItems();
  const itemStatus = "Decrease Stock";

  const result = items.filter((item) => {
    if(item.item_name == itemName){

      totalQty = item.quantity - stock;

      // for negative
      if(totalQty < 0){
        SpreadsheetApp.getUi().alert("Error. Please try again");
        return;
      }
      Logger.log(totalQty);

      item.quantity = totalQty;

      saveStockData(itemStatus, totalQty);
      SpreadsheetApp.getUi().alert("Decrease stock "+itemStock.getValue() +" by "+stock+" successfully");
      clearFieldsStock();

      return item
    }  
  });


  updateData(items);
}

function saveStockData(itemStatus, totalStock) {
  let blankRow = deploymentRecordSheet.getLastRow()+1;
  const datetime = `${dateStock.getValue()} ${time}`;
  const remarksString = `${remarksStock.getValue()}   Total: ${totalStock}`;

  deploymentRecordSheet.getRange(blankRow, 1).setValue(datetime);
  deploymentRecordSheet.getRange(blankRow, 2).setValue(itemStock.getValue());
  deploymentRecordSheet.getRange(blankRow, 3).setValue(userStock.getValue());
  deploymentRecordSheet.getRange(blankRow, 4).setValue(actionStock.getValue());
  deploymentRecordSheet.getRange(blankRow,5).setValue(itemStatus);
  deploymentRecordSheet.getRange(blankRow, 6).setValue(qtyStock.getValue());
  deploymentRecordSheet.getRange(blankRow, 7).setValue(remarksString);

  dateStock.isBlank() && deploymentRecordSheet.getRange(blankRow, 1).setValue(`${dateNow} ${time}`).setNumberFormat("yyyy-mm-dd");

}

function searchItemId() {
   let str = searchField.getValue().toString();  
   let values = itemsSheet.getDataRange().getValues();

   let valuesFound = false;

   for(let i=0; i<values.length; i++){
     let rowValue = values[i];
     const item_id_column = rowValue[0];
     const item_name_column = rowValue[1];

     if(item_id_column.toString().toLowerCase() == str.toLowerCase()){
       errorText.setFontColor('transparent');
       errorBorder(searchField, "black");

       item_total_label.setValue(rowValue[2]);
       item_name_label.setValue(rowValue[1]);
       item_id_label.setValue(rowValue[0]);
       return;
     }else if(item_name_column.toString().toLowerCase() == str.toLowerCase()){
       errorText.setFontColor('transparent');
      errorBorder(searchField, "black");

       item_total_label.setValue(rowValue[2]);
       item_name_label.setValue(rowValue[1]);
       item_id_label.setValue(rowValue[0]);
       return;
     }
     
   }

   if(valuesFound == false){
    
     errorText.setValue(`Item not found`).setFontColor('red');
     errorBorder(searchField, "red");

     item_total_label.clearContent();
     item_name_label.clearContent();
     item_id_label.clearContent();
     SpreadsheetApp.getUi().alert('No item found');
   }
}

function makeIDButton() {
  const item_id = generateItemId();
  const item_id_label = supportDataSheet.getRange('G4');
  const values = itemsSheet.getDataRange().getValues();

  item_id_label.setValue(item_id);

  for(let i=0; i<values.length; i++){
    let rowValue = values[i];

    if(rowValue[0] == item_id_label.getValue()){
      SpreadsheetApp.getUi().alert('Item ID already exists. Generate ID again');
      return;
    }
  }
}

function onlyNumber(str){
  const reg = /^\d+$/;
  return reg.test(str);
}
