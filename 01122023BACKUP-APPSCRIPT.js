//Google sheet
let googleSheet = SpreadsheetApp.getActiveSpreadsheet();

// Sheets
let inventory = googleSheet.getSheetByName("Inventory");
let addDecreaseSheet = googleSheet.getSheetByName("Add Decrease Stock");
let list = googleSheet.getSheetByName("List");
let stockSheet = googleSheet.getSheetByName("Stock");
let supportDataSheet = googleSheet.getSheetByName("Support Data");
let databaseSheet = googleSheet.getSheetByName("Database");
let itemsSheet = googleSheet.getSheetByName("Items");

let ui = SpreadsheetApp.getUi();
let dateNow = new Date();
let hours = dateNow.getHours();
let minutes = dateNow.getMinutes();
let time = `${hours}:${minutes}`;

Logger = BetterLog.useSpreadsheet('1RBAipO7DnPW948SRp7thANQMZ8qV_0XWbaynq9TtmO4'); // to log output

let actionString;
let status;
const item_id_set = 20;

let date = inventory.getRange('C4');
let actionColumn = inventory.getRange('C7');
let item = inventory.getRange('C5');
let qty = inventory.getRange('F5');
let user = inventory.getRange('C6');
let remarks = inventory.getRange('C9');
let statusColumn = inventory.getRange('C8');

const dateStock = addDecreaseSheet.getRange('C4');
const itemStock = addDecreaseSheet.getRange('C5');
const userStock = addDecreaseSheet.getRange('C6');
const remarksStock = addDecreaseSheet.getRange('C9');
const qtyStock = addDecreaseSheet.getRange('F5');

function main() {
  // Get the spreadsheet as an object we can manipulate
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Support Data");
  // const ss = SpreadsheetApp.getActiveSpreadsheet();


  /**
   * Get the data into an array. The .getValues() method returns data in a 2D
   * structure where each row is an array within the main array:
   * 
   * [ ["row1 col1", "row1 col2"], ["row2 col1", "row2 col2"]...]
   */
  const data = ss.getRange(1, 1, ss.getLastRow(), 2).getValues();

  // Open a loop so we can read each row individually
  for(var i=0; i<data.length; i++) {
    // Assign each row to a variable for readability
    let row = data[i]
    
    // If there is no value in cell 1 of each row (the ID cell), then make an ID
    if(!row[0]) {
      let id = generateItemId()

      // You need to use i+1 for the row range because arrays are zero-indexed,
      // but Google Sheets' .setValue() expects row numbers starting with 1.
      ss.getRange(i + 1, 1).setValue(id)
    }
  }
}

// When the script loads, add a menu item to the Spreadsheet so you can run the function
function onOpen() {
  SpreadsheetApp.getUi().createMenu("Make IDs").addItem("Run", "main").addToUi()
}

function randomFixedInteger(length) {
  return Math.floor(
    Math.pow(10, length - 1) +
      Math.random() * (Math.pow(10, length) - Math.pow(10, length - 1) - 1)
  );

}

function generateItemId() {
  let len = 2;
  const item_id = Number(`${item_id_set}${randomFixedInteger(len)}`);
  return item_id;
}

function showAction() {
  if(actionColumn.getValue() == 'Deploy'){
    inventory.hideRow(statusColumn);
  }
}


function clearFields() {
  date.clearContent();
  actionColumn.clearContent();
  item.clearContent();
  qty.clearContent();
  user.clearContent();
  remarks.clearContent();
  statusColumn.clearContent();
}

function clearFieldsStock() {
  dateStock.clearContent();
  itemStock.clearContent();
  userStock.clearContent();
  remarksStock.clearContent();
  qtyStock.clearContent();
  addDecreaseSheet.getRange('F7').clearContent();
  addDecreaseSheet.getRange('C2').clearContent();
  addDecreaseSheet.getRange('F2:G2').clearContent();
}

function getItems() {
  // A1:B24
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
    
    supportDataSheet.getRange(s,1).setValue(array[e][0]);
    supportDataSheet.getRange(s,2).setValue(array[e][1]);
    supportDataSheet.getRange(s,3).setValue(array[e][2]);
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

function deploySave(data, qty) {
  const items = getItems();
  let totalQty;
  
  const result = items.filter((item) => {
    if(item.item_name == data){

      if(item.quantity == 0) {
        ui.alert("No available this item "+ item.items);
        return;
      };

      totalQty = item.quantity - qty;
      item.quantity = totalQty;

      saveListData();
      ui.alert("Deploy "+item.item_name+" to Employee "+user.getValue()+" sucessfully");

      return item
    }
  });

  
  updateData(items);
  
}

function validate() {
  
  if(actionColumn.isBlank() === true){
    ui.alert('Please enter the action');
    return false;
  }else if(item.isBlank() === true){
    ui.alert('Please enter the item');
    return false;
  }else if(qty.isBlank() === true){
    ui.alert('Please enter the quantity');
    return false;
  }else if(user.isBlank() === true){
    ui.alert('Please select the user');
    return false;
  }else if(remarks.isBlank() === true){
    ui.alert('Please enter the remarks');
    return false;
  }
  
  return true;

}

function validateItemStock() {
  if(itemStock.isBlank() === true){
    ui.alert('Please enter the item');
    return false;
  }else if(remarksStock.isBlank() === true){
    ui.alert('Please enter the remarks');
    return false;
  }else if(qtyStock.isBlank() === true){
    ui.alert('Please enter the quantity');
    return false;
  }else if(userStock.isBlank() === true){
    ui.alert('Please select the user');
    return false;
  }

  return true;
}

function deployItem() {
  const confirmation = ui.alert("Submit", 'Do you want to submit the data?', ui.ButtonSet.YES_NO);
  actionString = "Deploy";
  status = "Working";

  if(confirmation == ui.Button.NO) {
    return;
  }

  if(statusColumn.isBlank() === false){
    ui.alert('Status works only when you undeploy item');
    return;
  }

  if(validate() === true){
    if(actionColumn.getValue() === actionString){
      deploySave(item.getValue(), qty.getValue());
    }else {
      ui.alert('Select Deploy in Action');
      return;
    }

    clearFields();
  }
  
}

function undeploySave(data, qty) {
  const items = getItems();
  let totalQty;
  const result = items.filter((item) => {
    if(item.items == data){
      if(item.quantity == 0) {
        ui.alert("No available this item "+ item.items);
        return;
      };

      if(statusColumn.getValue() == 'Defective'){
        totalQty = item.quantity;
        item.quantity = totalQty;
      }else {
        totalQty = item.quantity + qty;
        item.quantity = totalQty;
      }

      saveListData();
      ui.alert("Undeploy "+item.items+" to Employee "+user.getValue()+" sucessfully");

      return item;

    }
  });

  updateData(items);
}


function undeployItem() {
  const confirmation = ui.alert("Submit", 'Do you want to submit the data?', ui.ButtonSet.YES_NO);
  actionString = "Undeploy";

  if(confirmation == ui.Button.NO){
    return;
  }

  if(validate() === true){
    if(statusColumn.isBlank() === true){
      ui.alert('Please enter Status of the item when Undeploy Item');
      return;
    }

    if(actionColumn.getValue() === "Undeploy"){
      undeploySave(item.getValue(), qty.getValue());
    }else {
      ui.alert('Select Undeploy in Action');
      return;
    }

    clearFields();
  }
}


function saveListData() {
  let blankRow = list.getLastRow()+1;
  const datetime = `${date.getValue()} ${time}`;


  list.getRange(blankRow, 1).setValue(datetime);
  list.getRange(blankRow, 2).setValue(actionString);
  statusColumn.isBlank() ? list.getRange(blankRow, 3).setValue('Working') : list.getRange(blankRow, 3).setValue(statusColumn.getValue());
  list.getRange(blankRow, 4).setValue(item.getValue());
  list.getRange(blankRow, 5).setValue(qty.getValue());
  list.getRange(blankRow,6).setValue(user.getValue());
  list.getRange(blankRow,7).setValue(remarks.getValue());
  
  date.isBlank() && list.getRange(blankRow, 1).setValue(dateNow).setNumberFormat("yyyy-mm-dd");

}


function addStock() {
  const confirmation = ui.alert("Submit", 'Do you want to submit the data?', ui.ButtonSet.YES_NO);

  if(confirmation === ui.Button.NO) return;

  if(validateItemStock() === true){
    itemStockAdd(itemStock.getValue(), qtyStock.getValue());
  }

  clearFieldsStock();
}

function decreaseStock() {
  const confirmation = ui.alert("Submit", 'Do you want to submit the data?', ui.ButtonSet.YES_NO);

  if(confirmation === ui.Button.NO) return;

  if(validateItemStock() === true){
    itemStockDecrease(itemStock.getValue(), qtyStock.getValue());
  }

  clearFieldsStock();
}

function itemStockAdd(data, stock) {
  const items = getItems();
  const status = "Add Stock";

  const result = items.filter((item) => {
    if(item.items == data){

      totalQty = item.quantity + stock;
      item.quantity = totalQty;

      saveStockData(status, totalQty);
      ui.alert("Add stock "+itemStock.getValue() +" by "+stock+" ssucessfully");

      return item
    }  
  });


  updateData(items);
}

function itemStockDecrease(data, stock) {
  const items = getItems();
  const status = "Decrease Stock";

  const result = items.filter((item) => {
    if(item.items == data){

      totalQty = item.quantity - stock;
      item.quantity = totalQty;

      saveStockData(status, totalQty);
      ui.alert("Decrease stock "+itemStock.getValue() +" by "+stock+" ssucessfully");

      return item
    }  
  });


  updateData(items);
}

function saveStockData(status, totalStock) {
  let blankRow = stockSheet.getLastRow()+1;
  const datetime = `${dateStock.getValue()} ${time}`;


  stockSheet.getRange(blankRow, 1).setValue(datetime);
  stockSheet.getRange(blankRow, 2).setValue(itemStock.getValue());
  stockSheet.getRange(blankRow, 3).setValue(qtyStock.getValue());
  stockSheet.getRange(blankRow, 4).setValue(userStock.getValue());
  stockSheet.getRange(blankRow,5).setValue(remarksStock.getValue());
  stockSheet.getRange(blankRow,6).setValue(totalStock);
  stockSheet.getRange(blankRow,7).setValue(status);

  
  dateStock.isBlank() && list.getRange(blankRow, 1).setValue(dateNow).setNumberFormat("yyyy-mm-dd");

}

function searchItemId() {
   const totalQtyLabel = addDecreaseSheet.getRange('F7');
   const itemLabel = addDecreaseSheet.getRange('C5');
   const itemIdLabel = addDecreaseSheet.getRange('C2');

   let str = addDecreaseSheet.getRange('F2:G2').getValue();  
   let values = databaseSheet.getDataRange().getValues();

   let valuesFound = false;
    Logger.log(values);
   return;
   for(let i=0; i<values.length; i++){
     let rowValue = values[i];

     if(rowValue[0] == str){
       totalQtyLabel.setValue(rowValue[2]);
       itemLabel.setValue(rowValue[1]);
       itemIdLabel.setValue(rowValue[0]);
       return;
     }else if(rowValue[1] == str){
       totalQtyLabel.setValue(rowValue[2]);
       itemLabel.setValue(rowValue[1]);
       itemIdLabel.setValue(rowValue[0]);
       return;
     }
     
   }

   if(valuesFound == false){
     totalQtyLabel.clearContent();
     itemLabel.clearContent();
     itemIdLabel.clearContent();
     ui.alert('No item found');
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
      ui.alert('Item ID already exists. Generate ID again');
      return;
    }
  }


}

function onlyNumber(str){
  const reg = /^\d+$/;
  return reg.test(str);
}

function createItem() {
  const item_id_label = supportDataSheet.getRange('G4');
  const item_name = supportDataSheet.getRange('G5');
  const item_quantity = supportDataSheet.getRange('G6');

  const confirmation = ui.alert("Submit", 'Do you want to submit the data?', ui.ButtonSet.YES_NO);

  if(confirmation === ui.Button.NO) return;

  if(item_id_label.isBlank() == true){
    ui.alert('Please generate item id');
    return;
  }else if(item_name.isBlank() == true){
    ui.alert('Please enter item name.');
    return;
  }else if(item_quantity.isBlank() == true){
    ui.alert('Please enter quantity');
    return
  }else if(onlyNumber(item_quantity.getValue()) == false){
    ui.alert('Please enter number only');
    return
  }

  saveCreateItemInDatabase(item_id_label.getValue(), item_name.getValue(), item_quantity.getValue());
  ui.alert('Successfully create item '+ item_name.getValue());

}

function clearCreateItemField() {
  const item_id_label = supportDataSheet.getRange('G4');
  const item_name = supportDataSheet.getRange('G5');
  const item_quantity = supportDataSheet.getRange('G6');

  item_id_label.clearContent();
  item_name.clearContent();
  item_quantity.clearContent();
}

function saveCreateItemInDatabase(itemID, itemName, itemQty) {
  let blankRow = itemsSheet.getLastRow()+1;
  let inventoryRow = inventory.getLastRow()+1;

  itemsSheet.getRange(blankRow, 1).setValue(itemID);
  itemsSheet.getRange(blankRow, 2).setValue(itemName);
  itemsSheet.getRange(blankRow, 3).setValue(itemQty);

  inventory.getRange(inventoryRow, 2).setValue(itemID).setBackground('red').setFontColor('white');
  inventory.getRange(inventoryRow, 3).setValue(itemName).setBackground('red').setFontColor('white');
  inventory.getRange(inventoryRow, 4).setValue(itemQty).setBackground('red').setFontColor('white');
  inventory.getRange(inventoryRow, 5).setValue("New").setFontColor('blue');

  const items = getItems();


  const arr = items.map((obj) => {
    return [obj.item_id, obj.item_name, obj.quantity];
  });

  let len = arr.length + 2;

  for(let s = 2, e = 0; s<len, e<len; e++, s++) {
    
    supportDataSheet.getRange(s,1).setValue(arr[e][0]);
    supportDataSheet.getRange(s,2).setValue(arr[e][1]);
    supportDataSheet.getRange(s,3).setValue(arr[e][2]);


  }

}

function findLastRow() {
  let blankRow = inventory.getLastRow();
  Logger.log(blankRow);
}

function test() {
  const items = getAllItem();
  const arr = items.map((obj) => {
    return [obj.item_id, obj.items_name, obj.quantity];
  });

  let bc = arr.length + 2;
  Logger.log(`length: ${arr.length} --- bc: ${bc}`);
  for(let s = 2, e=0; e<bc, s<bc; e++, s++)   {  
    Logger.log(`${s}, 1  - ${arr[e][0]}`);
  }

}


