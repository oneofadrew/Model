# Model Library Google App Script - User Guide

## Description
The `Model` library is used to be able to perform create/replace/update operations on the records of a google sheet. It provides access to the necessary functions to be able to perform these operations as Javascript objects, and facilitates easy interactions for finding specific objects and scanning a list of objects for certain attributes and/or values. It is intended to provide a layer around a google sheet to turn it into a simple database table for scripting to interact with.

#### A note on thread safety...
While there are some simple document locks that ensure multiple users of the spreadsheet don't corrupt the data in place, there is no optimistic locking of records to ensure record integrity of multithreaded writes on the same row. This means that each row should be updated safely in place, but the script has a stale object in memory that is written the stale data will overwrite the existing record.

## Usage

To start using the Model library in your script projects you need to add the Model library to the sript project you would like to use it in. To do this takes a few simple steps:

1. click the + sign next to "Libraries" in the left hand navigation pane on you apps script editor. 
2. Paste the script ID for the model library (`1SS1tXYtDMx7D9d4PPzlGw-FC0-AMrXWakCe5jMwsYDlvq4MiGbJeaxY5`) into the text box and look up the Model library script.
3. Choose the maximum version of the Model library to lock in a specific version.
4. Enter the name of the variable you would like to refer to the library by (or accept the default of `Model`) and then click "Add".

That's it, you are now ready to get going.

### Sheet Configuration
There are several key configurations that are required to use the Model library in your scripting projects.
| Configuration | Type | Description |
| ------------- | ---- | ----------- |
| Data ID | `String` | This is the file ID of the spreadsheet that the sheet is contained within. If no data ID is defined the the current active sheet will be used |
| Sheet Name | `String` | This is the name of the sheet within the spreadsheet that information is stored within |
| Model Keys | `String[]` | These are the fields of the object that is represented by a row on the sheet. The objects will be returned with the fields available as properties on tthe model |
| Column Start | `String` | This is the column reference the model starts at in the sheet. |
| Column End | `String` | This is the column reference the model ends at in the sheet. |
| Has Header | `Boolean` | This is a boolean value that specifies whether or not the first row of the sheet is a header record and can be safely ignored. |
| Key Name | `String` | Where there is a need for a generated key (like a database sequence) the Key Name refers to a Named Range of a single cell that has the current maximum key value in it. If this value is not set then code will assume a natural unique key exists. |
| Enrich Function | `Function` | This is a function with the signature `functionName(model) { return model; }`. It takes a model that has been populated from a row in the sheet and perform post processing of the fields to add into it new fields and helper functions and then returns this enriched version of the model. |
| Rich Text Converters | `Function[]` | This is a list of functions each with the signature `functionName(value) { return RichTextValue; }`. The function has the task of converting the value for a particular field in the model into a rich text value for populating into the spreadsheet. The position of the function in the array needs to correlate to the position in the array of field keys for the key it needs to operate on. If no rich text converter function is supplied (`null`) then a default no-op function is used. |

### Builders
The builder functions are functions that create the simple functions for the basic creeate/replace/update interactivity of the model. Some special conventions that are used within the generated functions:
* The first column of a model is assumed to be the unique key of the model.
* If a Key Name is specified the model will try to generate a new numeric key by looking up the named range called the Key Name and incrementing the numeric value in that cell as the new primary key value for the model to be created.
* Every time a model is created or retrieved the row it has been stored on in the sheet is added to the model with the metadata field name of `row`

#### Build getModels

`buildGetModels(spreadsheetId, sheetName, keys, startCol, endCol, hasHeader, enricher)`

This function creates the `getXXXs()` function that returns all the current XXX models in the sheet.

#### Build findModelByKey

`buildFindModelByKey(spreadsheetId, sheetName, keys, startCol, endCol, enricher)`

This function creates the `findXXXByKey(key)` function that returns the precise XXX model in the sheet that correlates to the key provided. If the key is not found an `Error` will be thrown.

#### Build findModelByRow

`buildFindModelByRow(spreadsheetId, sheetName, keys, startCol, endCol, enricher)`

This function creates the `findXXXByRow(row)` function that returns the precise XXX model in the sheet that is contained on the row provided.

#### Build saveModel

`buildSaveModel(spreadsheetId, sheetName, keys, startCol, endCol, keyName, richTextConverters)`

This function creates the `saveXXX(xxx)` function that saves the XXX object. If the XXX object as denoted by the key field already exists it will overwrite the existing row. If not it will create a new row based on the fields of the object as defined in the keys.

#### Build bulkInsertModels

`buildBulkInsertModels(spreadsheetId, sheetName, keys, startCol, endCol, keyName, richTextConverters)`

This function creates the `bulkInsertXXXs(xxx[])` function that inserts the list of XXX objects. It is intended to be an optimisation on saving the models one at a time when it's understood by the script that all the models are new and can be safely created afresh.

#### Widgets example

As a practical example of how the Model library can be used, imagine we have a sheet that defines the values that we want to store for our Widgets. It looks like this:

Spreadsheet ID: `ABC123`
Sheet Name: `Widget`
Sheet structure:
| ID | Name | Folder Location | Status |
| -- | ---- | --------------- | ------ |
| 1 | Foo | New Widgets | Active |
| 2 | Bar | Old Widgets | Active |
| 3 | Big | New Widgets | Active |
| 4 | Time | New Widgets | Inactive |

In a separate sheet there is a named range called `WidgetID` that refers to a single cell with the value `4` in it.

Our script to interact with this to set up the simple functions would look the following:

```
const DATA_ID = 'ABC123';       //the spreadsheet that holds all the data for our widgets
const WIDGET_SHEET = 'Widget';  //the sheet that holds the widget records
const WIDGET_COL_START = "A";   //widgets start in column A with "ID"
const WIDGET_COL_END = "D";     //widgets end in column D with "Status"
const WIDGET_PK = 'WidgetID';   //the named range that stores the current maximum value of the generated widget primary key
const WIDGET_KEYS = ['id', 'name', 'folder', 'status'];
const WIDGET_RICH_CONVERTERS = [null, null, getFolderLink, null]; //folders need to be stored with their name as a clickable link.

function enrichWidget(widget) {
  widget['active'] = widget['status'] == 'Active';
  return widget;
}

function getFolderLink(value) {
  value = value ? value : "";
  let url = value == "New Widgets" ? 'https://drive.google.com/drive/u/0/folders/XYZ789' : 'https://drive.google.com/drive/u/0/folders/XYZ123';
  return SpreadsheetApp.newRichTextValue().setText(value).setLinkUrl(url).build();
}

let getWidgets = Model.buildGetModels(DATA_ID, WIDGET_SHEET, WIDGET_KEYS, WIDGET_COL_START, WIDGET_COL_END, true, enrichWidget);
let findWidgetByKey = Model.buildFindModelByKey(DATA_ID, WIDGET_SHEET, WIDGET_KEYS, WIDGET_COL_START, WIDGET_COL_END, enrichWidget);
let saveWidget = Model.buildSaveModel(DATA_ID, WIDGET_SHEET, WIDGET_KEYS, WIDGET_COL_START, WIDGET_COL_END, WIDGET_PK, CLIENT_RICH_CONVERTERS);
let buldInsertWidgets = Model.buildBulkInsertModel(DATA_ID, WIDGET_SHEET, WIDGET_KEYS, WIDGET_COL_START, WIDGET_COL_END, WIDGET_PK, CLIENT_RICH_CONVERTERS);
```

With the above set up it would be possible to do the following calls:
```
// Lookup an existing widget
let widget = findWidgetById(1); //gets the widget with the ID of 1 (found on the second row of the sheet)
let msg = widget.active ? 'This widget is still running!' : 'This widget has been disabled.'; //the active boolean comes from our enrich function
console.log(msg); //prints 'This widget is still running!'

//Create a new widget
let newWidget = {'name':'trouble','folder':'Old Folder','Status':'Inactive'}; //the ID will be automatically generated for us
createdWidget = saveWidget(newWidget); //this would add a new row to the widgets sheet with the values for our widget
console.log(`My new widget has been created with ID ${createdWidget.id}`); //prints 'My new widget has been created with ID 5'
```

### Helpers
The Model library also provides some helper functions and features that enable easier navigation and scanning of spreadsheets and objects 

#### Search
The search feature provides an optimised way of scanning a list of objects for a specified set of terms without needing to do multiple iterations of the list. In doing so it enables efficient tailored searching to improve performance for large lists by making search terms composable and reusable. It consists of two main components, the `Search` object and the `runSearch` function.

The `Search` object is created with the `Model.newSearch()` function. It has a function called `where{term, value)`. The `term` parameter specifies the field for the search to run on, and the `value` parameter specifies the value to look for in that search field. The `Search` object also contains the function`and(term, value)`. The `and` function is a simple synonym for `where` and results in more readable code when chaining function calls.

To run a search you provide a `Search` object and a list of models to the `Model.runSearch(search, models)` function. This will filter down the list of models to only those that contain the specified values in the terms as specified in the search. If no models match the search an empty list is returned. Putting these two things together we can get fairly simple scanning of objects for key search terms. For example:

```
function findActiveClientsByLastName(lastName) {
  let clients = getClients();
  let search = Model.newSearch().where('active', true).and('lastName', lastName);
  return Model.runSearch(search, clients);
}
```

#### findLastRow

`Model.findLastRow(spreadsheetId, sheetName, col)`

This is a convenience function that returns the last row of a spreadsheet to have data in it based on a specified column. Note that it will not check all columns, only the last row for the column in question. In actual fact this will return the last row before the first empty cell in a given column, even if there are more records that exist after that empty cell, so it is only safe to use on columns that will always have a value - typically the primary key column for a model.
