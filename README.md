# Model Google App Script Library - User Guide

## Description
The `Model` library is used to be able to perform create/replace/update operations on the records of a google sheet. It provides access to the necessary functions to be able to perform these operations as Javascript objects, and facilitates easy interactions for finding specific objects and scanning a list of objects for certain attributes and/or values. It is intended to provide a layer around a google sheet to turn it into a simple database table for scripting to interact with.

#### A note on thread safety...
While there are some simple document locks that ensure multiple users of the spreadsheet don't corrupt the data in place, there is no optimistic locking of records to ensure record integrity of multithreaded writes on the same row. This means that each row should be updated safely in place, but a script has a stale object in memory will overwrite a newer record when saved.

## Usage

To start using the Model library in your script projects you need to add the Model library to the sript project you would like to use it in. To do this takes a few simple steps:

1. click the + sign next to "Libraries" in the left hand navigation pane on you apps script editor. 
2. Paste the script ID for the model library (`1SS1tXYtDMx7D9d4PPzlGw-FC0-AMrXWakCe5jMwsYDlvq4MiGbJeaxY5`) into the text box and look up the Model library script.
3. Choose the maximum version of the Model library to lock in a specific version.
4. Enter the name of the variable you would like to refer to the library by (or accept the default of `Model`) and then click "Add".

That's it, you are now ready to get going.

### Sheet Configuration
There are several key configurations that are required to use the Model library in your scripting projects.
| Configuration | Type | Inferable | Description |
| ------------- | ---- | --------- | ----------- |
| Data ID | `String` | Yes | This is the file ID of the spreadsheet that the sheet is contained within. If no data ID is defined the the current active sheet will be used |
| Sheet Name | `String` | No | This is the name of the sheet within the spreadsheet that information is stored within |
| Model Keys | `String[]` | Yes | These are the fields of the object that is represented by a row on the sheet. The objects will be returned with the fields available as properties on the model |
| Start Column | `String` | Yes | This is the column reference the model starts at in the sheet. It must be between 'A' and 'ZZ'. |
| Has Header | `Boolean` | Yes | This is a boolean value that specifies whether or not the first row of the sheet is a header record and can be safely ignored. |
| Key Name | `String` | No | Where there is a need for a generated key (like a database sequence) the Key Name refers to a Named Range of a single cell that has the current maximum key value in it. If this value is not set then code will assume a natural unique key exists. |
| Enrich Function | `Function` | No | This is a function with the signature `functionName(model) { return model; }`. It takes a model that has been populated from a row in the sheet and perform post processing of the fields to add into it new fields and helper functions and then returns this enriched version of the model. |
| Rich Text Converters | `Function[]` | No | This is a list of functions each with the signature `functionName(value) { return RichTextValue; }`. The function has the task of converting the value for a particular field in the model into a rich text value for populating into the spreadsheet. The position of the function in the array needs to correlate to the position in the array of field keys for the key it needs to operate on. If no rich text converter function is supplied (`null`) then a default no-op function is used. |

In some cases the above fields can be inferred by convention. In this case they don't need to be provided and the `inferDao` function can be used (see Data Access Objects below).

## Core Concepts

### Model Objects
The Model library seeks to act as an object/relational mapper between a JSON object and a series of columns on a google sheet acting as a database table. It is possible to represent multiple tables on a single spreadsheet, but usually by convention each table would be on separate sheets. Some simple rules that need to be understood:
* The first column of a model is always assumed to be the unique key of the model. When saving a model, it will be looking up the record based on this field as the primary key that determines if this model will be updated or created.
* Key Name is intended to be the Model equivalent of a database sequence. When provided, it should represent a named range for a single cell that contains a numeric value. For data access objects that have the Key Name field provided, when the object being saved doesn't have a value in the primary key (first field) the model will try to generate a new numeric key by looking up the named range and incrementing the numeric value in that cell as the new primary key value for the new model to be created.
* Every time a model is created or retrieved the row it has been stored on in the sheet is added to the model with the metadata field name of `row`. The row can be used as a shorthand reference for simple calculations, but it should not be heavily relied on since it is not refreshed automatically and if someone modifies the data directly (such as adding a new row in the middle of the table the row value may be incorrect until read again.
* In general, sheets that contain data served by the model library should not be modified directly by users. The library does not defend against changes being made to data outside of it's definition, and so strange behaviour might be observed if this happens.

### Conventions
Following the convention of using a header row will reduce the amount of boilerplate code that needs to be written. You are recommended to always use a header row, and to treat the header values like code.
* The spreadsheet will be inferred as the currently active spreadsheet.
* The start column will be inferred as the first column that has a header row value in it. If a "from column" is specified the start column will be the header cell that that first has a value in it starting at the from column and going left.
* The number of fields will be taken from the inferred start column moving left until either an empty cell is reached or there are no more cells available.
* The header row values will be converted to camel case (e.g. camelCase) and used as the field names to the model object

### Data Access Objects
A Data Access Object wraps up a set of functions to allow easily interactivity across a model. It contains a combination of functions that can operate on the sheet, as well as metadata for easy access to the underlying information. A new Data Access Object can either be created or inferred, depending on ability/desire to follow the conventions.

Creating a DAO requires all of the fields to be passed in, but gives complete control over the configuration. For instance, it is possible to not have a header row, or to have header values that do not match the field names for the model object.
`Model.createDao(spreadsheetId, sheetName, keys, startCol, hasHeader, enricher, keyName, richTextConverters)`
Inferring a DAO means that some of the information will be deduced by the library through inspection of the spreadsheet metadata provided. The `fromCol` param only needs to be provided if there is data in columns ahead of the columns that the model object needs to start in which should be skipped. If not provided it defaults to 'A'.
`inferDao(spreadsheetId, sheetName, enricher, keyName, richTextConverters, fromCol = 'A')`

#### Available metadata

* `dao.DATA` - the spreadsheet ID for the model object represented by this Data Access Object
* `dao.SHEET`- the sheet for the model object represented by this Data Access Object
* `dao.KEYS` - the keys (field names) for the model object represented by this Data Access Object
* `dao.START_COL` - the start column reference in the spreadsheet for the model object represented by this Data Access Object
* `dao.KEY_NAME` - the key name (sequence) for the model object represented by this Data Access Object

If any of the fields are not provided tney will be null/undefined.

#### findAll()
This function returns all the current model objects in the sheet.

#### findByKey(key)
This function returns the precise model object in the sheet that correlates to the key provided. If the key is not found an `Error` will be thrown.

#### findByRow(row)
This function returns the precise model object in the sheet that is contained on the spreadsheet row provided.

#### save(model)
This function saves the model object. If it already exists (the primary key is present) it will overwrite the existing row. If not it will create a new row based on the fields of the object, and will generate a new numeric key from the "key name" sequence if needed.

#### bulkInsert([model])
This function inserts a list of model objects into the sheet. It is intended to be an optimisation on saving the model objects one at a time when it's understood by the script that all the objects are new and can be safely created afresh.

#### findLastRow()
This is a convenience function that returns the last populated row for the model objects. Note that it will simply check the first column which should contain the primary key and so always be populated.

#### search(terms)
Run a search across the model objectss represented by this Data Access Object for the terms specified in the search object and return the filtered list of objects.

### Search
The search feature provides an optimised way of scanning a list of objects for a specified set of terms without needing to do multiple iterations of the list. In doing so it enables efficient tailored searching to improve performance for large lists by making search terms composable and reusable. It consists of two main components, the `Search` object and the `runSearch` function.

The `Search` object is created with the `Model.newSearch()` function. It has a function called `where{term, value)`. The `term` parameter specifies the field for the search to run on, and the `value` parameter specifies the value to look for in that search field. The `Search` object also contains the function`and(term, value)`. The `and` function is a simple synonym for `where` and results in more readable code when chaining function calls.

There are two ways to run a search.
1. Provide a `Search` object and a list of models to the `Model.runSearch(terms, models)` function.
2. Call the function `dao.search(terms)` to search all the models represented by that Data Access Object

Both of these will filter down the list of models to only those that contain the specified values in the terms as specified in the search. If no models match the search an empty list is returned. Putting these two things together we can get fairly simple scanning of objects for key search terms. For example:

```
function findActiveHospitalStaff() {
  let doctors = doctorDao.findAll();
  let nurses = nurseDao.findAll();
  let hospitalStaff = doctors.concat(nurses);
  let terms = Model.newSearch().where('active', true);
  return Model.runSearch(terms, clients);
}

//alternatively it's possible to use the convenience method on a DAO
//when we aren't looking across other model objects
function findActiveDoctors() {
  let terms = Model.newSearch().where('active', true);
  return doctorDao.earch(terms);
}
```

### Widgets example

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
const WIDGET_PK = 'WidgetID';   //the named range that stores the current maximum value of the generated widget primary key
const WIDGET_KEYS = ['number', 'longName', 'filesystemLocation', 'status'];
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

//infer the dao from the contents of the sheet
let dao = Model.inferDao(DATA_ID, WIDGET_SHEET, enrichWidget, WIDGET_PK, WIDGET_RICH_CONVERTERS);
```

With the above set up it would be possible to do the following calls:

```
// Lookup an existing widget
let widget = dao.findById(1); //gets the widget with the ID of 1 (found on the second row of the sheet)
let msg = widget.active ? 'This widget is still running!' : 'This widget has been disabled.'; //the active boolean comes from our enrich function
console.log(msg); //prints 'This widget is still running!'

//Create a new widget
let newWidget = {'name':'trouble','folderLocation':'Old Folder','status':'Inactive'}; //the ID will be automatically generated for us
createdWidget = saveWidget(newWidget); //this would add a new row to the widgets sheet with the values for our widget
console.log(`My new widget has been created with ID ${createdWidget.id}`); //prints 'My new widget has been created with ID 5'
```

And here's an example of how that could change if you wanted to use a different configuration for the DAO, even if the data in the sheet doesn't change.

```
//create rather than infer the dao using some extra metadata provided to tailor the experience a bit more
const WIDGET_COL_START = "A";   //widgets start in column A with "ID"
const WIDGET_KEYS = ['number', 'longName', 'filesystemLocation', 'status'];
let customDao = Model.createDao(DATA_ID, WIDGET_SHEET, WIDGET_KEYS, WIDGET_COL_START, true, enrichWidget, WIDGET_PK, WIDGET_RICH_CONVERTERS);

//in this case...
let model = dao.findById(1);
let customModel = customerDao.findById(1);
console.log(model.folderLocation == customModel.filesystemLocation);  //prints true, because they are actually the same row in the same sheet.
```


