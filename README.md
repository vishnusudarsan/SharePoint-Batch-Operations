# SharePoint Batch Operations
It's a jQuery plugin for SharePoint batch requests using JavaScript Object Model (JSOM) in jQuery

## Features Included:
* SharePoint JSOM Batch requests for creation,updation and deletion.
* Retrieve list items using JSOM with <IN> usage in CAML query. 
* jQuery Deffered-Promise function implementation.

### Create List Items:
```
createListItems(<listname>, <array of list items to create>)

var listItems = [{"Title":"ReactJS","Category":"Learning"},
{"Title":"jQuery","Category":"Developing"},
{"Title":"AngularJS","Category":"Experienced"},
];

createListItems("SPList", listItems).then(function(){
alert("Items created.!");
});
```

### Update List Items:
```
batchUpdate(<listname>, <array of list items to update with ID specified>)

var listItems = [
{"ID":32,"Title":"Angular","Category":"Experienced"},
{"ID":33,"Title":"React","Category":"Developing"}
];

batchUpdate("SPList", listItems).then(function(){
 alert("Items updated !");
});
```

### Retrieve List Items:

##### Retrieval using Item Ids & select fields (List Columns) : 
```
retrieveListItems(<listname>,<array of list items to retrieve>,<array of list column names to retrieve>,<customQuery>)

retrieveListItems("SPList",[32,33,34],["ID","Title","Category"]).then(function(response){
response.get_item("ID");
});
```

##### Retrieval using custom query : 

```
var customQuery = "<ViewFields><FieldRef Name='Title'/></ViewFields><Query><Where><Geq><FieldRef Name='ID'/><Value Type='Number'>1</Value></Geq></Where></Query><RowLimit>10</RowLimit>";

retrieveListItems("SPList",[32,33,34],null,customQuery).then(function(response){
response.get_item("ID");
});
```

### Delete List Items:
```
// deleteListItems(<listname>, <array of list items to delete>)

deleteListItems("SPList", [35,36,37]).then(function(){
    alert("Deleted Items !");
});
```
