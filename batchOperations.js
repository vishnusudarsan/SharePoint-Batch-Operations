/*global SP,alert.console*/
(function() {
  "use strict";
  /*Creating items in batch*/
  var createListItems = function(listName, listItems) {
    var dfd = jQuery.Deferred(),
      itemArray = [],
      i = 0,
      k = 0,
      j = 0,
      columnNames = "",
      clientContext = SP.ClientContext.get_current(),
      oList = clientContext
        .get_web()
        .get_lists()
        .getByTitle("" + listName + ""),
      itemCreateInfo = null,
      oListItem = null;
    for (k = 0; k < listItems.length; k++) {
      itemCreateInfo = new SP.ListItemCreationInformation();
      oListItem = oList.addItem(itemCreateInfo);
      columnNames = Object.keys(listItems[k]);
      for (j = 0; j < columnNames.length; j++) {
        oListItem.set_item(
          "" + columnNames[j] + "",
          "" + listItems[k][columnNames[j]] + ""
        );
      }
      oListItem.update();
      itemArray[i] = oListItem;
      clientContext.load(itemArray[i]);
      i += 1;
    }
    clientContext.executeQueryAsync(
      function() {
        dfd.resolve();
      },
      function(sender, args) {
        console.log(
          "Request failed. " + args.get_message() + "\n" + args.get_stackTrace()
        );
        dfd.reject();
      }
    );
    return dfd.promise();
  };
  /*Delete List Items batch*/
  var deleteListItems = function(listName, deleteitems) {
    var dfd = jQuery.Deferred(),
      clientContext = SP.ClientContext.get_current(),
      oList = clientContext
        .get_web()
        .get_lists()
        .getByTitle("" + listName + ""),
      deleteItemsLength = deleteitems.length,
      oListItem = null;
    for (var i = 0; i < deleteItemsLength; i += 1) {
      oListItem = oList.getItemById(deleteitems[i]);
      oListItem.deleteObject();
    }
    clientContext.executeQueryAsync(
      function() {
        dfd.resolve();
      },
      function(sender, args) {
        console.log(
          "Request failed. " + args.get_message() + "\n" + args.get_stackTrace()
        );
        dfd.reject();
      }
    );
    return dfd.promise();
  };
  /*update product list item*/
  var batchUpdate = function(listName, listItems) {
    var dfd = jQuery.Deferred(),
      updatingID = [],
      k = 0,
      oListItem = null,
      clientContext = new SP.ClientContext.get_current(),
      oList = clientContext
        .get_web()
        .get_lists()
        .getByTitle(listName);
    jQuery.each(listItems, function(key, value) {
      jQuery.each(value, function(keys, val) {
        oListItem = oList.getItemById(value.ID);
        if (keys.toLowerCase() !== "id") {
          oListItem.set_item(keys, val);
          oListItem.update();
          updatingID[k] = oListItem;
          clientContext.load(updatingID[k]);
        }
      });
      k += 1;
    });
    clientContext.executeQueryAsync(
      function() {
        dfd.resolve();
      },
      function(sender, args) {
        console.log(
          "Request failed. " + args.get_message() + "\n" + args.get_stackTrace()
        );
        dfd.reject();
      }
    );
    return dfd.promise();
  };
  /* Retrieve list items in batch */
  var retrieveListItems = function(
    listName,
    idList,
    selectFields,
    customQuery
  ) {
    var dfd = jQuery.Deferred(),
      qureyValues = "",
      queryLogic = "",
      selectQuery = "",
      combinedQuery = "",
      i = 0,
      j = 0;
    if (idList && idList.length > 0) {
      switch (idList.length) {
        case 1:
          queryLogic =
            "<Query><Where><Eq><FieldRef Name='ID'/><Value Type='Counter'>" +
            idList[0] +
            "</Value>" +
            "</Eq></Where></Query>";
          break;
        default:
          var idListLength = idList.length;
          for (i = 0; i < idListLength; i += 1) {
            if (i !== idListLength - 1) {
              qureyValues += "<Value Type='Counter'>" + idList[i] + "</Value>";
            } else {
              qureyValues +=
                "</Values></In><In><FieldRef Name='ID'/><Values>" +
                "<Value Type='Counter'>" +
                idList[i] +
                "</Value></Values></In>";
            }
          }
          queryLogic =
            "<Query><Where><Or><In><FieldRef Name='ID'/><Values>" +
            qureyValues +
            "</Or></Where></Query>";
          break;
      }
      if (selectFields && selectFields.length > 0) {
        selectQuery += "<ViewFields>";
        for (j = 0; j < selectFields.length; j++) {
          selectQuery += "<FieldRef Name='" + selectFields[j] + "'/>";
        }
        selectQuery += "</ViewFields>";
      }
    }
    if (customQuery == null || customQuery == "") {
      combinedQuery = selectQuery + queryLogic;
    } else {
      combinedQuery = customQuery;
    }
    var clientContext = new SP.ClientContext.get_current(),
      oList = clientContext
        .get_web()
        .get_lists()
        .getByTitle(listName),
      camlQuery = new SP.CamlQuery();
    camlQuery.set_viewXml("<View>" + combinedQuery + "</View>");
    var collListItem = oList.getItems(camlQuery);
    clientContext.load(collListItem);
    clientContext.executeQueryAsync(
      function(sender, args) {
        var listItemEnumerator = collListItem.getEnumerator();
        while (listItemEnumerator.moveNext()) {
          var oListItem = listItemEnumerator.get_current();
          dfd.resolve(oListItem);
        }
      },
      function(sender, args) {
        console.log(
          "Request failed. " + args.get_message() + "\n" + args.get_stackTrace()
        );
        dfd.reject();
      }
    );
    return dfd.promise();
  };
});
