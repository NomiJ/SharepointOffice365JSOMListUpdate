'use strict';

var context = SP.ClientContext.get_current();
var web = context.get_web();
var user = web.get_currentUser();
var notRecappedItems = "";
var listName = 'Donations';
var fieldName = 'WGY_Recapped';

// This code runs when the DOM is ready and creates a context object which is needed to use the SharePoint object model
$(document).ready(function () {
    getUserName();
});

// This function prepares, loads, and then executes a SharePoint query to get the current users information
function getUserName() {
    context.load(user);
    context.executeQueryAsync(onGetUserNameSuccess, onGetUserNameFail);
}


function setMsg(t) {
    $('#message').append("<br />" + t);

}

// This function is executed if the above call is successful
// It replaces the contents of the 'message' element with the user name

function onGetUserNameSuccess() {
    setMsg('Hello ' + user.get_title());
}

function updateRecapped(){
   
    setMsg("Loading List " + listName);
    var tList = web.get_lists().getByTitle(listName);
   
    var query = "<View><Query><Where><Eq><FieldRef Name='"+fieldName+"' /><Value Type='Integer'>0</Value></Eq></Where></Query></View>";
    var camlQuery = new SP.CamlQuery();
    camlQuery.set_viewXml(query);
    notRecappedItems = tList.getItems(camlQuery);

   
    context.load(notRecappedItems);
    context.executeQueryAsync(Function.createDelegate(this, loadDonationsWithNoRecapped), Function.createDelegate(this, queryFailed));
}

function queryFailed(sender, args) {
    setMsg("Webservices call Failed :" + args.get_message());
}
function loadDonationsWithNoRecapped(sender, args) {
    setMsg("Donations has been loaded");

    var itemArray = [];
    var oList = context.get_web().get_lists().getByTitle(listName);
    var it = notRecappedItems.getEnumerator();

    var counter = 0;
    while(it.moveNext())
    {
        counter = counter + 1;
        var oItem = it.get_current();
        var oListItem = oList.getItemById(oItem.get_id());
        oListItem.refreshLoad();
        oListItem.set_item(fieldName, '1');
        oListItem.update();
        itemArray.push(oListItem);
        context.load(itemArray[itemArray.length - 1]);
    }
    if (counter == 0) {

        setMsg('No Item to update');
    } else {
        setMsg('Updating ... ' + counter);

    context.executeQueryAsync(updateMultipleListItemsSuccess, queryFailed);}
}
function updateMultipleListItemsSuccess() 
{    
    setMsg('Items Updated');

}


// This function is executed if the above call fails
function onGetUserNameFail(sender, args) {
    alert('Failed to get user name. Error:' + args.get_message());
}


var collList;
function retrieveAllListProperties() {

    var clientContext = SP.ClientContext.get_current();
    var oWebsite = clientContext.get_web();
    collList = oWebsite.get_lists();

    clientContext.load(collList);

    clientContext.executeQueryAsync(Function.createDelegate(this, onQuerySucceeded), Function.createDelegate(this, onQueryFailed));
}

function onQuerySucceeded() {

    var listInfo = '';

    var listEnumerator = collList.getEnumerator();

    while (listEnumerator.moveNext()) {
        var oList = listEnumerator.get_current();
        listInfo += 'Title: ' + oList.get_title()  + '<br />';
    }
    setMsg(listInfo);
}

function onQueryFailed(sender, args) {
    alert('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
}