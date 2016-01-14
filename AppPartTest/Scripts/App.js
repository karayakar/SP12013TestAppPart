'use strict';

var context = SP.ClientContext.get_current();
var user = context.get_web().get_currentUser();
// Alternative method to get Url Parameter
// var hostUrl = GetUrlKeyValue('SPHostUrl');
var hostUrl = decodeURIComponent(getQueryStringParameter("SPHostUrl"));
var appWebUrl = decodeURIComponent(getQueryStringParameter("SPAppWebUrl"));
var scriptbase = hostUrl + "/_layouts/15/";

// This code runs when the DOM is ready and creates a context object which is needed to use the SharePoint object model
$(document).ready(function () {
    getUserName();
    // Load the js files and continue to the successHandler
    $.getScript(scriptbase + "SP.RequestExecutor.js", getAllUserProfiles);
});

// This function prepares, loads, and then executes a SharePoint query to get the current users information
function getUserName() {
    context.load(user);
    context.executeQueryAsync(onGetUserNameSuccess, onGetUserNameFail);
}

// This function is executed if the above call is successful
// It replaces the contents of the 'message' element with the user name
function onGetUserNameSuccess() {
    $('#message').text('Hello ' + user.get_title());
}

// This function is executed if the above call fails
function onGetUserNameFail(sender, args) {
    alert('Failed to get user name. Error:' + args.get_message());
}

function getAllUserProfiles() {
    // executor: The RequestExecutor object
    // Initialize the RequestExecutor with the add-in web URL.
    var executor = new SP.RequestExecutor(appWebUrl);
    var fullUrl = appWebUrl + "/_api/SP.AppContextSite(@target)/web/lists?$orderby=Title&$select=Title&@target='" + encodeURIComponent(hostUrl) + "'";

    $('#pageContext').text(fullUrl);

    // Issue the call against the add-in web.
    executor.executeAsync(
        {
            url: fullUrl,
            type: "GET",
            headers: {
                "accept": "application/json; odata=verbose",
                "content-type": "application/json; odata=verbose"
            },
            success: onGetAllUserProfilesSucceeded,
            error: onGetAllUserProfilesFailed
        }
    );
}

function onGetAllUserProfilesSucceeded(data) {
    $('#data').empty()
    var jsonObject = JSON.parse(data.body);
    var results = jsonObject.d.results;
    for (var i = 0; i < results.length; i++) {
        $("<p>" + i + " " + results[i].Title + "</p>").appendTo('#data')
    }

}

function onGetAllUserProfilesFailed(data) {
    $('#data').text('Failed to load data');
}

function getQueryStringParameter(urlParameterKey) {
    var params = document.URL.split('?')[1].split('&');
    var strParams = '';
    for (var i = 0; i < params.length; i = i + 1) {
        var singleParam = params[i].split('=');
        if (singleParam[0] == urlParameterKey)
            return decodeURIComponent(singleParam[1]);
    }
}
