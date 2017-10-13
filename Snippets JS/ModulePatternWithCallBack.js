var PropertyBagUtility = (function () {

    var requestHeaders = {
        "accept": "application/json;odata=verbose",
        "content-type": "application/json;odata=verbose",
    };

    _getCustomSearchUrl = function (callback) {
        $.ajax({
            url: _spPageContextInfo.siteAbsoluteUrl + "/_api/web/allProperties",
            method: "GET",
            headers: requestHeaders
        }).done(function (data) {
            console.log(data);
            var searchResultPageObject = JSON.parse(data.d.SRCH_x005f_SB_x005f_SET_x005f_WEB);
            callback(searchResultPageObject.ResultsPageAddress);
        });
    };



    return {
        getCustomSearchUrl: function (callback)
        {
            return _getCustomSearchUrl(callback);
        }
    };

})();

PropertyBagUtility.getCustomSearchUrl(function(ove){alert(ove);});



var AutoSuggestUtility = (function () {

    var requestHeaders = {
        "accept": "application/json;odata=verbose",
        "content-type": "application/json;odata=verbose",
    };

    _getCustomSearchUrl = function (callback) {
        $.ajax({
            url: _spPageContextInfo.siteAbsoluteUrl + "/_api/web/allProperties",
            method: "GET",
            headers: requestHeaders
        }).done(function (data) {
            console.log(data);
            var searchResultPageObject = JSON.parse(data.d.SRCH_x005f_SB_x005f_SET_x005f_WEB);
            callback(searchResultPageObject.ResultsPageAddress);
        });
    };



    return {
        getCustomSearchUrl: function (callback)
        {
            return _getCustomSearchUrl(callback);
        }
    };

})();
