var pickerToFix = 'Godkänd av';
var groupIdToRender = 7;
var currentUser;

function SetAndResolvePeoplePicker(userAccountName) {

    //Här kan vi sätta en user

    var controlName = pickerToFix;

    var peoplePickerDiv = $("[id$='ClientPeoplePicker'][title='" + controlName + "']");

    var peoplePickerEditor = peoplePickerDiv.find("[title='" + controlName + "']");

    var spPeoplePicker = SPClientPeoplePicker.SPClientPeoplePickerDict[peoplePickerDiv[0].id];

    peoplePickerEditor.val(userAccountName);

    spPeoplePicker.AddUnresolvedUserFromEditor(true);

    //disable the field

    spPeoplePicker.SetEnabledState(true);

    //hide the delete/remove use image from the people picker

    //$('.sp-peoplepicker-delImage').css('display','none');

}

function DeletePeoplePickerValue()
{
    //Detta tömmer nuvarande värde i en picker
    var peoplePickerDiv = $("[id$='ClientPeoplePicker'][title='" + pickerToFix + "']");
    $(peoplePickerDiv).find(".sp-peoplepicker-delImage").click();
}


function GetPeoplePickerName() {
//Här får vi vem som är vald just nu
    var controlName = pickerToFix;

    var peoplePickerDiv = $("[id$='ClientPeoplePicker'][title='" + controlName + "']");
    var peoplePickerEditor = peoplePickerDiv.find("[title='" + controlName + "']");

    var spPeoplePicker = SPClientPeoplePicker.SPClientPeoplePickerDict[peoplePickerDiv[0].id];
    var userInfo = spPeoplePicker.GetAllUserInfo(); 
    if(userInfo.length === 1)
        return userInfo[0].DisplayText;
    else
        return null;

}
function GetCurrentUser(userid) {

    
    var requestUri = _spPageContextInfo.webAbsoluteUrl + "/_api/web/getuserbyid(" + userid + ")";

    var requestHeaders = { "accept": "application/json;odata=verbose" };

    $.ajax({
        url: requestUri,
        contentType: "application/json;odata=verbose",
        headers: requestHeaders,
        success: onSuccess,
        error: onError
    });
}

function onSuccess(data, request) {
    console.log(data);
    var loginName = data.d.LoginName.split('|')[2];
    //alert(loginName);
    //SetAndResolvePeoplePicker(loginName);
    // return loginName;
}

function onError(error) {
    alert(error);
}

function RenderSelect(htmlString)
{
    $("#pickerSelect").append(htmlString);
}

function UpdateApprover()
{

document.getElementById("selApprovedBy").disabled = true;

    var pick = $("#bgPicker").find("div[role='textbox']");
$(pick).text("");
$(pick).text(document.getElementById("selApprovedBy").value);
$("#bgPicker").find("img[src='/_layouts/15/images/checknames.png']").click();

var interval = setInterval(function() {
      if ($("#bgPicker").find("div[role='textbox']").text().indexOf(';') > -1)
        {
            
            clearInterval(interval);
            console.log("Picker resolved");
document.getElementById("selApprovedBy").disabled = false;
            
        }
        else
        {
            console.log("Picker NOT resolved");
        }
}, 500);

}





function InitMEPicker(){

    $("#currApprover").hide();
    $(".spMetadataLayoutLabel").hide();


    $.when(

        $.getJSON( _spPageContextInfo.webAbsoluteUrl + "/_api/web/lists(guid'076EBC40-B8C2-49DE-98D6-76D1AC29EF26')/items?$select=Roll,Id,Title,Person/Title,Person/EMail,Person/Id&$expand=Person/Id", 
        function( data ) {

        var isEmpty = true;
        var selectHtml = [];
        selectHtml.push("<select id='selApprovedBy' onchange='UpdateApprover()'>");
        currentUser = $("#currApprover").find("a[class='ms-subtleLink']").text();

        $.each(data.value, function( index, value ) {
            //Här fyller vi upp vår drop down
            if(currentUser == value.Person.Title)
            {
                selectHtml.push("<option selected value='" + value.Person.Title + "'>" + value.Person.Title + "</option>"); 
                isEmpty =false;
            }
             else
                selectHtml.push("<option value='" + value.Person.Title + "'>" + value.Person.Title + "</option>"); 
             
        });

        if(isEmpty)
            selectHtml.push("<option value='empty' selected>Välj Godkännare</option>")

        selectHtml.push("</select>");
        var htmlString = selectHtml.join("");
        RenderSelect(htmlString);

        })

    ).then(function( data, textStatus, jqXHR ) {
    
    
var found = [];
$("#selApprovedBy option").each(function() {
  if($.inArray(this.value, found) != -1) $(this).remove();
  found.push(this.value);
});
    });
}

























