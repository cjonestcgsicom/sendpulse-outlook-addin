$(document).ready(() => {

    $('#run').hide();
    $('#run').click(run);

    var Auth2AccountData = Auth2AccountData || {};
    var dialog;
    var contacts = [];
    var folders = [];
    var baseURL = localStorage.getItem("baseURL");

    Auth2AccountData.secret = 'nnoHKD75!)gwzrSICW728?[';
    Auth2AccountData.clientID = 'c80d93b6-a45b-405c-add9-388a25beadb9';

// The Auth0 subdomain and client ID need to be shared with the popup dialog
    localStorage.setItem('clientID', Auth2AccountData.clientID);
    localStorage.setItem('secret', Auth2AccountData.secret);
    localStorage.setItem('redirectURI', baseURL + 'function-file/function-file.html');
    localStorage.removeItem('selectedMicrosoftContacts');


// The initialize function must be run each time a new page is loaded
    Office.initialize = (reason) => {
        $('#sideload-msg').hide();
        $('.alert').hide();
        $('#progress_dynamics').hide();
        $('#sync_div').hide();
        $('#spinner').hide();
        $('#app-body').show();
        $('#run').hide();
        $('.close').click((e)=>{
            $('.alert').hide();
        });
        $('[data-dismiss]').click((e)=>{
            $('.alert').hide();
        });
        onStateChange();
    };

    //auth State's been changed so we just check if it is appropriate to get contacts and synch them
    function onStateChange()
    {
        if ( localStorage.getItem("microsoftToken") === null || typeof localStorage.getItem("microsoftToken") === 'undefined' ) {
            console.log("Still not authorized with Microsoft account");
            //$('#syncBtnTitle').html('Select Outlook Contacts Folder');

        }
        else
        {
            console.log('Token we got: ' + localStorage.getItem("microsoftToken"));
            getFolders(localStorage.getItem("microsoftToken"));
        }

    };

    // Use the Office dialog API to open a pop-up and display the sign-in page for choosing an identity provider.
    function showLoginPopup() {

        // Create the popup URL and open it.
        var fullUrl = location.protocol + '//' + location.hostname + (location.port ? ':' + location.port : '') + '/popup.html';

        // height and width are percentages of the size of the screen.
        Office.context.ui.displayDialogAsync(fullUrl,
            {height: 45, width: 55},
            function (result) {
                dialog = result.value;
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);

                /*Events are sent by the platform in response to user actions or errors. For example, the dialog is closed via the 'x' button*/
                dialog.addEventHandler(Office.EventType.DialogEventReceived, eventHandler);
            });
    }

// This handler responds to the success or failure message that the pop-up dialog receives from the identity provider.
    function processMessage(arg) {
        var messageFromPopupDialog = JSON.parse(arg.message);
        console.log('Dialog sent a message: ' + JSON.stringify(messageFromPopupDialog));
        if (messageFromPopupDialog === "success") {

            // The Auth0 token has been received, so close the dialog, use
            // the token to get user information, and redirect the task
            // pane to the landing page.

            dialog.close();
            console.log("Auth2 response: " + JSON.stringify(messageFromPopupDialog));

        } else {

            // Something went wrong with authentication or the authorization of the web application,
            // either with Auth0 or with the provider.
            dialog.close();
            console.error("User authentication and application authorization",
                "Unable to successfully authenticate user or authorize application: " + messageFromPopupDialog.error);
        }
    }

    //this handler reponds to whatever happened with auth dialog
    // if dialog's been closed (itself or by user), onStateChange is called
    function eventHandler(arg) {

        // In addition to general system errors, there are 2 specific errors
        // and one event that you can handle individually.
        switch (arg.error) {
            case 12002:
                console.error("Cannot load URL, no such page or bad URL syntax.");
                break;
            case 12003:
                console.error("HTTPS is required.");
                break;
            case 12006:
                // The dialog was closed, typically because the user the pressed X button.
                console.log("Dialog closed by user");
                onStateChange();
                break;
            default:
                onStateChange();
                console.error("Undefined error in dialog window");
                break;
        }
    }


    //reload contacts list
    function reloadFolders() {
        $('#sync_div').show();
        $("#foldersList").empty();

        if(Array.isArray(folders))
        {
            if(folders.length > 0)
            {
                folders.forEach((folder, index) => {
                    var str = '<option ' + (index == 0 ? 'selected' : '') +'  value="' + index + '"' +
                        '>' + folder.displayName + '</option>';
                    $("#foldersList").append(str);
                });
            }
        }

    };

    async function run() {
        //if we're not authorized to token yet
        if (localStorage.getItem("microsoftToken") === null || typeof localStorage.getItem("microsoftToken") === 'undefined') {
            showLoginPopup();
        }
        else {
            console.log('Token we got: ' + localStorage.getItem("microsoftToken"));
            //get Microsoft Acc contacts first
            var folder_index =  $("#foldersList").val();
            var selectedFolder = folders[folder_index];
            console.log("ContactFolder selected: " + folder_index );
            getContacts(localStorage.getItem("microsoftToken"), selectedFolder, (success) => {
                if (contacts.length > 0)
                {
                    localStorage.setItem('selectedMicrosoftContacts', JSON.stringify(contacts));
                }
                else {
                    console.log("No contacts selected to sync!");
                }
                if(success){
                    //here we got to start synching
                }
            });
        }
    };

    //show alert
    function showAlert(message){
        $('.alert').show();
        $("#error_text").html(message);
    }

//microsoft Outlook Graph API ==========================================================================================

    //get microsoft folders
    function getFolders(accessToken) {

        //$('#syncBtnTitle').html('Getting Outlok Folders...');
        $('#run').hide();

        $('#spinner').show();

        $.ajax({
            url: 'https://graph.microsoft.com/beta/me/contactFolders', //'https://graph.microsoft.com/v1.0/me/contacts',
            dataType: 'json',
            headers: { 'Authorization': 'Bearer ' + accessToken }
        }).done(function(items){

            console.log("Folders: " + JSON.stringify(items));
            folders = items.value;
            reloadFolders();
            //$('#syncBtnTitle').html('Refresh Outlook Folders');
            $('#run').show();
            $('#spinner').hide();
        }).fail(function(error){
            $('#spinner').hide();
            console.log("Error: " + JSON.stringify(error));
            if(error.status)
            {
                $('#syncBtnTitle').html('Select Outlook Folders');
                if(error.status != 200)
                {
                    showLoginPopup();
                }
            }
        });

    };

    //get microsoft contacts
    function getContacts(accessToken, selectedFolder, callback) {

        //$('#syncBtnTitle').html('Getting Folder Contacts...');
        $('#run').hide();

        $('#spinner').show();

        $.ajax({
            url: 'https://graph.microsoft.com/v1.0/me/contactfolders/' + selectedFolder.id + '/contacts', //'https://graph.microsoft.com/v1.0/me/contacts',
            dataType: 'json',
            headers: { 'Authorization': 'Bearer ' + accessToken }
        }).done(function(items){

            console.log("Contacts: " + JSON.stringify(items));
            contacts = items.value;
            callback(true);
            $('#syncBtnTitle').html('Refresh Outlook Folders');
            $('#run').show();
            $('#spinner').hide();

        }).fail(function(error){

            $('#spinner').hide();
            console.log("Error: " + JSON.stringify(error));
            callback(false);
            if(error.status)
            {
                $('#buttonTitle').html('Select Outlook Folders');
                if(error.status != 200)
                {
                    showLoginPopup();
                }
            }
        });

    }
//=======================================================================================================================================



});