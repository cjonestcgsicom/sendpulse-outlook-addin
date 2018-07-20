import {SendPulseClient} from './sendpulse_api';

$(document).ready(() => {

    $('#run').hide();
    $('#run').click(run);

    var Auth2AccountData = Auth2AccountData || {};
    var dialog;
    var contacts = [];
    var folders = [];
    var selectedFolder = {};
    var baseURL = localStorage.getItem("baseURL");
    var sendPulseID = localStorage.getItem("sendPulseID") || '';
    var sendPulseSecret = localStorage.getItem("sendPulseSecret") || '';

    var counter_max = 0;
    var counter = 0;

    Auth2AccountData.secret = 'nnoHKD75!)gwzrSICW728?[';
    Auth2AccountData.clientID = 'c80d93b6-a45b-405c-add9-388a25beadb9';

// The Auth0 subdomain and client ID need to be shared with the popup dialog
    localStorage.setItem('clientID', Auth2AccountData.clientID);
    localStorage.setItem('secret', Auth2AccountData.secret);
    localStorage.setItem('redirectURI', baseURL + 'function-file/function-file.html');
    localStorage.removeItem('selectedMicrosoftContacts');

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

// The initialize function must be run each time a new page is loaded
    Office.initialize = (reason) => {
        /*
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
        */
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
            selectedFolder = folders[folder_index];
            console.log("ContactFolder selected: " + folder_index );
            //
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
                    syncContacts();
                }
            });
        }
    };

    /**
    *
    * synching contacts with given master mode (sendPulse of outlook)
    * */
    function syncContacts() {

        //1) get the addressBook ID for SendPulse
        let addressBookId = localStorage.getItem("bookID");

        //2) get the contacts from Outlook
        let outlookContacts = contacts;

        //3) get the contacts from Sendpulse
        let sendPulseContacts = JSON.parse(localStorage.getItem("emailsSelected"));

        //4) the master mode (1 - import from SendPulse, 2 - export from Outlook)
        let masterMode = localStorage.getItem("masterMode");

        //5) syncronize

        //5.1 - when we export from outlook to sendPulse
        if (masterMode === '2') {

            if (!(Array.isArray(outlookContacts)) || outlookContacts.length === 0) {
                showAlert("You don't have any Outlook contacts to eхport from selected folder!");
                return;
            }

            $('#spinner').show();
            $('#run').hide();
            $('#progress_dynamics').show();
            counter_max = outlookContacts.length;
            counter = 0;

            var emails = [];
            outlookContacts.forEach((contact, index) => {
                contact.emailAddresses.forEach((email) => {
                    if (email.address) {
                        console.log("Outlook contact email : " + JSON.stringify(email.address.toLowerCase()));
                        let email_object = {
                            'email': email.address.toLowerCase(),
                            'variables': {'firstName': '', 'lastName': ''}
                        };
                        if(contact.givenName) {
                            email_object.variables.firstName = contact.givenName;
                        }
                        if(contact.surname) {
                            email_object.variables.lastName = contact.surname;
                        }
                        emails.push(email_object);
                    }
                });
                update_counter();
            });
            let sendPulseClient = new SendPulseClient(sendPulseID, sendPulseSecret, baseURL);

            sendPulseClient.addEmailsToAddressBook(addressBookId, emails, (res) => {
                $('#spinner').hide();
                $('#progress_dynamics').hide();
                $('#run').show();
                if (res.error) {
                        showAlert(res.message ? res.message : res.error_description);
                }
                else {
                    console.log("added contacts to sendpulse: ", emails.length);
                }
            });


        } else //5.2 - when we import from sendPulse
        if (masterMode === '1') {

            console.log("Export from SendPulse");
            if (!(Array.isArray(sendPulseContacts)) || sendPulseContacts.length === 0) {
                showAlert("You don't have any SendPulse emails to import!");
                return;
            }
            $('#spinner').show();
            $('#run').hide();
            $('#progress_dynamics').show();
            counter_max = sendPulseContacts.length;
            counter = 0;

            sendPulseContacts.forEach((contact) => {
                let newMicosoftContact = {
                    "GivenName": contact["email"].split("@")[0],
                    "EmailAddresses": [
                        {
                            "Address": contact["email"],
                            "Name":  contact["email"].split("@")[0]
                        }
                    ]
                };

                if(contact.variables)
                {
                    if(contact.variables.firstName)
                    {
                        newMicosoftContact.GivenName = contact.variables.firstName;
                    }
                    if(contact.variables.lastName)
                    {
                        newMicosoftContact["surname"] = contact.variables.lastName;
                    }
                };

                outlookContacts.forEach((microsoftContact) => {

                    if (isSendPulseContactDuplicated(contact, microsoftContact)) {
                       newMicosoftContact = null;
                    }

                });
                if (newMicosoftContact != null) {
                    addOutlookContact(localStorage.getItem("microsoftToken"), selectedFolder, newMicosoftContact, (success) => {
                        console.log(success);
                        update_counter();
                    });
                }


            });


        }
    }

    function isSendPulseContactDuplicated(contact, microsoftContact) {
        let duplicate = false;
        microsoftContact.emailAddresses.forEach((email) => {
           if(email.address && contact["email"]) {
               if (email.address.toLowerCase() === contact["email"].toLowerCase()) {
                   duplicate = true;
                   console.log("Outlook contact email : " + JSON.stringify(email.address.toLowerCase()));
                   console.log("SP Contact email: " + contact.Email.toLowerCase());
                   return true;
               }
           }
        });
        return duplicate;
    }


    //show alert
    function showAlert(message){
        $('.alert').show();
        $("#error_text").html(message);
    }

    //update synched conteacts counter
    function update_counter() {
        counter += 1;
        console.log("Counter: " + counter + " out of " + counter_max);

        //update progress
        let current_progress = Math.ceil((counter * 100.0)/(counter_max * 1.0));
        console.log("Progress: " + current_progress);

        $("#dynamic")
            .css("width", current_progress + "%")
            .attr("aria-valuenow", current_progress)
            .text(current_progress + "% Complete");

        if(counter == counter_max){
            $('#spinner').hide();
            $('#run').show();
            $('#progress_dynamics').hide();
        }
    };

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

        $('#syncBtnTitle').html('Getting Outlook Folder Contacts...');
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
            $('#syncBtnTitle').html('Synchronize');
            $('#run').show();
            $('#spinner').hide();

        }).fail(function(error){

            $('#spinner').hide();
            console.log("Error: " + JSON.stringify(error));
            callback(false);
            if(error.status)
            {
                $('#syncBtnTitle').html('Refresh Outlook Folders');
                if(error.status != 200)
                {
                    showLoginPopup();
                }
            }
        });

    }

    //create microsoft contact
    function addOutlookContact(accessToken, selectedFolder, contact, callback) {

        $.ajax({
            url: 'https://graph.microsoft.com/v1.0/me/contactfolders/' + selectedFolder.id + '/contacts',
            type: "POST",
            contentType: "application/json",
            data: JSON.stringify(contact),
            headers: { 'Authorization': 'Bearer ' + accessToken,
                'Content-Type': 'application/json' }
        }).done(function(item){

            console.log("Contact created: " + JSON.stringify(item));
            callback(true);

        }).fail(function(error){
            console.log("Error: " + JSON.stringify(error));
            callback(false);
            if(error.status)
            {
                if(error.status != 200)
                {

                }
            }
        });

    }

//create microsoft contact
    function updateOutlookContact(accessToken, selectedFolder, contact, callback) {

        if(!contact.id)
        {
            console.error("Trying to update contact w/i id!");
            return;
        }

        var url = 'https://graph.microsoft.com/v1.0/me/contactfolders/' + selectedFolder.id + '/contacts/' + contact.id

        $.ajax({
            url: url,
            dataType: 'json',
            type: "PATCH",
            contentType: "application/json",
            data: JSON.stringify(contact),
            headers: { 'Authorization': 'Bearer ' + accessToken,
                'Content-Type': 'application/json' }
        }).done(function(item){

            console.log("Contact updated: " + JSON.stringify(item));
            callback(true);

        }).fail(function(error){
            console.log("Error: " + JSON.stringify(error));
            callback(false);
            if(error.status)
            {
                if(error.status != 200)
                {

                }
            }
        });

    }

//=======================================================================================================================================



});