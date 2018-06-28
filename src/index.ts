/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import {SendPulseClient} from './sendpulse_api';
import listTable2_Accent1 = Word.Style.listTable2_Accent1;

$(document).ready(() => {

   var sendPulseToken = '';
   var sendPulseID = '';
   var sendPulseSecret = '';
   var baseURL = localStorage.getItem('baseURL');
   var addressBooks = [];
   var emails = [];
   var selectedAddressBook = null;

    if (localStorage.getItem("baseURL") === null) {
      baseURL = 'https://localhost:3000/';
      localStorage.setItem('baseURL', baseURL);
    }


    $('#next').click(next);
    $('#logout').click(logout);
    $('#spinner').hide();
    $("#next").hide();
    $('#sync_div').hide();
    $('#login_div').hide();
    $('#logout').hide();

// The initialize function must be run each time a new page is loaded
    Office.initialize = (reason) => {
        $('#sideload-msg').hide();
        $('#app-body').show();

        $('#spAccountButton').click(login);
        $('.alert').hide();

        $('.close').click((e)=>{
            $('.alert').hide();
        });
        $('[data-dismiss]').click((e)=>{
            $('.alert').hide();
        });

    };

    statusCheck();




    function statusCheck()
    {
        if (localStorage.getItem("sendPulseToken") === null) {
            $('#login_div').show();
            $('#logout').hide();
            //set stored email an password
            sendPulseID = localStorage.getItem("sendPulseID");
            sendPulseSecret = localStorage.getItem("sendPulseSecret");
            $('#client_id').val(sendPulseID);
            $('#secret').val(sendPulseSecret);
        }
        else {
            sendPulseToken = localStorage.getItem("sendPulseToken");
            sendPulseID = localStorage.getItem("sendPulseID");
            sendPulseSecret = localStorage.getItem("sendPulseSecret");
            $('#client_id').val(sendPulseID);
            $('#secret').val(sendPulseSecret);

            let sendPulseClient = new SendPulseClient(sendPulseID, sendPulseSecret, baseURL);

            getAddressBooksList(sendPulseClient, true, (success) => {
                if(!success)
                {
                    getToken(sendPulseClient, (success) => {});
                }
                $('#logout').show();

            });

        }
    }

    function getToken(sendPulseClient: SendPulseClient, callback) {

        $('#spinner').show();

        sendPulseClient.getToken((response) => {
            $('#spinner').hide();

            if(response.access_token.length){
                sendPulseToken = response.access_token;
                $('#login_div').hide();
                saveSendPulseToken(sendPulseToken);
                saveSendPulseID(sendPulseID);
                saveSendPulseSecret(sendPulseSecret);
                getAddressBooksList(sendPulseClient, false, (success)=> {});
                $('#logout').show();

                callback(true);
            }
            else {
                callback(false);
                $('#login_div').show();
                $('#logout').hide();
                showAlert("Authentication failed! " + response.message ? response.message : "");
            }
        });
    }

    function getAddressBooksList(sendPulseClient: SendPulseClient, mute, callback)
    {
        $('#spinner').show();
        sendPulseClient.getAddressBooks((res)=> {
            $('#spinner').hide();
            if(res.error)
            {
                if(!mute)
                {
                    showAlert(res.message ? res.message : res.error_description);
                }
                callback(false);
            }
            else
            {
                $("#next").show();
                callback(true);
                //here make all list visible
                showAddressBooksList(res)
            }
        });
    }

    function getAddressBooksContacts(sendPulseClient: SendPulseClient, bookID, mute, callback)
    {
        $('#spinner').show();
        sendPulseClient.getAddressBookContacts(bookID,(res)=> {
            $('#spinner').hide();
            if(res.error)
            {
                if(!mute)
                {
                    showAlert(res.message ? res.message : res.error_description);
                }
                callback(false);
            }
            else
            {
                emails = res;
                saveSendPulseEmails();
                if(Array.isArray(emails))
                {
                    console.log(emails.length + " emails retrieved from address book");
                }
                callback(true);
            }
        });
    }


    function showAddressBooksList(list_array)
    {
        addressBooks = list_array;

        $("#book_select").empty();
        $('#sync_div').show();

        list_array.forEach((book, index) => {
            var str = '<option ' + (index == 0 ? 'selected' : '') +'  value="' + index + '"' +
                '>' + book.name + '</option>';
            $("#book_select").append(str);
        });
    }





    function logout(){
        localStorage.removeItem('sendPulseToken');
        localStorage.removeItem('sendPulseID');
        localStorage.removeItem('sendPulseSecret');
        localStorage.removeItem('emailsSelected');

        $('#login_div').show();
        $('#next').hide();
        $('#sync_div').hide();
    }

    function login(){

        sendPulseID =  $('#client_id').val();
        sendPulseSecret = $('#secret').val();

        let sendPulseClient = new SendPulseClient(sendPulseID, sendPulseSecret, baseURL);
        getToken(sendPulseClient, (success) => {});
    }


    async function next() {

        //get and save master mode
        var master_mode = $('input[name="inlineRadioOptions"]:checked').val();
        console.log("Master mode: " + (master_mode == 1 ? 'SendPulse is Master' : 'Outlook is Master'));
        saveMasterMode(master_mode);

        //get the selected address-book

        var book_index =  $("#book_select").val();;
        console.log("AddressBooks selected index: " + book_index );

        if(book_index < addressBooks.length)
        {
            selectedAddressBook = addressBooks[book_index];
        }
        else
        {
            showAlert("You haven'tselected andy address book to synch your contacts with");
            return;
        }

        //get the contacts
        let sendPulseClient = new SendPulseClient(sendPulseID, sendPulseSecret, baseURL);
        getAddressBooksContacts(sendPulseClient, selectedAddressBook.id, false, (success) => {
            if(success)
            {
                //go to next page
                var nextUrl = location.protocol + '//' + location.hostname + (location.port ? ':' + location.port : '') + '/synchdialog.html';
                window.location.replace(nextUrl);
            }
            else
            {
                statusCheck();
            }
        });

    }

    //------------ data storing
    //set new greenRope token
    function saveSendPulseToken(tokenStr)
    {
        sendPulseToken = tokenStr;
        localStorage.removeItem('sendPulseToken');
        localStorage.setItem('sendPulseToken', tokenStr);
    }

    //save email to localStorage
    function saveSendPulseID(idStr)
    {
        sendPulseID = idStr;
        localStorage.removeItem('sendPulseID');
        localStorage.setItem('sendPulseID', idStr);
    }

    //save login to localStorage
    function saveSendPulseSecret(passswdStr)
    {
        sendPulseSecret = passswdStr;
        localStorage.removeItem('sendPulseSecret');
        localStorage.setItem('sendPulseSecret', passswdStr);
    }

    //save emails we got to localStorage
    function saveSendPulseEmails()
    {
        let emails_stringified = JSON.stringify(emails);
        localStorage.removeItem('emailsSelected');
        localStorage.setItem('emailsSelected', emails_stringified);
    }

    function saveMasterMode(masterMode: number){
        localStorage.removeItem('masterMode');
        localStorage.setItem('masterMode', masterMode.toString());
    }

    //show alert
    function showAlert(message){
        $('.alert').show();
        $("#error_text").html(message);
    }



});