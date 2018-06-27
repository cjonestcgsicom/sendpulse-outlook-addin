/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import {SendPulseClient} from './sendpulse_api';

$(document).ready(() => {

   var sendPulseToken = '';
   var sendPulseID = '';
   var sendPulseSecret = '';
   var baseURL = localStorage.getItem('baseURL');

    if (localStorage.getItem("baseURL") === null) {
      baseURL = 'https://localhost:3000/';
      localStorage.setItem('baseURL', baseURL);
    }


    $('#next').click(next);
    $('#spinner').hide();
    $("#next").hide();


// The initialize function must be run each time a new page is loaded
    Office.initialize = (reason) => {
        $('#sideload-msg').hide();
        $('#app-body').show();

        $('#spAccountButton').click(login);
        $('.alert').hide();
        $('#sync_div').hide();
        $('.close').click((e)=>{
            $('.alert').hide();
        });
        $('[data-dismiss]').click((e)=>{
            $('.alert').hide();
        });

    };


    if (localStorage.getItem("sendPulseToken") === null) {
        $('#login_div').show();
    }
    else {
        //$('#login_div').hide();
        //$('#sync_div').show();
        sendPulseToken = localStorage.getItem("sendPulseToken");

        //set stored email an password
         sendPulseID = localStorage.getItem("sendPulseID");
         sendPulseSecret = localStorage.getItem("sendPulseSecret");
        $('#client_id').val(sendPulseID);
        $('#secret').val(sendPulseSecret);

    }


    async function next() {

    }

    function login(){

        sendPulseID =  $('#client_id').val();
        sendPulseSecret = $('#secret').val();

        $('#spinner').show();

        let sendPulseClient = new SendPulseClient(sendPulseID, sendPulseSecret, baseURL);

        sendPulseClient.getToken((response) => {
            $('#spinner').hide();

            if(response.access_token.length){

                sendPulseToken = response.access_token;
                $('#login_div').hide();
                $("#next").show();
                saveSendPulseToken(sendPulseToken);
                saveSendPulseID(sendPulseID);
                saveSendPulseSecret(sendPulseSecret);
                sendPulseClient.getAddressBooks((res)=> {

                });
            }
            else {
                showAlert("Authentication failed! " + response.message ? response.message : "");
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

    //show alert
    function showAlert(message){
        $('.alert').show();
        $("#error_text").html(message);
    }

});