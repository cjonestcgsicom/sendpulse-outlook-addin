$(document).ready(() => {

    var baseURL = localStorage.getItem('baseURL');

    // The initialize function must be run each time a new page is loaded
    Office.initialize = () => {

        var url_string = window.location.href;

        var pass = '';
        var url = new URL(url_string);
        var data = {
            code : url.searchParams.get("code"),
            redirect_uri : localStorage.getItem('redirectURI'),
            client_id : localStorage.getItem('clientID'),
            client_secret : localStorage.getItem("secret"),
            grant_type : 'authorization_code'
        };


        if(data.code) {
            setTempCode(data.code);
            getToken(data);
        } else
        {
            $("#auth_response").html("Auth error: " + url.searchParams.get("error_description"));
            closeWindowOnTimeout(2);
        }
    };


    function setTempCode(code)
    {
        localStorage.removeItem('code');
        localStorage.setItem('code', code);

        // var data = 'success';
        // Office.context.ui.messageParent(JSON.stringify(data));

    }

    function getToken(data)
    {

        var getTokenURL = baseURL + 'token';//'https://login.microsoftonline.com/common/oauth2/v2.0/token';
        var data_str = 'grant_type=authorization_code&code=' + data.code
            + '&redirect_uri=' + data.redirect_uri
            + '&client_id=' + data.client_id
            + '&client_secret=' + data.client_secret;


        var data_json = {
            'grant_type' : 'authorization_code',
            'code' : data.code,
            'redirect_uri': data.redirect_uri,
            'client_id': data.client_id,
            'client_secret' : data.client_secret
        };


        console.log(JSON.stringify(data_json));

        var xhr = new XMLHttpRequest();
        var urlset = getTokenURL;
        xhr.open("POST", urlset, true);
        xhr.setRequestHeader('Content-Type', 'application/json');
        xhr.send(JSON.stringify(data_json));

        xhr.addEventListener('load', function(e) {
            var jsonResponse = JSON.parse(xhr.responseText);
            console.log("Response: " + JSON.stringify(jsonResponse));

            if(! (jsonResponse.access_token === null ))
            {
                console.log("Token: " + jsonResponse.access_token);
                saveToken(jsonResponse.access_token);
            }
            else {
                $("#auth_response").html("Auth error: " + JSON.stringify(jsonResponse));
                closeWindowOnTimeout(0.5);
            }

        });

    }

    //set new token
    function saveToken(tokenStr)
    {
        localStorage.removeItem('microsoftToken');
        localStorage.setItem('microsoftToken', tokenStr);

        //var data = 'success';
        //Office.context.ui.messageParent(JSON.stringify(data));
        closeWindowOnTimeout(0.5);

    };

    function closeWindowOnTimeout(seconds) {
        setTimeout(function() {
            window.close();
        }, seconds*1000);
    };

});