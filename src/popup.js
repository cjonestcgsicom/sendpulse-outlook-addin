$(document).ready(() => {

    try {
        // Redirect to Auth2 and tell it which provider to use.

        var auth2AuthorizeEndPoint = 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=' + localStorage.getItem('clientID')
            + '&redirect_uri=' + localStorage.getItem('redirectURI')
            + '&response_type=code&scope=Contacts.ReadWrite'
            + '&response_mode=query'
            + '&state=12345';


        $("#msAccountButton").click(function () {
            redirectToIdentityProvider('windowslive');
            $("#btnTitle").html('Loading auth ' + auth2AuthorizeEndPoint);

        });

        function redirectToIdentityProvider(provider) {

            console.log("go to auth: " + auth2AuthorizeEndPoint);
            window.location.replace(auth2AuthorizeEndPoint);
        }

    }
    catch (err) {
        console.log(err.message);
    }
});
