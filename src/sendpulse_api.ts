class SendPulseClient {

    private clientID: string = '';
    private clientSecret: string = '';
    public token: string = '';
    public baseURL: string = '';

    constructor (id: string, secret: string, baseurl: string) {
        this.clientID = id;
        this.clientSecret = secret;
        this.token ='';
        this.baseURL = baseurl;
    }

    getToken(callback) {

        let parameters = {
            'grant_type': 'client_credentials',
            'client_id' : this.clientID,
            'client_secret': this.clientSecret
        };

        //'https://api.sendpulse.com/oauth/access_token'
        this.makeRequest(parameters, 'POST', this.baseURL + 'sendpulsetoken', false, callback);
    }

    getAddressBooks(callback){

        this.token = localStorage.getItem("sendPulseToken");
        let parameters = {
            'url': 'listAddressBooks',
            'token': this.token
        }
        //' https://api.sendpulse.com/addressbooks'
        this.makeRequest(parameters, 'POST', this.baseURL + 'sendpulse', true, callback);
    }

    getAddressBookContacts(bookID, callback) {

        //'https://api.sendpulse.com/addressbooks/'+ bookID +'/emails';
        this.token = localStorage.getItem("sendPulseToken");
        let parameters = {
            'url': 'getEmailsFromBook',
            'id' : bookID,
            'token': this.token
        }

        this.makeRequest(parameters, 'POST', this.baseURL + 'sendpulse', true, callback);
    }

    makeRequest(data, method: string, url: string, authorized: boolean, callback) {

        let self = this;

        var xhr = new XMLHttpRequest();
        xhr.open(method, url, true);
        xhr.setRequestHeader('Content-Type', 'application/json');
        if(authorized){
            xhr.setRequestHeader('Authorization', 'Bearer ' + this.token);
        }
        xhr.send(JSON.stringify(data));

        xhr.addEventListener('load', function(e) {
            var jsonResponse = JSON.parse(xhr.responseText);
            console.log("Response: " + JSON.stringify(jsonResponse));

            if(! (jsonResponse.access_token === null ))
            {
                console.log("Token: " + jsonResponse.access_token);
                self.token =  jsonResponse.access_token;
            }
            else {
                console.log("Auth error: " + JSON.stringify(jsonResponse));
            }

            callback(jsonResponse);

        });


    }

}

export { SendPulseClient };