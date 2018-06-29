var SendPulseClient = /** @class */ (function () {
    function SendPulseClient(id, secret, baseurl) {
        this.clientID = '';
        this.clientSecret = '';
        this.token = '';
        this.baseURL = '';
        this.clientID = id;
        this.clientSecret = secret;
        this.token = '';
        this.baseURL = baseurl;
    }
    SendPulseClient.prototype.getToken = function (callback) {
        var parameters = {
            'grant_type': 'client_credentials',
            'client_id': this.clientID,
            'client_secret': this.clientSecret
        };
        this.makeRequest(parameters, 'POST', 'https://api.sendpulse.com/oauth/access_token', false, callback);
    };
    SendPulseClient.prototype.getAddressBooks = function (callback) {
        this.makeRequest(null, 'GET', ' https://api.sendpulse.com/addressbooks', true, callback);
    };
    SendPulseClient.prototype.getAddressBookContacts = function (bookID, callback) {
        var url = 'https://api.sendpulse.com/addressbooks/' + bookID + '/emails';
        this.makeRequest(null, 'GET', url, true, callback);
    };
    SendPulseClient.prototype.makeRequest = function (data, method, url, authorized, callback) {
        var settings = {
            url: url,
            type: method,
            headers: { 'Content-Type': 'application/json' }
        };
        if (data) {
            settings['contentType'] = "application/json";
            settings['data'] = JSON.stringify(data);
        }
        if (authorized) {
            settings.headers['Authorization'] = this.token;
        }
        $.ajax(settings).done(function (data) {
            console.log("Token received: " + JSON.stringify(data));
            callback(data);
        }).fail(function (error) {
            console.log("Error: " + JSON.stringify(error));
            callback(error);
            if (error.status) {
                if (error.status != 200) {
                }
            }
        });
    };
    return SendPulseClient;
}());
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map
//# sourceMappingURL=sendpulse_api.js.map