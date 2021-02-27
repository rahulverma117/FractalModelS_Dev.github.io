
(function () {
    "use strict";

    var cellToHighlight;
    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        console.log('here');
        //var auth = initCognitoSDK();
     //   console.log(auth);

        $(document).ready(function () {


            $('#LoginButton').click(OnClickLoginButton);
            $('#CloseButton').click(OnCloseButton);



        });
    };
})();


function OnCloseButton() {
    Office.context.ui.messageParent(false);
}

function OnClickLoginButton() {

    var username = $("#username").val();  //dougsawyer
    var password = $("#password").val(); //9*ck2CT@waY6

    var authenticationData = {
        Username: username,
        Password: password,
    };
    var authenticationDetails = new AmazonCognitoIdentity.AuthenticationDetails(authenticationData);
    var poolData = {
        UserPoolId: 'us-west-2_HmAsbnylC',
        ClientId: 'u23tqn215tlmhac1ju8e9khhq'
    };
    var userPool = new AmazonCognitoIdentity.CognitoUserPool(poolData);
    var userData = {
        Username: username,
        Pool: userPool
    };
    var cognitoUser = new AmazonCognitoIdentity.CognitoUser(userData);
    cognitoUser.authenticateUser(authenticationDetails, {
        onSuccess: function (result) {


            var AccessToken = result.getAccessToken().getJwtToken();
            var IdToken = result.getIdToken().getJwtToken();
            var RefreshToken = result.getRefreshToken().getToken();



            var tokens = {
                IdToken: IdToken,
                AccessToken: AccessToken,
                RefreshToken: RefreshToken,
                Username: username
            };

            var result = IdToken + "," + AccessToken + "," + RefreshToken + "," + username;

            console.log(result);
            //
            //console.log(accessToken);
            ///* Use the idToken for Logins Map when Federating User Pools with identity pools or when passing through an Authorization Header to an API Gateway Authorizer */
            //var idToken = result.idToken.jwtToken;
            Office.context.ui.messageParent(result);
        },

        onFailure: function (err) {
            $("#errorMsg").text(err);

        },

    });
}