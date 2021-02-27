
(function () {
    "use strict";

    var cellToHighlight;
    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
 //   Office.initialize = function (reason) {
        console.log('here');
        //var auth = initCognitoSDK();
     //   console.log(auth);
        //var user = await auth.signIn('dougsawyer', '9*ck2CT@waY6');
        $(document).ready(function () {
          

            $('#LoginButton').click(OnClickLoginButton);
            $('#CloseButton').click(OnCloseButton);



        });
    //};
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


            const AccessToken = result.getAccessToken().getJwtToken();
            const IdToken = result.getIdToken().getJwtToken();
            const RefreshToken = result.getRefreshToken().getToken();

        

            const tokens = {
                IdToken: IdToken,
                AccessToken: AccessToken,
                RefreshToken: RefreshToken,
                Username: username
            };
          
            //
            //console.log(accessToken);
            ///* Use the idToken for Logins Map when Federating User Pools with identity pools or when passing through an Authorization Header to an API Gateway Authorizer */
            //var idToken = result.idToken.jwtToken;
            Office.context.ui.messageParent(tokens);
        },

        onFailure: function (err) {
            $("#errorMsg").text(err);
            
        },

    });
}