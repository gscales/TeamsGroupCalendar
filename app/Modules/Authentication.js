const Authuser = (upn,appconfig) => {
    return new Promise(
        (resolve, reject) => {
            let config = {
                clientId: appConfig.clientId,
                redirectUri: window.location.origin + appConfig.redirectUri,       // This should be in the list of redirect uris for the AAD app
                cacheLocation: "localStorage",
                navigateToLoginRequestUrl: false,
                endpoints: {
                    "https://graph.microsoft.com": "https://graph.microsoft.com"
                }
            };
            if (upn) {
                config.extraQueryParameters = "scope=openid+profile&login_hint=" + encodeURIComponent(upn);
            } else {
                config.extraQueryParameters = "scope=openid+profile";
            }
            let authContext = new AuthenticationContext(config);
            let user = authContext.getCachedUser();
            if (user) {
                if (user.userName !== upn) {
                    // User doesn't match, clear the cache
                    authContext.clearCache();
                }
            }
            // Get the id token (which is the access token for resource = clientId)
            let token = authContext.getCachedToken(config.clientId);
            if (token) {
                authContext.acquireToken("https://graph.microsoft.com", function (error, idtoken) {
                    if (error || !idtoken) {
                       reject(error);
                    }
                    else
                        resolve(idtoken);
                });
            } else {
                // No token, or token is expired
                authContext._renewIdToken(function (err, idToken) {
                    if (err) {
                        console.log("Renewal failed: " + err);
                        microsoftTeams.authentication.authenticate({
                            url: window.location.origin + appConfig.authwindow,
                            width: 400,
                            height: 400,
                            successCallback: function (t) {
                                // Note: token is only good for one hour
                                token = t;
                                resolve(token);
                            },
                            failureCallback: function (err) {
                                  reject(err);
                            }
                        });
                    } else {
                        authContext.acquireToken("https://graph.microsoft.com", function (error, idtoken) {
                            if (error || !idtoken) {
                               reject(error);
                            }
                            else
                                resolve(idtoken);
                        });
                    }
                });
            }



        }
        );
}
