export const Configs = {
    appId: 'YOUR-CLIENTID-OF-REGISTERED-APP',
    scope: 'User.Read Mail.Send',
    graphAuthUrl: 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize',
    graphV1Url: 'https://graph.microsoft.com/v1.0/me/',
    graphBetaUrl: 'https://graph.microsoft.com/beta/me/',
    Language: 'en-US'
};

export const explorerValues: any = {
    selectedOption: 'GET',
    selectedVersion: 'v1.0',
    authentication: {
        user: {}
    },
    showImage: false,
    requestInProgress: false,
    headers: [],
    postBody: ''
};
