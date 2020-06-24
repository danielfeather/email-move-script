const axios = require('axios');
const qs = require('querystring');

const authConfig = {
    client_id: 'f102452b-b793-4ba8-aab0-4974495c3f84',
    client_secret: 'pU~QlyL3hv~5Dv_AUpo6bh1m0~RJvu-.-8',
    token_endpoint: 'https://login.microsoftonline.com/6e26a76f-adf1-4e9e-8eea-25f26bf52a0f/oauth2/v2.0/token',
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials'
};

const graphConfig = {
    host: 'https://graph.microsoft.com/',
    version: 'v1.0',
    userId: 'cb922f30-bd4d-48f9-b3e3-c9b892c6294e',
    getEndpoint(){
        return this.host.concat(this.version);
    }
};

// Function to acquire an access token for Microsoft Graph using Client Credentials grant type.
async function getAccessToken(){
    try {
        return (await axios.post(
            authConfig.token_endpoint,
            qs.stringify(
                {
                    client_id: authConfig.client_id,
                    client_secret: authConfig.client_secret,
                    scope: authConfig.scope,
                    grant_type: authConfig.grant_type
                }
            ),
            {
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded'
                }
            }
        )).data.access_token;
    } catch (e) {
        return e;
    }
}



(async function () {

    const graphClient = axios.create({
        headers: {
            Authorization: `Bearer ${await getAccessToken()}`
        },
        baseURL: graphConfig.getEndpoint()
    })

    const { data: archive } = await graphClient.get(`/users/${graphConfig.userId}/mailFolders/archive`);

    const storedFolderIds = {};

    const { data: childFolders } = await graphClient.get(`/users/${graphConfig.userId}/mailFolders/archive/childfolders`)

    childFolders.value.every(folder => {
        storedFolderIds[folder.displayName] = folder.id;
    })

    // Variable we can set inside the loop when we check if there is a nextLink.
    // If there is no next link, we set this to true so the loop will break on
    // its next iteration.
    let isEnd = false;

    // We define the endpoint so that it can be overwritten inside the look if
    // a next link is provided.
    let endpoint = `/users/${graphConfig.userId}/mailFolders/archive/messages?$select=id`;

    let { data: messages } = await graphClient.get(endpoint)

    let responses = [];

    let lastGraphBatchResult;

    while(!isEnd){

        let batchRequestBody = {
            requests: []
        };

        let requestId = 0;
        for (const message of messages.value) {

            batchRequestBody.requests.push(
                {
                    id: requestId++,
                    method: 'POST',
                    url: `/users/${graphConfig.userId}/messages/${message.id}/move`,
                    body: {
                        destinationId: 'AAMkADZhYjI0ZTMwLTIzNTMtNDgzMS04ODJjLTNhMjAzYzYwY2NlNgAuAAAAAAAENC3I2XUATZdPFyH_DYPLAQC-CUleI1JNQpuxqWV3wPhOAAAmEWoDAAA='
                    },
                    headers: {
                        "Content-Type": 'application/json'
                    }
                }
            )

            await graphClient.post(
                `/users/${graphConfig.userId}/messages/${message.id}/move`,
                {
                    destinationId: 'AAMkADZhYjI0ZTMwLTIzNTMtNDgzMS04ODJjLTNhMjAzYzYwY2NlNgAuAAAAAAAENC3I2XUATZdPFyH_DYPLAQC-CUleI1JNQpuxqWV3wPhOAAAmEWoDAAA='
                }
            )

        }

        // lastGraphBatchResult = await graphClient.post('/$batch', batchRequestBody)

        if(!messages['@odata.nextLink']){
            isEnd = true;
            continue
        }

        const nextLink = messages['@odata.nextLink'].replace(graphConfig.getEndpoint(), '');
        messages = (await graphClient.get(nextLink)).data;

    }

console.dir(responses);

})()