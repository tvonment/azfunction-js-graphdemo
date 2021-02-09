import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import axios, { AxiosRequestConfig } from 'axios';
import qs = require('qs');

const APP_ID = process.env["APP_ID"];
const APP_SECRET = process.env["APP_SECRET"];
const TENANT_ID = process.env["TENANT_ID"];

const TOKEN_ENDPOINT = 'https://login.microsoftonline.com/' + TENANT_ID + '/oauth2/v2.0/token';
const MS_GRAPH_SCOPE = 'https://graph.microsoft.com/.default';
const MS_GRAPH_ENDPOINT = 'https://graph.microsoft.com/v1.0/';


const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    context.log('HTTP trigger function processed a request.');
    
    // Set Default Header for Axios Requests
    axios.defaults.headers.post['Content-Type'] = 'application/x-www-form-urlencoded';

    // Get Token for MS Graph
    let token = await getToken();

    context.res = {
        // status: 200, /* Defaults to 200 */
        body: token
    };

};

export default httpTrigger;

/**
 * Get Token for MS Graph
 */
async function getToken(): Promise<string> {
    const postData = {
        client_id: APP_ID,
        scope: MS_GRAPH_SCOPE,
        client_secret: APP_SECRET,
        grant_type: 'client_credentials'
    };

    return await axios
        .post(TOKEN_ENDPOINT, qs.stringify(postData))
        .then(response => {
            // console.log(response.data);
            return response.data.access_token;
        })
        .catch(error => {
            console.log(error);
        });
}