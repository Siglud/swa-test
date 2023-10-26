import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { OnBehalfOfCredentialAuthConfig, OnBehalfOfUserCredential, UserInfo } from "@microsoft/teamsfx";
import config from "../../config";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
import { Client } from "@microsoft/microsoft-graph-client";
import * as jwt from "jsonwebtoken";
import * as jwksClient from "jwks-rsa";

function validateJwt(token: string): Promise<jwt.JwtPayload> {

    const getSigningKeys = (header: jwt.JwtHeader, callback: jwt.SigningKeyCallback) => {
        var client = jwksClient({
            jwksUri: `https://login.microsoftonline.com/${config.tenantId}/discovery/v2.0/keys`,
        });

        client.getSigningKey(header.kid, function (err, key: jwksClient.SigningKey) {
            var signingKey = key.getPublicKey();
            callback(null, signingKey);
        });
    }

    const validationOptions = {
        audience: config.clientId,
        issuer: `https://login.microsoftonline.com/${config.tenantId}/v2.0`
    };
    return new Promise((resolve, _) => {
        jwt.verify(token, getSigningKeys, validationOptions, (err, decoded) => {
            if (err) {
                resolve(null);
            }
            resolve(decoded as jwt.JwtPayload);
        });
    });
    
}

export async function getUserProfile(req: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    // Initialize response.
    const res: HttpResponseInit = {
        status: 200
    };

    res.cookies = [{name: "myCookie", value: "cookieValue"}];

    const body = Object();

    // Put an echo into response body.
    body.receivedHTTPRequestBody = await req.text() || "";

    // Prepare access token.
    const accessToken: string = req.headers.get("X-Teams-Accesstoken")?.trim();
    if (!accessToken) {
        return {
            status: 400,
            body: JSON.stringify({
                error: "No access token was found in request header.",
            }),
        };
    }

    const validate = await validateJwt(accessToken);
    if (!validate) {
        return {
            status: 401,
            body: JSON.stringify({
                error: "Access token is invalid.",
            }),
        };
    }

    body.accessToken = accessToken;

    const oboAuthConfig: OnBehalfOfCredentialAuthConfig = {
        authorityHost: config.authorityHost,
        clientId: config.clientId,
        tenantId: config.tenantId,
        clientSecret: config.clientSecret,
    };

    body.oboAuthConfig = oboAuthConfig;

    let oboCredential: OnBehalfOfUserCredential;
    try {
        oboCredential = new OnBehalfOfUserCredential(accessToken, oboAuthConfig);
    } catch (e) {
        context.log(e);
        return {
            status: 500,
            body: JSON.stringify({
                error:
                    "Failed to construct OnBehalfOfUserCredential using your accessToken. " +
                    "Ensure your function app is configured with the right Azure AD App registration.",
            }),
        };
    }

    // Query user's information from the access token.
    try {
        const currentUser: UserInfo = await oboCredential.getUserInfo();
        if (currentUser && currentUser.displayName) {
            body.userInfoMessage = `User display name is ${currentUser.displayName}.`;
        } else {
            body.userInfoMessage = "No user information was found in access token.";
        }
    } catch (e) {
        context.log(e);
        return {
            status: 400,
            body: JSON.stringify({
                error: "Access token is invalid." + e,
            }),
        };
    }

    // Create a graph client with default scope to access user's Microsoft 365 data after user has consented.
    try {
        // Create an instance of the TokenCredentialAuthenticationProvider by passing the tokenCredential instance and options to the constructor
        const authProvider = new TokenCredentialAuthenticationProvider(oboCredential, {
            scopes: ["https://graph.microsoft.com/.default"],
        });

        // Initialize Graph client instance with authProvider
        const graphClient = Client.initWithMiddleware({
            authProvider: authProvider,
        });

        const profile: any = await graphClient.api("/me").get();
        body.graphClientMessage = profile;
    } catch (e) {
        context.log(e);
        return {
            status: 500,
            body: JSON.stringify({
                error:
                    "Failed to retrieve user profile from Microsoft Graph. The application may not be authorized." + e,
            }),
        };
    }

    res.body = JSON.stringify(body);

    return res;
}


export async function headerTest(req: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    const res: HttpResponseInit = {
        status: 200,
        body: "",
    };
    let hd = {}
    req.headers.forEach((value, name) => {
        hd[name] = value;
    });
    res.body = JSON.stringify(hd);
    return res;
}

app.http('getUserProfile', {
    methods: ['GET', 'POST'],
    authLevel: 'anonymous',
    handler: getUserProfile
});

app.http('headerTest', {
    methods: ['GET', 'POST'],
    authLevel: 'anonymous',
    handler: headerTest
});