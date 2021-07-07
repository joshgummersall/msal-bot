# msal bot

This repo includes a Microsoft Bot Framework Bot using Cloud Adapter and @azure/msal-node with an HTTP proxy.

`CloudAdapter` was introduced in `4.14.0`, so it is relatively new.

The `index.ts` file is a little different as it includes the proper code to create a `CloudAdapter` instance that
uses `@azure/msal-node` authentication. Much of that code is derived from https://github.com/microsoft/botbuilder-js/pull/3848.

The HTTP client implementation is adapted from https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-node/src/network/HttpClient.ts.

In order to run this sample, update the `.env` file with a valid App ID and password.
