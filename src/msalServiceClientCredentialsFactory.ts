// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { ConfidentialClientApplication } from "@azure/msal-node";
import { MsalAppCredentials } from "./msalAppCredentials";
import { ServiceClientCredentials } from "@azure/ms-rest-js";

import {
  AuthenticationConstants,
  MicrosoftAppCredentials,
  ServiceClientCredentialsFactory,
} from "botframework-connector";

/**
 * An implementation of ServiceClientCredentialsFactory that generates MsalAppCredentials
 */
export class MsalServiceClientCredentialsFactory
  implements ServiceClientCredentialsFactory {
  constructor(
    private readonly appId: string,
    private readonly clientApplication: ConfidentialClientApplication
  ) {}

  async isValidAppId(appId: string): Promise<boolean> {
    return appId === this.appId;
  }

  async isAuthenticationDisabled(): Promise<boolean> {
    return !this.appId;
  }

  async createCredentials(
    appId: string,
    audience: string,
    _loginEndpoint: string,
    _validateAuthority: boolean
  ): Promise<ServiceClientCredentials> {
    if (!(await this.isValidAppId(appId))) {
      throw new Error("Invalid appId.");
    }

    return new MsalAppCredentials(
      this.clientApplication,
      appId,
      audience ?? AuthenticationConstants.ToBotFromChannelTokenIssuer
    );
  }
}
