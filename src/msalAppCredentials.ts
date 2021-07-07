import { AppCredentials } from "botframework-connector";
import { ConfidentialClientApplication } from "@azure/msal-node";
import { TokenResponse } from "adal-node";

/**
 * An implementation of AppCredentials that uses @azure/msal-node to fetch tokens.
 */
export class MsalAppCredentials extends AppCredentials {
  constructor(
    private readonly clientApplication: ConfidentialClientApplication,
    appId: string,
    scope: string
  ) {
    super(appId, undefined, scope);
  }

  async getToken(forceRefresh: boolean): Promise<string> {
    const scopePostfix = "/.default";
    let scope = this.oAuthScope;
    if (!scope.endsWith(scopePostfix)) {
      scope = `${scope}${scopePostfix}`;
    }

    const token = await this.clientApplication.acquireTokenByClientCredential({
      scopes: [scope],
      skipCache: forceRefresh,
    });

    const { accessToken } = token ?? {};
    if (typeof accessToken !== "string") {
      throw new Error("Authentication: No access token received from MSAL.");
    }

    return accessToken;
  }

  protected async refreshToken(): Promise<TokenResponse> {
    throw new Error("NotImplemented");
  }
}
