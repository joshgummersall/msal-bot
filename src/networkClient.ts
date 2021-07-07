import axios, { AxiosProxyConfig } from "axios";

import {
  INetworkModule,
  NetworkRequestOptions,
  NetworkResponse,
} from "@azure/msal-common";

// Derived from
// https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-node/src/network/HttpClient.ts
export class ProxyHttpClient implements INetworkModule {
  constructor(private readonly proxy?: AxiosProxyConfig) {}

  async sendGetRequestAsync<T>(
    url: string,
    options?: NetworkRequestOptions
  ): Promise<NetworkResponse<T>> {
    const response = await axios({
      headers: options?.headers,
      method: "GET",
      proxy: this.proxy,
      url,
      validateStatus: () => true,
    });

    return {
      headers: response.headers,
      body: response.data as T,
      status: response.status,
    };
  }

  async sendPostRequestAsync<T>(
    url: string,
    options?: NetworkRequestOptions,
    cancellationToken?: number
  ): Promise<NetworkResponse<T>> {
    const response = await axios({
      data: options?.body ?? "",
      headers: options && options.headers,
      method: "POST",
      proxy: this.proxy,
      timeout: cancellationToken,
      url: url,
      validateStatus: () => true,
    });

    return {
      headers: response.headers,
      body: response.data as T,
      status: response.status,
    };
  }
}
