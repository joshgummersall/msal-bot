import axios, { AxiosProxyConfig } from 'axios';
import { memoize } from 'lodash';
import { OpenIdMetadata } from 'botframework-connector/lib/auth/openIdMetadata';

const base64url = require('base64url');
const getPem = require('rsa-pem-from-mod-exp');

export class ProxyOpenIdMetadata extends OpenIdMetadata {
  constructor(metadataUrl: string, proxy: AxiosProxyConfig) {
    super(metadataUrl);

    this.getKey = memoize(async (keyId) => {
      let response = await axios.get(metadataUrl, { proxy });

      const metadata = response.data;
      const { jwks_uri } = metadata ?? {};
      if (typeof jwks_uri !== 'string') {
        return null;
      }

      response = await axios.get(jwks_uri, { proxy });
      const { keys = [] } = response.data ?? {};
      const key = keys.find((key: any) => key.kid === keyId) ?? {};

      // Return null for non-RSA keys
      if (!(key?.n && key?.e)) {
          return null;
      }

      return {
          key: getPem(base64url.toBase64(key.n), key.e),
          endorsements: key.endorsements,
      };
    });
  }
}
