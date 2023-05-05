import DynamicsWebApi from 'dynamics-web-api';
import * as msal from '@azure/msal-node';

export interface Dynamics365Config {
  resource: string;
  applicationId: string;
  clientSecret: string;
  tenant: string;
  authorityHostUrl: string;
  authorityUrl: string;
  webApiUrl: string;
};


export class Dynamics365 {
  config = {} as Dynamics365Config;
  msalConfig = {
    auth: {
      authority: this.config.authorityUrl,
      clientId: this.config.applicationId,
      clientSecret: this.config.clientSecret,
      knownAuthorities: [this.config.authorityHostUrl],
    },
  };
  cca = null as any;
  dynamicsWebApi = null as any;
  constructor(config: Dynamics365Config) {
    this.config = config;
    this.msalConfig = {
      auth: {
        authority: this.config.authorityUrl,
        clientId: this.config.applicationId,
        clientSecret: this.config.clientSecret,
        knownAuthorities: [this.config.authorityHostUrl],
      },
    };
    this.cca = new msal.ConfidentialClientApplication(this.msalConfig);
    this.dynamicsWebApi = new DynamicsWebApi({
      webApiUrl: this.config.webApiUrl,
      onTokenRefresh: this.acquireToken,
    });
  }

  acquireToken = (dynamicsWebApiCallback: any) => {
    this.cca.acquireTokenByClientCredential({
        scopes: [`${this.config.resource}/.default`],
      })
      .then((response: any) => {
        //call DynamicsWebApi callback only when a token has been retrieved successfully
        dynamicsWebApiCallback(response.accessToken);
      })
      .catch((error: any) => {
        console.debug(JSON.stringify(error));
      });
  };
}
