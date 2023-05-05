# dynamic365_msal

use msal to connect dynamic365

import { Dynamics365 } from'./src';



const dynamic365 = new Dynamics365({

resource: //CRM Organization URL,

applicationId: 'xxxxxxxx-aaaa-bbbb-cccc-dddddddddddd', // Application Id of app registered under AAD.//Dynamics 365 Client Id when registered in Azure

clientSecret: 'fslfjdskljksdljfksdjf;dsjf;kjsdf;kj', // Secret generated for app. Read this environment constiable.

tenant: 'xxxxxxxx-aaaa-bbbb-cccc-dddddddddddd', // AAD Tenant name

authorityHostUrl:'https://login.microsoftonline.com',

authorityUrl:`https://login.microsoftonline.com/${tenant}`,

webApiUrl:  `${resource}/api/data/v9.0/`,

  });

constresults = awaitdynamic365.dynamicsWebApi.retrieveAllRequest(properties);
