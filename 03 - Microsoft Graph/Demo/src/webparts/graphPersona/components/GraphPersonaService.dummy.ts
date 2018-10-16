import { User } from '@microsoft/microsoft-graph-types';
import { IGraphPersonaService } from './IGraphPersonaService';

export class GraphPersonaService implements IGraphPersonaService {
    constructor() {
    }

    public getProfileInfo(): Promise<User> {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users/$entity",
          "id": "48d31887-5fad-4d73-a9f5-3c356e68a038",
          "businessPhones": [
              "+1 412 555 0109"
          ],
          "displayName": "Megan Bowen",
          "givenName": "Megan",
          "jobTitle": "Auditor",
          "mail": "MeganB@M365x214355.onmicrosoft.com",
          "mobilePhone": null,
          "officeLocation": "12/1110",
          "preferredLanguage": "en-US",
          "surname": "Bowen",
          "userPrincipalName": "MeganB@M365x214355.onmicrosoft.com",
          // The following are erroneously required
          "assignedLicenses": [],
          "assignedPlans": [],
          "provisionedPlans": [],
          "proxyAddresses": [],
          "birthday": null,
          "hireDate": null,
          "deviceEnrollmentLimit": 0
        });
    }
    public getPhoto(): Promise<string> {
        return Promise.resolve("https://localhost:4321/megan.jpg");
    }
}