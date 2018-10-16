import { User } from '@microsoft/microsoft-graph-types';
import { IGraphPersonaService } from './IGraphPersonaService';
import { MSGraphClient } from '@microsoft/sp-http';

export class GraphPersonaService implements IGraphPersonaService {
    constructor(private graphClient: MSGraphClient) {
    }

    public getProfileInfo(): Promise<User> {
        return this.graphClient
        .api('me')
        .get();
    }
    public getPhoto(): Promise<string> {
        return this.graphClient
        .api('/me/photo/$value')
        .responseType('blob')
        .get()
        .then(res => {
            return window.URL.createObjectURL(res);
        });
    }
}