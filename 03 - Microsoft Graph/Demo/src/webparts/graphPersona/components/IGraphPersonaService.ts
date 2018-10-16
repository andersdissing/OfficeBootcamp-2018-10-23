import { User } from '@microsoft/microsoft-graph-types';
export interface IGraphPersonaService {
    getProfileInfo():Promise<User>;
    getPhoto():Promise<string>;
}