import { MSGraphClient, HttpClient } from '@microsoft/sp-http';

export interface ICreateO365GroupsProps {
  graphClient: MSGraphClient;
  httpClient: HttpClient;
  userEmail: string;
  context: any;
}