import { SPHttpClient } from '@microsoft/sp-http';

export interface IApartment {
  id: number;
  title: string;
  apartmentNumber: number;
  floor: number;
  status: string;
}

export interface IHandoverManagmentProps {
  siteUrl: string;
  listName: string;
  apartmentsPerFloor: number;
  spHttpClient: SPHttpClient;
  statusColors: { [status: string]: string };
}
