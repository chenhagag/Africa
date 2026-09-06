import { SPHttpClient } from '@microsoft/sp-http';

export interface IApartment {
  id: number;
  title: string;
  apartmentNumber: number;
  floor: number;
  tower: string;
  saleStatus: string;
  handoverStatus: string;
  finishStatus: string;
  apartmentType: string;
  rooms: number;
  standard: string;
  shadowNumber: number;
  isPublicUnit: boolean;
  spaceUnits: number;
  publicUnitType: string;
}

export interface IPublicUnitColors {
  [type: string]: string;
}

export interface IHandoverManagmentProps {
  siteUrl: string;
  listName: string;
  spHttpClient: SPHttpClient;
  saleColors: { [status: string]: string };
  finishColors: { [status: string]: string };
  publicUnitColors: IPublicUnitColors;
}
