import LocalizedStrings, { LocalizedStringsMethods } from "react-localization";

import { enStrings } from "./en-us";
import { ptStrings } from "./pt-pt";

export interface IStrings extends LocalizedStringsMethods {
  PropertyPaneDescription: string;
  PropertiesGroupName: string;
  SharePointRelativeListUrlFieldLabel: string;
  NumberOfItemsToShowLabel: string;
  NumberOfDaysToRetrieveLabel: string;
  InformationTypeLabel: string;
  BirthdaysInformationTypeLabel: string;
  WorkAnniversariesInformationTypeLabel: string;
  NewHiresInformationTypeLabel: string;
  TodayLabel: string;
  WebPartTitleLabel: string;
  NoBirthdaysLabel: string;
  NoWorkAnniversariesLabel: string;
  NoNewHiresLabel: string;
  WebPartConfigurationMissing: string;
}

export const localizedStrings: IStrings = new LocalizedStrings({
  en: enStrings,
  pt: ptStrings
});