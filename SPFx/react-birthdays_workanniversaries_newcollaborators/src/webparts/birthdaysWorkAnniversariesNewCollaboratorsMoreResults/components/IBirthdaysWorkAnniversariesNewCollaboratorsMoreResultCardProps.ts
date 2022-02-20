import { DisplayMode } from "@microsoft/sp-core-library";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IBirthdaysWorkAnniversariesNewCollaboratorsMoreResultCardProps {
  sharePointRelativeListUrl: string;
  numberOfDaysToRetrieveForBirthdays: number;
  numberOfDaysToRetrieveForWorkAnniveraries: number;
  numberOfDaysToRetrieveForNewCollaborators: number;
  context: WebPartContext;
  title: string;
  displayMode: DisplayMode;
  updateTitleProperty: (value: string) => void;
}
