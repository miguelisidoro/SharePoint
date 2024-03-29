import { DisplayMode } from "@microsoft/sp-core-library";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IBirthdaysWorkAnniversariesNewCollaboratorsProps {
  sharePointRelativeListUrl: string;
  showMoreUrl: string;
  informationType: string;
  numberOfItemsToShow: number;
  numberOfDaysToRetrieve: number;
  context: WebPartContext;
  title: string;
  displayMode: DisplayMode;
  updateTitleProperty: (value: string) => void;
}
