import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IBirthdaysWorkAnniverariesNewHiresProps {
  sharePointRelativeListUrl: string;
  informationType: string;
  numberOfItemsToShow: number;
  numberOfDaysToRetrieve: number;
  context: WebPartContext;
}
