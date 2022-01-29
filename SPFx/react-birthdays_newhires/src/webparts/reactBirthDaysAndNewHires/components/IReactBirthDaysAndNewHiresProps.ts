import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IReactBirthDaysAndNewHiresProps {
  personalInformationListUrl: string;
  numberOfItemsToShow: number;
  birthDayNumberOfUpcomingDays: number;
  context: WebPartContext;
}
