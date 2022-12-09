import { InformationType } from "@app/enums";
import { UserInformation } from "@app/models";

export interface IBirthdaysWorkAnniversariesNewCollaboratorsMoreResultsState {
    users: UserInformation[];
    nextPageUrl: string;
    informationType: InformationType;
  }