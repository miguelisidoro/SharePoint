import { InformationType } from "../../../enums";
import { PersonaInformation, UserInformation } from "../../../models";

export interface IBirthdaysWorkAnniversariesNewCollaboratorsMoreResultsState {
    users: UserInformation[];
    informationType: InformationType;
  }