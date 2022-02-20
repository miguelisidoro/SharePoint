import { InformationType } from "../../../enums";
import { PersonaInformation, UserInformation } from "../../../models";

export interface IBirthdaysWorkAnniversariesNewCollaboratorsMoreResultsState {
    allUsers: UserInformation[];
    pagedUsers: UserInformation[];
    usersToShow: PersonaInformation[];
    informationType: InformationType;
  }