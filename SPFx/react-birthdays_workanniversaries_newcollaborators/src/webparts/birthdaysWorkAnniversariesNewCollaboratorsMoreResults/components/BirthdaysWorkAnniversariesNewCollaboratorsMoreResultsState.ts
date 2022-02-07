import { InformationType } from "../../../enums";
import { PersonaInformation, UserInformation } from "../../../models";

export interface BirthdaysWorkAnniversariesNewCollaboratorsMoreResultsState {
    allUsers: UserInformation[];
    pagedUsers: UserInformation[];
    usersToShow: PersonaInformation[];
    informationType: InformationType;
  }