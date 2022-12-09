
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { DisplayMode } from "@microsoft/sp-core-library";
import { panelMode } from "../../../../spservices/IEnumPanel";
import spservices from "../../../../spservices/spservices";
import {IInOfficeAppointment} from "../../../../interfaces/IInOfficeAppointment";

export interface IInOfficeSpFxProps {
  context: WebPartContext;
  panelMode: panelMode;
  officeAppointment: IInOfficeAppointment;
  spservices: spservices;
  onDismissPanel?: (refresh:boolean) => void;
  ShowPanel: boolean;
}

