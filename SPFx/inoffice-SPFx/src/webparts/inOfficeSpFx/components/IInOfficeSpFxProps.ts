
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { DisplayMode } from "@microsoft/sp-core-library";
import { panelMode } from "../../../spservices/IEnumPanel";
import spservices from "../../../spservices/spservices";
import {IOfficeRecord} from "../../interfaces/IOfficeRecord";

export interface IInOfficeSpFxProps {
  context: WebPartContext;
  pnanelMode: panelMode;
  officeRecord: IOfficeRecord;
  spservices: spservices;
  onDismissPanel?: (refresh:boolean) => void;
  ShowPanel: boolean;
}
