import { WebPartContext } from "@microsoft/sp-webpart-base";
import { DisplayMode } from "@microsoft/sp-core-library";
import { panelMode } from "../../../../spservices/IEnumPanel";

export interface IInOfficeSpFxProps {
  title: string;
  context: WebPartContext;
  displayMode:DisplayMode;
  updateProperty(value:string):void;
}

