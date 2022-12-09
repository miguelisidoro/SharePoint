import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IContact, panelMode } from '../../../models';

export interface IReactDetailsItemPanelProps {
  mode: panelMode;
  showPanel: boolean;
  Contact: IContact;
  onDismiss(ev?: React.SyntheticEvent<HTMLElement>): void;
  context: WebPartContext;
  readOnly?: boolean;
  refreshData(ev, contact: IContact): void;
}
