import * as React from 'react';
import { IMyTeamsGraphProps } from './IMyTeamsGraphProps';
import {
  CommandBar,
  DefaultButton,
  Dialog,
  DialogFooter,
  DialogType,
  IContextualMenuProps,
  IconButton,
  ImageFit,
  Label,
  Link,
  MessageBar,
  MessageBarType,
  Persona,
  PersonaSize,
  PrimaryButton,
  SearchBox,
  Separator,
  ShimmeredDetailsList,
  Spinner,
  SpinnerSize,
  Stack,
  getTheme,
  Dropdown,
  IDropdownStyles,
  IDropdownOption,
  themeRulesStandardCreator,
  IPersonaProps
} from '@fluentui/react';
import { PersonInformation } from '../../../data/PersonInformation';
import { GraphServiceProvider, SharePointServiceProvider } from '../../../api';
import { Microsoft365Group } from '../../../data';
import { IMyTeamsGraphState } from './IMyTeamsGraphPropsState';
import { DropDownItemMapper, PersonInformationMapper } from '../../../mapper';
import { Microsoft365GroupUser } from '../../../data/Microsoft365GroupUser';
import { LivePersona } from "@pnp/spfx-controls-react/lib/controls/LivePersona";
import { ServiceScope } from '@microsoft/sp-core-library';
import styles from './MyTeamsGraph.module.scss';

const dropdownStyles: Partial<IDropdownStyles> = {
  dropdown: { width: '300px', marginBottom: 10 },
  dropdownOptionText: { overflow: 'visible', whiteSpace: 'normal' },
  dropdownItem: { height: 'auto' },
};

const personaProps: IPersonaProps = {
  size: PersonaSize.size48,
  styles: {
    root: {
      width: 350,
      margin: 5,
    },
  },
};

export default class MyTeamsGraph extends React.Component<IMyTeamsGraphProps, IMyTeamsGraphState> {

  private _graphServiceProvider: GraphServiceProvider;
  private _serviceScope: ServiceScope;

  constructor(props) {
    super(props);

    this._serviceScope = null;

    this.state = ({
      microsoft365Groups: null,
      microsoftGroupOptions: null,
      selectedGroupMembers: null,
    });

    this._graphServiceProvider = new GraphServiceProvider(this.props.context);

    this.loadMicrosoft365Groups = this.loadMicrosoft365Groups.bind(this);
    this.onGroupChange = this.onGroupChange.bind(this);
  }

  public async componentDidMount(): Promise<void> {
    await this.loadMicrosoft365Groups();

    let groupOptions: any[] = DropDownItemMapper.mapToDropDownItems(this.state.microsoft365Groups);

    this.setState({
      microsoftGroupOptions: groupOptions,
    });
  }

  private async loadMicrosoft365Groups() {
    console.log("componentDidMount...");

    let microsoft365Groups: Microsoft365Group[] = await this._graphServiceProvider.getMicrosoft365Groups();

    this.setState({
      microsoft365Groups: microsoft365Groups,
    });
  }

  private onGroupChange = async (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): Promise<void> => {
    debugger;
    const groupMembers: Microsoft365GroupUser[] = await this._graphServiceProvider.getMicrosoft365GroupMembers(item.key.toString());

    const groupMembersPersonInformation = PersonInformationMapper.mapToPersonInformations(groupMembers);

    this.setState({ selectedGroupMembers: groupMembersPersonInformation });
  };

  public render(): React.ReactElement<IMyTeamsGraphProps> {

    if (this.state.microsoftGroupOptions !== null) {
      return (
        <div>
          <Label>Select Team</Label>
          <Dropdown styles={dropdownStyles} options={this.state.microsoftGroupOptions} placeholder="Team" onChange={this.onGroupChange} />
          <Label>Team Members</Label>
          {
            <div className={styles.myTeamsContainer}>
              {
                this.state.selectedGroupMembers !== null && this.state.selectedGroupMembers.map(groupMember =>
                    <Stack horizontal wrap>
                      <LivePersona serviceScope={this._serviceScope} upn={groupMember.userPrincipalName}
                        template={
                          <>
                            <Persona {...groupMember} {...personaProps} coinSize={48} />
                          </>
                        }
                        context={this.props.context}
                      />
                    </Stack>
                )
              }
            </div>
          }
        </div>
      );
    }
    else {
      return (<div></div>);
    }
  }
}

