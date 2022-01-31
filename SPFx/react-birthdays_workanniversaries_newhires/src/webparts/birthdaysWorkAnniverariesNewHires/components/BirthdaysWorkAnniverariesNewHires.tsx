import * as React from 'react';
import styles from './BirthdaysWorkAnniverariesNewHires.module.scss';
import { IBirthdaysWorkAnniverariesNewHiresProps } from './IBirthdaysWorkAnniverariesNewHiresProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SharePointServiceProvider } from '../../../api';
import { UserInformation } from '../../../models';

export default class BirthdaysWorkAnniverariesNewHires extends React.Component<IBirthdaysWorkAnniverariesNewHiresProps, {}> {
  
  private sharePointServiceProvider: SharePointServiceProvider;

  constructor(props: IBirthdaysWorkAnniverariesNewHiresProps) {
    super(props);

    this.sharePointServiceProvider = new SharePointServiceProvider(this.props.context, 
      this.props.sharePointRelativeListUrl,
      this.props.numberOfItemsToShow,
      this.props.numberOfDaysToRetrieve);

    this.state = {
      users: null,
    };
  }

  public async componentDidMount(): Promise<void> {
    // Populate with items for demos.

    await this.loadUsers();
  }

  private async loadUsers() {
    debugger;
    let users: UserInformation[] = await this.sharePointServiceProvider.getUsers();

    this.setState({ 
      users: users, 
     });
  }

  public render(): React.ReactElement<IBirthdaysWorkAnniverariesNewHiresProps> {
    return (
      <div className={ styles.birthdaysWorkAnniverariesNewHires }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <p className={ styles.description }>{escape(this.props.informationType)}</p>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
