import * as React from 'react';
import styles from './ReactBirthDaysAndNewHires.module.scss';
import { IReactBirthDaysAndNewHiresProps } from './IReactBirthDaysAndNewHiresProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class ReactBirthDaysAndNewHires extends React.Component<IReactBirthDaysAndNewHiresProps, {}> {
  public render(): React.ReactElement<IReactBirthDaysAndNewHiresProps> {
    return (
      <div className={ styles.reactBirthDaysAndNewHires }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <p className={ styles.description }>{escape(this.props.personalInformationListUrl)}</p>
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
