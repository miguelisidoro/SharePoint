import * as React from 'react';
import styles from './App.module.scss';
import { IAppProps } from './IAppProps';
import { DetailsListCustomers } from './customer/list/DetailsListCustomers';
import { escape } from '@microsoft/sp-lodash-subset';

export default class App extends React.Component<IAppProps, {}> {
  public render(): React.ReactElement<IAppProps> {
    return (
      <DetailsListCustomers />
    );
  }
}
