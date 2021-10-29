import * as React from 'react';
import styles from './InOfficeSpFx.module.scss';
import { IInOfficeSpFxProps } from './IInOfficeSpFxProps';
import { escape } from '@microsoft/sp-lodash-subset';

const LOG_SOURCE: string = "In Office SPFx";
const IMAGE_LIST_NO_ITEMS: string =
  "https://static2.sharepointonline.com/files/fabric/office-ui-fabric-react-assets/images/emptyfolder/empty_list.svg";

import {
  CommandBar,
  DefaultButton,
  DetailsList,
  DetailsListLayoutMode,
  Dialog,
  DialogFooter,
  DialogType,
  FontIcon,
  IColumn,
  IContextualMenuProps,
  Icon,
  IconButton,
  IconType,
  Label,
  MessageBar,
  MessageBarType,
  Persona,
  PersonaSize,
  PrimaryButton,
  SearchBox,
  Selection,
  SelectionMode,
  Separator,
  ShimmeredDetailsList,
  Spinner,
  SpinnerSize,
  Stack,
  Text,
  TextField,
  getTheme,
  mergeStyles
} from "office-ui-fabric-react";
import { FontSizes, FontWeights } from "@uifabric/styling";
import InfiniteScroll from "react-infinite-scroller";

export default class InOfficeSpFx extends React.Component<
IInOfficeSpFxProps,
IInOfficeSpFxState> {
  public render(): React.ReactElement<IInOfficeSpFxProps> {
    return (
      // <div className={ styles.inOfficeSpFx }>
      //   <div className={ styles.container }>
      //     <div className={ styles.row }>
      //       <div className={ styles.column }>
      //         <span className={ styles.title }>Welcome to SharePoint!</span>
      //         <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
      //         <p className={ styles.description }>{escape(this.props.description)}</p>
      //         <a href="https://aka.ms/spfx" className={ styles.button }>
      //           <span className={ styles.label }>Learn more</span>
      //         </a>
      //       </div>
      //     </div>
      //   </div>
      // </div>

      <div >
            <InfiniteScroll
              pageStart={0}
              threshold={50}
              //loadMore={this._getPedidosCompraNextPage}
              //hasMore={this.state.hasMore}
              useWindow={false}//funciona
            >
              <ShimmeredDetailsList
                items={this.state.items}
                compact={false}
                columns={this.state.columns}
                selectionMode={SelectionMode.single}
                setKey="items"
                layoutMode={DetailsListLayoutMode.justified}
                isHeaderVisible={true}
                enableShimmer={this.state.isLoading}
                enableUpdateAnimations={false}
                selection={this._selection}
              />
            </InfiniteScroll>
            {this.state.items.length == 0 && !this.state.isLoading && (
              <Stack tokens={{ childrenGap: 0 }}>
                <Stack.Item align="center">
                  <Icon
                    className={styles.noListItemsImageStyle}
                    iconType={IconType.Image}
                    imageProps={{
                      src: IMAGE_LIST_NO_ITEMS
                    }}
                  />
                </Stack.Item>
                <Stack.Item align="center">
                  <Label className={styles.title}>
                    {strings.NoItemsListMessage}
                  </Label>
                </Stack.Item>
              </Stack>
            )}
          </div>
    );
  }
}
