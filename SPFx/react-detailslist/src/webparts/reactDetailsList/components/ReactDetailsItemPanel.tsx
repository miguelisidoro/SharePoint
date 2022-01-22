import * as React from 'react';

import {
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
    Panel,
    Persona,
    PersonaSize,
    PrimaryButton,
    SearchBox,
    Separator,
    ShimmeredDetailsList,
    Spinner,
    SpinnerSize,
    Stack,
    getTheme
  } from '@fluentui/react';
import { IReactDetailsItemPanelState } from './IReactDetailsItemPanelState';
import { IContact, panelMode } from '../../../models';
import * as strings from 'ReactDetailsListWebPartStrings';
import SharePointServiceProvider from '../../../api/SharePointServiceProvider';
import { IReactDetailsItemPanelProps } from './IReactDetailsItemPanelProps';

  export default class ReactDetailsItemPanel extends React.Component<IReactDetailsItemPanelProps, IReactDetailsItemPanelState> {
    
    private sharePointServiceProvider: SharePointServiceProvider;

    constructor(props) {
      super(props);

      this.sharePointServiceProvider = new SharePointServiceProvider(this.props.context);
    }

    public async componentDidMount(): Promise<void> {
    }

    private _onRenderFooterContent = (): JSX.Element => {
      return (
        <div>
          <Separator></Separator>
          <PrimaryButton text={this.props.mode === panelMode.Edit ? strings.PrimaryButtonLabelSave : strings.PrimaryButtonLabelInsert} disabled={this.props.readOnly} onClick={this.onSave} style={{ marginRight: "8px" }}>
            {/* {this.state.isAddingPedidoCompra ? <Spinner size={SpinnerSize.small} /> : 'Criar'} */}
          </PrimaryButton>
          <DefaultButton text={strings.PrimaryButtonLabelCancel} onClick={this.onCancel} />
        </div>
      );
    }

    // Cancel Panel
    private onCancel(ev: React.MouseEvent<HTMLButtonElement>) {
      ev.preventDefault();
      this.props.onDismiss();
    }

    private validateForm(): boolean {
      let contact: IContact = this.state.Contact;
      if (contact == null)
        throw "Erro ao processar o pedido";
      if (contact.Email == null)
        throw "Invalid Email";
      if (contact.MobileNumber == null)
      throw "Invalid Mobile Number";
      if (contact.Name == null)
      throw "Invalid Name";
     
      return true;
    }

    private async onSave(ev: React.MouseEvent<HTMLButtonElement>) {
      ev.preventDefault();
      try {
        if (this.validateForm()) {
          switch (this.props.mode) {
            // add immobilized request
            case (panelMode.New):
              try {
                await this.sharePointServiceProvider.addContact(this.state.Contact);
                this.props.addItemToList(ev, this.state.Contact);
                this.setState({ showPanel: false });
                // this.setState({
                //   showPanelConfirmation: true,
                // });
              } catch (error) {
                this.setState({ errorMessage: error });
              }
              break;
            //edit immobilized request
            case (panelMode.Edit):
              try {
                await this.sharePointServiceProvider.updateContact(this.state.Contact);
                this.setState({ showPanel: false });
              } catch (error) {
                this.setState({ errorMessage: error });
              }
              break;
            default:
              break;
          }
        }
        //TODO ERRROR
        else {
  
        }
      } catch (e) {
        this.setState({
          errorMessage: e
        });
      }
    }

    // // Render
    // public render(): React.ReactElement<IReactDetailsItemPanelProps> {
    //   const currentDate = format(parseISO(new Date().toISOString()), 'P', { locale: pt });

    //   const requiredTag = <label style={{ color: "rgb(168, 0, 0)" }}>*</label>;
    //   return (
    //     <div>
    //       <Panel
    //         closeButtonAriaLabel="Fechar"
    //         isOpen={this.state.showPanel}
    //         type={PanelType.medium}
    //         // customWidth={'800px'}
    //         onDismiss={this.props.onDismiss}
    //         isFooterAtBottom={true}
    //         headerText={this.props.mode == panelMode.Edit ? (this.props.readOnly ? strings.PanelHeaderTextVisualize : strings.PanelHeaderTextEdit) : strings.PanelHeaderTextAdd}
    //         onRenderFooterContent={this._onRenderFooterContent}
    //         onRenderNavigationContent={this._onRenderNavigationContent} >
    //         {
    //           this.state.errorMessage && (
    //             <MessageBar messageBarType={MessageBarType.error}>{this.state.errorMessage}</MessageBar>
    //           )
    //         }
    //         {
    //           this.state.isLoading ? (
    //             <Stack disableShrink styles={tsStyles.stackStyles}   >
    //               <Stack.Item align="stretch">
    //                 <Spinner size={SpinnerSize.large} />
    //               </Stack.Item>
    //             </Stack>
    //           ) : (
    //               <div>
    //                 <Stack styles={{ root: { width: "100%", marginTop: 5 } }}>
    //                   <Stack horizontal horizontalAlign="start" gap="15" >
    //                     <Stack.Item grow={2} >
    //                       {/*  company */}
    //                       <div>
    //                         <label style={{ marginTop: 5, marginBottom: 5, display: "block" }} >{strings.ListViewColumnCompanyLabel}  {requiredTag} </label>
    //                         <ListItemPicker listId={this.state.ImmobilizedListsIds !== null ? this.state.ImmobilizedListsIds.companyId : null}
    //                           disabled={this.props.readOnly}
    //                           defaultSelectedItems={this.state.FullImmobilizedRequest && this.state.FullImmobilizedRequest.company ?
    //                             [{ key: this.state.FullImmobilizedRequest.company.description, name: `${this.state.FullImmobilizedRequest.company.code} ${this.state.FullImmobilizedRequest.company.description}` }] : null}
                              
    //                           columnInternalName={strings.ImmobilizedSPListPropertyDescriptionItem}
    //                           concatColumns={true}
    //                           itemLimit={1}
    //                           keyColumnInternalName={strings.ImmobilizedSPListPropertyCompany}
    //                           filterList={strings.filterCompany}
    //                           onSelectedItem={this.setCompany}
    //                           context={this.props.context}

    //                         />
    //                         {/* {this.state.FullImmobilizedRequest && this.state.FullImmobilizedRequest.company && strings.ImmobilizedSPListPropertyDescriptionItem &&
    //                           <label className={tsStyles.classNames.labelStyles} >{this.state.FullImmobilizedRequest.company.description} </label>
    //                         } */}
    //                       </div>
    //                     </Stack.Item>
    //                     <Stack.Item grow={2} >

    //                     </Stack.Item>
    //                   </Stack>
    //                   <Stack horizontal horizontalAlign="start" gap="10">
    //                     <Stack.Item grow={2}>
    //                       {/* denomination */}
    //                       <TextField
    //                         label={strings.ListViewColumnDenominationLabel}
    //                         readOnly={this.props.readOnly}
    //                         required={true}
    //                         style={this.props.readOnly ? { background: "#d3d3d3" } : {}}
    //                         deferredValidationTime={1500}
    //                         defaultValue={this.state.FullImmobilizedRequest ? this.state.FullImmobilizedRequest.denomination : ''}
    //                         onGetErrorMessage={(value) => {
    //                           this.setState({
    //                             FullImmobilizedRequest: {
    //                               ...this.state.FullImmobilizedRequest,
    //                               denomination: value
    //                             }
    //                           });
    //                           return '';
    //                         }} />

    //                     </Stack.Item>
    //                   </Stack>
    //                   <Stack horizontal horizontalAlign="start" gap="10">
    //                     <Stack.Item grow={2}>
    //                       {/* immobilizedClass */}
    //                       <div>
    //                         <label style={{ marginTop: 5, marginBottom: 5, display: "block" }}>{strings.ListViewColumnImmobilizedClassLabel} {requiredTag} </label>
    //                         <ListItemPicker listId={this.state.ImmobilizedListsIds !== null ? this.state.ImmobilizedListsIds.immobilizedClassId : null}
    //                           disabled={this.props.readOnly}
    //                           defaultSelectedItems={this.state.FullImmobilizedRequest && this.state.FullImmobilizedRequest.immobilizedClass ?
    //                             [{ key: this.state.FullImmobilizedRequest.immobilizedClass.code, name: `${this.state.FullImmobilizedRequest.immobilizedClass.code} ${this.state.FullImmobilizedRequest.immobilizedClass.description}` }] : null}
                              
    //                           columnInternalName={strings.ImmobilizedSPListPropertyDescriptionItem}
    //                           concatColumns = {true}
    //                           itemLimit={1}
    //                           keyColumnInternalName={strings.ImmobilizedSPListPropertyImmobilizedClass}
    //                           onSelectedItem={this.setImmobilizedClass}
    //                           context={this.props.context}
    //                           filterList={this.state.FullImmobilizedRequest !== null ? this.state.FullImmobilizedRequest.company != null ?
    //                             "Empresa eq '" + this.state.FullImmobilizedRequest.company.code + "'" : "" : ""} />

    //                         {/* {this.state.FullImmobilizedRequest && this.state.FullImmobilizedRequest.immobilizedClass &&
    //                           <label className={tsStyles.classNames.labelStyles} >{this.state.FullImmobilizedRequest.immobilizedClass.description} </label>
    //                         } */}
    //                       </div>
    //                     </Stack.Item>
    //                     <Stack.Item grow={2}>
    //                       {/* quantity */}
    //                       <label style={{ marginTop: 5, marginBottom: 5, display: "block" }}>{strings.ListViewColumnQuantityLabel} {requiredTag}</label>
    //                       <div className={tsStyles.classNames.Spin}>
    //                         <input
    //                           style={{ "MozAppearance": "textfield", "WebkitAppearance": "textfield" }}
    //                           disabled={this.props.readOnly}
    //                           onChange={(value) => {
    //                             this.setState({
    //                               FullImmobilizedRequest: {
    //                                 ...this.state.FullImmobilizedRequest,
    //                                 quantity: value.currentTarget.value
    //                               }
    //                             });
    //                             return '';
    //                           }}
    //                           value={this.state.FullImmobilizedRequest ? this.state.FullImmobilizedRequest.quantity : ''}
    //                           type="number"
    //                           lang="pt-PT"
    //                           step={1}
    //                           min={0}
    //                           required
    //                           className={tsStyles.classNames.SpinButton}></input>
    //                       </div>
    //                     </Stack.Item>
    //                   </Stack>

    //                   <Stack horizontal horizontalAlign="start" gap="10">
    //                     {/* pep element */}
    //                     {
    //                       this.state.FullImmobilizedRequest != null && this.state.FullImmobilizedRequest.immobilizedClass != null && this.state.FullImmobilizedRequest.immobilizedClass.pep &&
    //                       <Stack.Item grow={2}>
    //                         <div>
    //                           <label style={{ marginTop: 5, marginBottom: 5, display: "block" }}>{strings.ListViewColumnPepElementLabel}{requiredTag} </label>
    //                           <ListItemPicker listId={this.state.ImmobilizedListsIds !== null ? this.state.ImmobilizedListsIds.pepElementId : null}
    //                             disabled={this.props.readOnly}
    //                             defaultSelectedItems={this.state.FullImmobilizedRequest && this.state.FullImmobilizedRequest.pepElement ?
    //                               [{ key: this.state.FullImmobilizedRequest.pepElement.code, name: `${this.state.FullImmobilizedRequest.pepElement.code} ${this.state.FullImmobilizedRequest.pepElement.description}`}] : null}
    //                             columnInternalName={strings.ImmobilizedSPListPropertyDescriptionItem}
    //                             concatColumns = {true}
    //                             itemLimit={1}
    //                             keyColumnInternalName={strings.ImmobilizedSPListPropertyPepElement}
    //                             onSelectedItem={this.setPepElement}
    //                             filterList={this.state.FullImmobilizedRequest !== null ? this.state.FullImmobilizedRequest.company != null ?
    //                               "Empresa eq '" + this.state.FullImmobilizedRequest.company.code + "'" : "" : ""}
    //                             context={this.props.context} />

    //                           {this.state.FullImmobilizedRequest && this.state.FullImmobilizedRequest.pepElement &&
    //                             <label className={tsStyles.classNames.labelStyles} >{this.state.FullImmobilizedRequest.pepElement.description} </label>
    //                           }
    //                         </div>
    //                       </Stack.Item>
    //                     }

    //                     {/* class8 */}
    //                     {
    //                       this.state.FullImmobilizedRequest != null && this.state.FullImmobilizedRequest.immobilizedClass != null && this.state.FullImmobilizedRequest.immobilizedClass.class8 &&
    //                       <Stack.Item grow={2}>
    //                         <div>
    //                           <label style={{ marginTop: 5, marginBottom: 5, display: "block" }}>{strings.ListViewColumnClassification8Label}  {requiredTag} </label>
    //                           <ListItemPicker listId={this.state.ImmobilizedListsIds !== null ? this.state.ImmobilizedListsIds.class8Id : null}
    //                             disabled={this.props.readOnly}
    //                             defaultSelectedItems={this.state.FullImmobilizedRequest && this.state.FullImmobilizedRequest.class8 ?
    //                               [{ key: this.state.FullImmobilizedRequest.class8.code, name: this.state.FullImmobilizedRequest.class8.description }] : null}
    //                             columnInternalName={strings.ImmobilizedSPListPropertyDescriptionItem}
    //                             itemLimit={1}
    //                             keyColumnInternalName={strings.ImmobilizedSPListPropertyClass8}
    //                             onSelectedItem={this.setClass8}
    //                             context={this.props.context} />
    //                         </div>
    //                       </Stack.Item>
    //                     }

    //                   </Stack>
    //                   <Stack horizontal horizontalAlign="start" gap="10">
    //                     <Stack.Item grow={2} >
    //                       {/* <div style={{ display: "block" }}> */}
                          
    //                       {/* createSubImmobilized */}
    //                       <div style={{ marginTop: 40, marginBottom: 20 }}>
    //                         <Checkbox label="Criar subn.º"
    //                           checked={this.state.FullImmobilizedRequest ? this.state.FullImmobilizedRequest.createSubImmobilized : false}
    //                           disabled={this.props.readOnly}

    //                           onChange={(ev, value) => this.setState({
    //                             FullImmobilizedRequest: {
    //                               ...this.state.FullImmobilizedRequest,
    //                               createSubImmobilized: value,
    //                               subImmobilized: !this.state.FullImmobilizedRequest.createSubImmobilized ? null : this.state.FullImmobilizedRequest.subImmobilized
    //                             }
    //                           })} />
    //                       </div>
    //                     </Stack.Item>

    //                     {/* subImmobilized */}
    //                     {
    //                       this.state.FullImmobilizedRequest != null && this.state.FullImmobilizedRequest.createSubImmobilized &&
    //                       <Stack.Item grow={2}>
    //                       {/* subImmobilized */}
    //                       <label style={{ marginTop: 5, marginBottom: 5, display: "block", marginLeft: 80 }}>{strings.ListViewColumnSubImmobilizedLabel} {requiredTag}</label>
    //                       <div className={tsStyles.classNames.Spin} style={{marginLeft: 80}}>
    //                         <input
    //                           style={{ "MozAppearance": "textfield", "WebkitAppearance": "textfield" }}
    //                           disabled={this.props.readOnly}
    //                           onChange={(value) => {
    //                             var SubImmobilized: ISubImmobilized = {
    //                               code: value.currentTarget.value,
                                  
    //                             };
    //                             this.setState({
    //                               FullImmobilizedRequest: {
    //                                 ...this.state.FullImmobilizedRequest,
    //                                 subImmobilized: SubImmobilized
    //                               }
    //                             });
    //                             return '';
    //                           }}
    //                           value={this.state.FullImmobilizedRequest && this.state.FullImmobilizedRequest.subImmobilized ? this.state.FullImmobilizedRequest.subImmobilized.code : ''}
    //                           type="number"
    //                           lang="pt-PT"
    //                           step={1}
    //                           min={0}
    //                           required
    //                           className={tsStyles.classNames.SpinButton}></input>
    //                       </div>
    //                     </Stack.Item>
    //                     }
                      

    //                   </Stack>
    //                   <Stack horizontal horizontalAlign="start" gap="10">
    //                     <Stack.Item grow={2}>
    //                       {/* costcenter */}
    //                       <div>
    //                         <label style={{ marginTop: 5, marginBottom: 5, display: "block" }}>{strings.ListViewColumnCostCenterLabel}{requiredTag} </label>
    //                         <ListItemPicker listId={this.state.ImmobilizedListsIds !== null ? this.state.ImmobilizedListsIds.costCenterId : null}
    //                           disabled={this.props.readOnly}
    //                           defaultSelectedItems={this.state.FullImmobilizedRequest && this.state.FullImmobilizedRequest.costCenter ?
    //                             [{ key: this.state.FullImmobilizedRequest.costCenter.code, name: `${this.state.FullImmobilizedRequest.costCenter.code} ${this.state.FullImmobilizedRequest.costCenter.description}` }] : null}
                              
    //                           columnInternalName={strings.ImmobilizedSPListPropertyDescriptionItem}
    //                           concatColumns={true}

    //                           itemLimit={1}
    //                           keyColumnInternalName={strings.ImmobilizedSPListPropertyCostCenter}
    //                           onSelectedItem={this.setCostCenter}
    //                           filterList={this.state.FullImmobilizedRequest !== null ? this.state.FullImmobilizedRequest.company != null ?
    //                             "Empresa eq '" + this.state.FullImmobilizedRequest.company.code + "'" : "" : ""}
    //                           context={this.props.context} />
    //                         {/* {this.state.FullImmobilizedRequest && this.state.FullImmobilizedRequest.costCenter &&
    //                           <label className={tsStyles.classNames.labelStyles}>{this.state.FullImmobilizedRequest.costCenter.description} </label>
    //                         } */}
    //                       </div>
    //                     </Stack.Item>
    //                     <Stack.Item grow={2}>
    //                       <div>
    //                         <TextField
    //                           label={strings.responsableLabel}
    //                           readOnly={true}
    //                           style={{ background: "#d3d3d3" }}
    //                           defaultValue={this.state.FullImmobilizedRequest && this.state.FullImmobilizedRequest.costCenter ? this.state.FullImmobilizedRequest.costCenter.responsable : ''}
    //                         />
    //                       </div>
    //                     </Stack.Item>
    //                   </Stack>

    //                   <Stack horizontal horizontalAlign="start" gap="10">
    //                     <Stack.Item grow={2}>
    //                       {/* center */}
    //                       <div>
    //                         <label style={{ marginTop: 5, marginBottom: 5, display: "block" }}>{strings.ListViewColumnCenterLabel}</label>
    //                         <ListItemPicker listId={this.state.ImmobilizedListsIds !== null ? this.state.ImmobilizedListsIds.centerId : null}
    //                           disabled={this.props.readOnly}

    //                           defaultSelectedItems={this.state.FullImmobilizedRequest && this.state.FullImmobilizedRequest.center && this.state.FullImmobilizedRequest.center.code?
    //                             [{ key: this.state.FullImmobilizedRequest.center.code, name: `${this.state.FullImmobilizedRequest.center.code} ${this.state.FullImmobilizedRequest.center.description}` }] 
    //                               : this._centroDefaultSelectedItems}
                              
    //                           columnInternalName={strings.ImmobilizedSPListPropertyDescriptionItem}
    //                           itemLimit={1}
    //                           keyColumnInternalName={strings.ImmobilizedSPListPropertyCenter}
    //                           onSelectedItem={this.setCenter}
    //                           concatColumns = {true}
    //                           context={this.props.context} />
    //                         {/* {this.state.FullImmobilizedRequest && this.state.FullImmobilizedRequest.center &&
    //                           <label className={tsStyles.classNames.labelStyles} >{this.state.FullImmobilizedRequest.center.description} </label>
    //                         } */}
    //                       </div>
    //                     </Stack.Item>

    //                     <Stack.Item grow={2}>
    //                       {/* localization */}
    //                       <div>
    //                       {
    //                       this.state.FullImmobilizedRequest != null && this.state.FullImmobilizedRequest.immobilizedClass != null && this.state.FullImmobilizedRequest.immobilizedClass.mandatoryLocation &&
    //                         <label style={{ marginTop: 5, marginBottom: 5, display: "block" }}>{strings.ListViewColumnLocalizationLabel}{requiredTag}</label>}
    //                       {
    //                         this.state.FullImmobilizedRequest != null && this.state.FullImmobilizedRequest.immobilizedClass != null && !this.state.FullImmobilizedRequest.immobilizedClass.mandatoryLocation &&
                          
    //                         <label style={{ marginTop: 5, marginBottom: 5, display: "block" }}>{strings.ListViewColumnLocalizationLabel}</label>}
    //                         <ListItemPicker listId={this.state.ImmobilizedListsIds !== null ? this.state.ImmobilizedListsIds.localizationId : null}
    //                           disabled={this.props.readOnly}
    //                           defaultSelectedItems={this.state.FullImmobilizedRequest && this.state.FullImmobilizedRequest.localization && this.state.FullImmobilizedRequest.localization.code?
    //                             [{ key: this.state.FullImmobilizedRequest.localization.code, name: `${this.state.FullImmobilizedRequest.localization.code} ${this.state.FullImmobilizedRequest.localization.description}` }] : null}
                              
    //                           columnInternalName={strings.ImmobilizedSPListPropertyDescriptionItem}
    //                           concatColumns={true}

    //                           itemLimit={1}
    //                           keyColumnInternalName={strings.ImmobilizedSPListPropertyLocalization}
    //                           onSelectedItem={this.setLocalization}
    //                           filterList={this.state.FullImmobilizedRequest !== null ? this.state.FullImmobilizedRequest.center != null ?
    //                             "Centro eq '" + this.state.FullImmobilizedRequest.center.code + "'" : "" : ""}
    //                           context={this.props.context} />
                            
    //                         {/* {this.state.FullImmobilizedRequest && this.state.FullImmobilizedRequest.localization &&
    //                           <label style={{ marginTop: 5, marginBottom: 5, display: "block", color: 'Blue' }}>{this.state.FullImmobilizedRequest.localization.description} </label>
    //                         } */}
    //                       </div>
    //                     </Stack.Item>
    //                   </Stack>

    //                   <Stack horizontal horizontalAlign="start" gap="10">
    //                     {this.state.FullImmobilizedRequest != null && this.state.FullImmobilizedRequest.immobilizedClass != null && this.state.FullImmobilizedRequest.immobilizedClass.internalOrder &&

    //                       <Stack.Item grow={2}>
    //                         {/* internal order */}
    //                         <div>
    //                           <label style={{ marginTop: 5, marginBottom: 5, display: "block" }}> {strings.ListViewColumnInternalOrderLabel}</label>
    //                           <ListItemPicker listId={this.state.ImmobilizedListsIds !== null ? this.state.ImmobilizedListsIds.internalOrderId : null}
    //                             disabled={this.props.readOnly}
    //                             defaultSelectedItems={this.state.FullImmobilizedRequest && this.state.FullImmobilizedRequest.internalOrder && this.state.FullImmobilizedRequest.internalOrder.code?
    //                               [{ key: this.state.FullImmobilizedRequest.internalOrder.code, name: `${this.state.FullImmobilizedRequest.internalOrder.code} ${this.state.FullImmobilizedRequest.internalOrder.description}` }] : null}
                                
    //                             columnInternalName={strings.ImmobilizedSPListPropertyDescriptionItem}
    //                             concatColumns={true}

    //                             itemLimit={1}
    //                             keyColumnInternalName={strings.ImmobilizedSPListPropertyInternalOrder}
    //                             onSelectedItem={this.setInternalOrder}
    //                             context={this.props.context} />

    //                           {/* {this.state.FullImmobilizedRequest && this.state.FullImmobilizedRequest.internalOrder &&
    //                             <label className={tsStyles.classNames.labelStyles} >{this.state.FullImmobilizedRequest.internalOrder.description} </label>
    //                           } */}

    //                         </div>
    //                       </Stack.Item>
    //                     }
    //                     <Stack.Item grow={2}>
    //                       {/* licenseplate */}
    //                       <TextField
    //                         label={strings.ListViewColumnLicensePlateLabel}
    //                         readOnly={this.props.readOnly}
    //                         required={false}
    //                         style={this.props.readOnly ? { background: "#d3d3d3" } : {}}
    //                         defaultValue={this.state.FullImmobilizedRequest ? this.state.FullImmobilizedRequest.licensePlate : ''}
    //                         deferredValidationTime={1500}
    //                         onGetErrorMessage={(value) => {
    //                           this.setState({
    //                             FullImmobilizedRequest: {
    //                               ...this.state.FullImmobilizedRequest,
    //                               licensePlate: value
    //                             }
    //                           });
    //                           return '';
    //                         }} />
    //                     </Stack.Item>
    //                   </Stack>

    //                   <Stack horizontal horizontalAlign="start" gap="10">
    //                     <Stack.Item grow={2}>
    //                       <div>
    //                         {/* dgci code */}
    //                         <label style={{ marginTop: 5, marginBottom: 5, display: "block" }}>{strings.ListViewColumnDGCICodeLabel} {requiredTag} </label>
    //                         <ListItemPicker listId={this.state.ImmobilizedListsIds !== null ? this.state.ImmobilizedListsIds.dgciCodeId : null}
    //                           disabled={this.props.readOnly}
                              
    //                           defaultSelectedItems={this.state.FullImmobilizedRequest && this.state.FullImmobilizedRequest.dgciCode ?
    //                             [{ key: this.state.FullImmobilizedRequest.dgciCode.code, name: `${this.state.FullImmobilizedRequest.dgciCode.code} ${this.state.FullImmobilizedRequest.dgciCode.description}` }] : null}
                              
    //                           columnInternalName={strings.ImmobilizedSPListPropertyDescriptionItem}
    //                           concatColumns={true}

    //                           itemLimit={1}
    //                           keyColumnInternalName={strings.ImmobilizedSPListPropertyDgciCode}
    //                           onSelectedItem={this.setDGCICode}
    //                           filterList={this.dgciFilter()}
    //                           context={this.props.context}
    //                         />
    //                         {this.state.FullImmobilizedRequest && this.state.FullImmobilizedRequest.dgciCode &&
    //                           <label className={tsStyles.classNames.labelStyles} >{this.state.FullImmobilizedRequest.dgciCode.description} </label>
    //                         }
    //                       </div>
    //                     </Stack.Item>

    //                     <Stack.Item grow={2}>
    //                       <div>
    //                         {/* investiment motive */}
    //                         <label style={{ marginTop: 5, marginBottom: 5, display: "block" }}>{strings.ListViewColumnInvestimentMotiveLabel}</label>
    //                         <ListItemPicker listId={this.state.ImmobilizedListsIds !== null ? this.state.ImmobilizedListsIds.investimentMotiveId : null}
    //                           disabled={this.props.readOnly}
    //                           defaultSelectedItems={this.state.FullImmobilizedRequest && this.state.FullImmobilizedRequest.investimentMotive && this.state.FullImmobilizedRequest.investimentMotive.code ?
    //                             [{ key: this.state.FullImmobilizedRequest.investimentMotive.code, name: this.state.FullImmobilizedRequest.investimentMotive.description }] : null}
    //                           columnInternalName={strings.ImmobilizedSPListPropertyDescriptionItem}
    //                           itemLimit={1}
    //                           keyColumnInternalName={strings.ImmobilizedSPListPropertyInvestimentMotive}
    //                           onSelectedItem={this.setInvestimentMotive}
    //                           context={this.props.context} />
    //                       </div>
    //                     </Stack.Item>
    //                   </Stack>

    //                   <Stack horizontal horizontalAlign="start" gap="10">
    //                     <Stack.Item grow={2}>
    //                       <div>
    //                         {/* state of good  */}
    //                         <label style={{ marginTop: 5, marginBottom: 5, display: "block" }}>{strings.ListViewColumnRequestStateOfGood}</label>
    //                         <ListItemPicker listId={this.state.ImmobilizedListsIds !== null ? this.state.ImmobilizedListsIds.stateOfGoodId : null}
    //                           disabled={this.props.readOnly}
    //                           defaultSelectedItems={this.state.FullImmobilizedRequest && this.state.FullImmobilizedRequest.stateOfGood && this.state.FullImmobilizedRequest.stateOfGood.code?
    //                             [{ key: this.state.FullImmobilizedRequest.stateOfGood.code, name: this.state.FullImmobilizedRequest.stateOfGood.description }] : null}
    //                           columnInternalName={strings.ImmobilizedSPListPropertyDescriptionItem}
    //                           itemLimit={1}
    //                           keyColumnInternalName={strings.ImmobilizedSPListPropertyStateOfGood}
    //                           onSelectedItem={this.setStateOfGood}
    //                           context={this.props.context} />
    //                       </div>
    //                     </Stack.Item>

    //                     <Stack.Item grow={2}>
    //                       <div>
    //                         {/* partner company */}
    //                         <label style={{ marginTop: 5, marginBottom: 5, display: "block" }}>{strings.ListViewColumnPartnerCompanyLabel}</label>
    //                         <ListItemPicker listId={this.state.ImmobilizedListsIds !== null ? this.state.ImmobilizedListsIds.companyId : null}
    //                           disabled={this.props.readOnly}
    //                           defaultSelectedItems={this.state.FullImmobilizedRequest && this.state.FullImmobilizedRequest.partnerCompany && this.state.FullImmobilizedRequest.partnerCompany.code?
    //                             [{ key: this.state.FullImmobilizedRequest.partnerCompany.code, name: this.state.FullImmobilizedRequest.partnerCompany.description }] : null}
    //                           columnInternalName={strings.ImmobilizedSPListPropertyDescriptionItem}
    //                           itemLimit={1}
    //                           keyColumnInternalName={strings.ImmobilizedSPListPropertyCompany}
    //                           onSelectedItem={this.setPartnerCompany}
    //                           filterList={this.state.FullImmobilizedRequest !== null ? this.state.FullImmobilizedRequest.company != null ?
    //                             "Empresa ne '" + this.state.FullImmobilizedRequest.company.code + "'" : "" : ""}
    //                           context={this.props.context} />
    //                       </div>
    //                     </Stack.Item>

    //                   </Stack>
    //                   {/* lifespan */}
    //                   <Stack horizontal horizontalAlign="start" gap="10">
    //                     <Stack.Item grow={2}>
    //                       {/* lifespan */}
    //                       <div>                  <label style={{ marginTop: 5, marginBottom: 5, display: "block" }}>{strings.ListViewColumnLifespanLabel} {requiredTag}</label>
    //                         <div className={tsStyles.classNames.Spin}>
    //                           <input
    //                             style={{ "MozAppearance": "textfield", "WebkitAppearance": "textfield" }}
    //                             disabled={this.props.readOnly}
    //                             onChange={(value) => {
    //                               this.setState({
    //                                 FullImmobilizedRequest: {
    //                                   ...this.state.FullImmobilizedRequest,
    //                                   lifespan: value.currentTarget.value
    //                                 }
    //                               });
    //                               return '';
    //                             }}
    //                             min={0}
    //                             max={1000000}
    //                             value={this.state.FullImmobilizedRequest ? this.state.FullImmobilizedRequest.lifespan : ''}
    //                             type="number"
    //                             lang="pt-PT"
    //                             required
    //                             className={tsStyles.classNames.SpinButton}></input>
    //                         </div>
    //                       </div>
    //                     </Stack.Item>
    //                   </Stack>


    //                   {/* numbers attributed */}
    //                   {/* <TextField
    //             label={strings.ListViewColumnNumbersAttributedLabel}
    //             readOnly={false}
    //             required={false}
    //             multiline={true}
    //             defaultValue={this.state.FullImmobilizedRequest ? this.state.FullImmobilizedRequest.numbersAttributed : ''}
    //             deferredValidationTime={1500}
    //             onGetErrorMessage={(value) => {
    //               this.setState({
    //                 FullImmobilizedRequest: {
    //                   ...this.state.FullImmobilizedRequest,
    //                   numbersAttributed: value
    //                 }
    //               });
    //               return '';
    //             }} /> */}

    //                   {/* solicitant comments */}

    //                   <TextField
    //                     label={strings.ListViewColumnSolicitantCommentsLabel}
    //                     required={false}
    //                     multiline={true}
    //                     readOnly={this.props.readOnly}
    //                     defaultValue={this.state.FullImmobilizedRequest ? this.state.FullImmobilizedRequest.solicitantComments : ''}
    //                     deferredValidationTime={1500}
    //                     style={this.props.readOnly ? { background: "#d3d3d3" } : {}}
    //                     onGetErrorMessage={(value) => {
    //                       this.setState({
    //                         FullImmobilizedRequest: {
    //                           ...this.state.FullImmobilizedRequest,
    //                           solicitantComments: value
    //                         }
    //                       });
    //                       return '';
    //                     }} />
    //                   <label style={{ marginTop: 5, marginBottom: 5, display: "block" }}>(Para as classes de exclusividade e direitos de uso indicar em número de meses, para as restantes em anos)</label>
    //                   {(this.props.mode === panelMode.Edit || this.props.readOnly) &&
    //                     <Stack horizontal horizontalAlign="start" gap="10">
    //                       <Stack.Item grow={2}>
    //                         <TextField
    //                           label={strings.ListViewColumnApprovalCommentsLabel}
    //                           required={false}
    //                           multiline={true}
    //                           readOnly={true}
    //                           defaultValue={this.state.FullImmobilizedRequest ? this.state.FullImmobilizedRequest.approvalComments : ''}
    //                           style={{ background: "#d3d3d3" }}
    //                         />
    //                       </Stack.Item>
    //                     </Stack>
    //                   }
    //                 </Stack>
    //                 {
    //                   this.state.errorMessage && (
    //                     <Stack horizontal horizontalAlign={"start"} styles={tsStyles.stackStyles}>
    //                       <Stack.Item grow={2}>
    //                         <MessageBar messageBarType={MessageBarType.error}>{this.state.errorMessage}</MessageBar>
    //                       </Stack.Item>
    //                     </Stack>
    //                   )
    //                 }
    //                 <Separator></Separator>
    //                 <Stack horizontal horizontalAlign={"start"} styles={tsStyles.stackStyles}>

    //                   <Label style={{ marginBottom: 5 }}>Solicitante</Label>
    //                   {/* <Label>{strings.dateLabel}</Label>
    //           <Label style={{ marginLeft: 10 }}>{currentDate}</Label> */}
    //                 </Stack>
    //                 <Stack horizontal horizontalAlign={"start"} styles={tsStyles.stackStyles}>

    //                   <Persona {...this.state.solicitant} size={PersonaSize.size48} styles={{ root: { margin: 15 } }} />
    //                   <Separator></Separator>
    //                 </Stack>
    //                 <Stack horizontal horizontalAlign={"start"} styles={tsStyles.stackStyles}>
    //                   {!this.props.readOnly &&
    //                     <div>
    //                       <Checkbox label="O adquirente declara ter preenchido a ficha em conformidade  com o Manual de Gestão Administrativa do Imobilizado"
    //                         checked={this.state.permissionCheckbox}
    //                         disabled={this.props.readOnly}
    //                         styles={{label:{maxWidth: 549}}}
    //                         onChange={(ev, value) => this.setState({
    //                           permissionCheckbox: value
    //                         }
    //                         )} />
    //                     </div>}
    //                 </Stack>
    //               </div>
    //             )
    //         }
    //       </Panel >
    //     </div >
    //   );
    // }
  }