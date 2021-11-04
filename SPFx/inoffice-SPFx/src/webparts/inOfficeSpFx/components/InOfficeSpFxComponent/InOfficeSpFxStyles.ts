import { DefaultPalette, FontSizes, FontWeights, mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';
import {
  IButtonStyles,
  IDatePickerStyles,
  IDropdownStyles,
  IModalStyles,
  IStackItemStyles,
  IStackStyles,
  IStackTokens,
  IStyle,
  ITextFieldProps,
  ITextFieldStyles,
  ITextFieldSubComponentStyles,
  ImageLoadState,
  calculatePrecision,
} from 'office-ui-fabric-react';

//import { CommunicationColors } from '@uifabric/fluent-theme/lib/fluent/FluentColors';

// Styles definition
export const stackStyles: IStackStyles = {
  root: {

    marginTop: 10
  }
};
export const stackItemStyles: IStackItemStyles = {
  root: {
    padding: 5,
    display: 'flex',
    width: 172,
    height: 32,
    fontWeight: FontWeights.regular,
  }
};

export const stackTokens: IStackTokens = {
  childrenGap: 10,

};

export const textFielStartDateDatePickerStyles: ITextFieldProps = {
  styles: {
    field: { backgroundColor: DefaultPalette.neutralLighter },
    root: {},
    wrapper: {},
    subComponentStyles: undefined
  }

};

export const textFielDueDateDatePickerStyles: ITextFieldProps = {
  styles: {
    field: { backgroundColor: DefaultPalette.neutralLighter },
    root: {},
    wrapper: {},
    subComponentStyles: undefined
  }

};

// export const textFieldDescriptionStyles: ITextFieldStyles = {
//   field: { backgroundColor: DefaultPalette.neutralLighter },
//   root: {},
//   description: {},
//   errorMessage: {},
//   fieldGroup: {},
//   icon: {},
//   prefix: {},
//   suffix: {},
//   wrapper: {},
//   subComponentStyles: undefined
// };



export const dropDownBucketStyles: IDropdownStyles = {

  root: { margin: 0 } ,
  title: {backgroundColor: '#f4f4f4', borderWidth:0},
  callout: {},
  caretDown: {},
  caretDownWrapper: {},
  dropdown:{},
  dropdownDivider: {},
  dropdownItem: {},
  dropdownItemDisabled: {},
  dropdownItemHeader:{},
  dropdownItemHidden: {},
  dropdownItemSelected:{},
  dropdownItemSelectedAndDisabled:{},
  dropdownItems:{},
  dropdownItemsWrapper:{},
  dropdownOptionText:{},
  errorMessage:{},
  label:{},
  panel:{},
  subComponentStyles: undefined,
};


export const commandBarActionStyles :  IButtonStyles = {
  root: {
    backgroundColor: 'rgb(244, 244, 244)'
  },
  rootHovered: { backgroundColor: 'rgb(234, 234, 234)' },
  rootDisabled: { backgroundColor: 'rgb(244, 244, 244)'},
};

export const classNames = mergeStyleSets({
  centerColumn: { display: 'flex', alignItems:'center', height:'100%'},
  fileIconHeaderIcon: {
    padding: 0,
    fontSize: '16px'
  },
  fileIconCell: {
    textAlign: 'center',
    selectors: {
      '&:before': {
        content: '.',
        display: 'inline-block',
        verticalAlign: 'middle',
        height: '100%',
        width: '0px',
        visibility: 'hidden'
      }
    }
  }
});
