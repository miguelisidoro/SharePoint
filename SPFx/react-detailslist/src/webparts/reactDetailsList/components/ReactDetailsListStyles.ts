import { IStackStyles } from '@fluentui/react';
import { FontSizes, FontWeights, DefaultPalette, mergeStyleSets } from '@fluentui/react/lib/Styling';

export const stackStyles: IStackStyles = {
  root: {
    alignItems: 'flex-start',
    margin: 0,
    width: '100%'
  }
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
          visibility: 'hidden',
        }
      }
    }
  });