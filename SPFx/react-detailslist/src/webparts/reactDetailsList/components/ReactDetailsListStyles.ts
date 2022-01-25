import { FontSizes, FontWeights, DefaultPalette, mergeStyleSets } from '@fluentui/react/lib/Styling';

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