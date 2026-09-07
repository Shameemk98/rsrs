import { mergeStyleSets } from '@fluentui/react';
import { BrandColors } from '../../../common/theme/RsrsTheme';

export const getRsrsHeaderStyles = (tokens: any) => {
  return mergeStyleSets({
    root: {
      position: 'relative',
      left: '50%',
      right: '50%',
      marginLeft: '-50vw',
      marginRight: '-50vw',
      width: '100vw',
      zIndex: 100,
      marginTop:'-10px'
    },
    outer: {
     background: `
        linear-gradient(
          120deg,
          #005A9E 0%,
          #0078D4 40%,
          #1194DC 70%,
          #2BB3F3 100%
        )
      `,
      // padding: `${tokens.spacingHorizontalL}px ${tokens.spacingHorizontalXL}px`,
      borderRadius: tokens.radius.lg,
     color: tokens.colors.white,
/*     borderRadius:tokens.borderRadiusLarge,
      color: tokens.colorBrandStroke2, */
      position: 'relative',
      overflow: 'hidden',
      boxShadow: '0 4px 12px rgba(0,0,0,.12)',
      selectors: {
        ':first-child': {},
        ':before': {
            content: '""',
            position: 'absolute',          
            right: '-120px',
            top: '-80px',          
            width: '500px',
            height: '500px',          
            background:
              'radial-gradient(circle, rgba(255,255,255,.12) 0%, transparent 70%)',          
            borderRadius: '50%',          
            pointerEvents: 'none'
          },
          
          ':after': {
            content: '""',
            position: 'absolute',          
            left: '25%',
            bottom: '-120px',          
            width: '700px',
            height: '250px',          
            background:
              'radial-gradient(ellipse, rgba(255,255,255,.08) 0%, transparent 70%)',          
            borderRadius: '50%',          
            pointerEvents: 'none'
          },
          '.topHighlight': {}         
          
      },
      width: '100%',
      padding: tokens.spacingVerticalM,
    },
    // outer: {
    //   width: '100%',
    //   padding: tokens.spacingVerticalM,
    // },
  
    inner: {
      maxWidth: '1100px',
      margin: '0 auto',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'space-between',
      gap: tokens.spacingHorizontalL,
    },
  
    left: {
      display: 'flex',
      alignItems: 'center',
      gap: tokens.spacingHorizontalM,
    },
    right: {
      display: 'flex',
      alignItems: 'center',
      gap: tokens.spacingHorizontalS,
    },
    
    brand: {
      display: 'flex',
      alignItems: 'center',    
      paddingRight: 20,
      marginRight: 20,    
      borderRight: '1px solid rgba(255,255,255,.25)',    
      flexShrink: 0
    },
    siteInfo: {
      display: 'flex',
      flexDirection: 'column',    
      minWidth: 0
    },
    logo: {
      height: 48,
      width: 'auto',
      display: 'block'
    },

    // title: {
    //     color: BrandColors.white,
    //     fontSize: 18,
    //     fontWeight: 600,
    //     lineHeight: 24
    //   },
    title: {
      color: '#fff',    
      fontWeight: 600,
      fontSize: 20,    
      whiteSpace: 'nowrap',
      overflow: 'hidden',
      textOverflow: 'ellipsis',    
      maxWidth: '500px',    
      selectors: {
        '@media (max-width: 768px)': {
          maxWidth: '220px'
        }
      }
    },
    titleWrapper: {
      display: 'flex',
      flexDirection: 'column',
      gap: tokens.spacingVerticalXS,
    },
    quarterText: {
        display: 'inline-flex',
        alignItems: 'center',      
        padding: '6px 14px',      
        borderRadius: 16,      
        background: 'rgba(255,255,255,.14)',      
        border: '1px solid rgba(255,255,255,.12)',      
        backdropFilter: 'blur(8px)',      
        fontSize: 12,
        fontWeight: 600,      
        color: '#fff'
    },
    headerIcon: {
      color: '#ffffff',

      selectors: {
        ':hover': {
          backgroundColor: 'rgba(255,255,255,.15)',
          color: '#ffffff'
        }
      }
    },
    settingsButton: {
      color: '#ffffff',

      selectors: {
        ':hover': {
          backgroundColor: 'rgba(255,255,255,.15)'
        }
      }
    },
    /* personaPrimaryText: {
      color: '#ffffff',
      fontWeight: 600
    },

    personaSecondaryText: {
      color: 'rgba(255,255,255,.85)'
    } */
  });
};