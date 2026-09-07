export const getTokens = (theme: any) => {
  const p = theme.palette;

  return {
    colors: {
      background1: p.themePrimary,
      background2: p.neutralLighterAlt,

      foregroundPrimary: p.neutralPrimary,
      foregroundSecondary: p.neutralSecondary,

      brandForeground: p.themePrimary,

      borderPrimary: p.neutralLight,
      borderSecondary: p.neutralLight,

      brandBorder: p.themePrimary
    },

    spacingVerticalXS: 4,
    spacingVerticalS: 8,
    spacingVerticalM: 12,
    spacingVerticalL: 16,
    spacingVerticalXL: 20,
  
    spacingHorizontalS: 8,
    spacingHorizontalM: 12,
    spacingHorizontalL: 16,
    spacingHorizontalXL: 20,
  
    radius: {
      medium: 4,
      large: 6
    },
  
    shadows: {
      shadow4: '0 1px 3px rgba(0,0,0,0.04)',
      shadow16: '0 -2px 8px rgba(0,0,0,0.12)'
    },
  
    typography: {
      fontSize200: 12,
      fontSize300: 14
    }
  };
};
