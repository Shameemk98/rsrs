import { createTheme } from "@fluentui/react";

export const BrandColors = {
    // Darkest Base
    primaryDark:    "#005A9E", // Pure corporate dark blue (From gradient 0%)
    themeDarkAlt:   "#006CBF", // Smooth step between Dark and Primary
  
    // The Anchor
    primary:        "#0078D4", // Standardized core SharePoint Blue (From gradient 40%)
  
    // Mid-Light Variants
    primaryLight:   "#1194DC", // Vibrant medium cyan-blue (From gradient 70%)
    primaryLighter: "#2BB3F3", // Bright sky blue highlight (From gradient 100%)
  
    // Soft UI Background Tints
    themeLight:     "#CFEFFF", // Clean alert/pill background tint
    themeLighter:   "#EAF8FF", // Ultra-soft page/container wash tint  

    white: "#FFFFFF",

    neutralLighter: "#f3f9fd",

    border: "#E1DFDD",

    success: "#107C10",
    warning: "#FFB900",
    error: "#D13438",

    pageBackground: "#F5F7FA"
};

export const RsrsTheme = createTheme({
  palette: {
    themePrimary: BrandColors.primary,
    themeDark: BrandColors.primaryDark,
    themeDarkAlt: BrandColors.themeDarkAlt,

    white: BrandColors.white,

    neutralLighter: BrandColors.neutralLighter,
  }
});