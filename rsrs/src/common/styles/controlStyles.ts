import { IComboBoxStyles, ILabelStyles, ITextFieldStyles } from '@fluentui/react';
import { IDropdownStyles } from '@fluentui/react';

export const getTextFieldStyles = (t: any): Partial<ITextFieldStyles> => ({
  fieldGroup: {
    borderRadius: 2,
    // border: `1px solid ${t.colorNeutralStroke2}`,
    backgroundColor: t.colorNeutralBackground1,
  },
  field: {
    borderRadius: 4,
  },});


  export const getDropdownStyles = (t: any): Partial<IDropdownStyles> => ({
    dropdown: {
      borderRadius: 2,
    },
  
    title: {
      borderRadius: 2,
      // border: `1px solid ${t.colorNeutralStroke1}`,      
      backgroundColor: t.colorNeutralBackground1,
    },
  
    caretDown: {
      color: t.colorNeutralForeground3,
    }
  });
  export const commonDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: "100%"
    },
    dropdown: {
      background: "#f8f8f8",
      border: 'none',//"1px solid rgb(96, 94, 92)",
      borderRadius: 4
    },
    title: {
      background: "#f8f8f8",
      fontSize: 14,
      fontWeight: "400",
      color: "#323130",
      fontFamily:
        '"Segoe UI", SegoeUI, "Helvetica Neue", Helvetica, Arial, sans-serif'
    },
    caretDownWrapper: {
      color: "#605e5c"
    }
    
  
};

  const commonLabelSelector = {
  ".ms-Label": {
    fontSize: "14px",
    fontWeight: 600,
    color: "#323130",
    marginBottom: "6px",
  },
};

export const formStyles = {
  sectionHeader: {
    root: {
      fontSize: 20,
      fontWeight: 600,
      color: "#323130",
      marginTop: 16,
      marginBottom: 20,
      borderBottom: "2px solid #0078d4",
      paddingBottom: 8,
    },
  } as ILabelStyles,

  textField: {
    root: {
      // marginBottom: 16,
      selectors: commonLabelSelector,
    },
    fieldGroup: {
      //minHeight: 40,
      borderRadius: 4,
    },
    field: {
      fontSize: 14,
    },
  } as Partial<ITextFieldStyles>,

  numberField: {
    root: {
      marginBottom: 16,
      selectors: commonLabelSelector,
    },
    fieldGroup: {
      //minHeight: 40,
      borderRadius: 4,
    },
    field: {
      fontSize: 14,
      textAlign: "right",
    },
  } as Partial<ITextFieldStyles>,

  multilineField: {
    root: {
      // marginBottom: 16,
      selectors: commonLabelSelector,
    },
    fieldGroup: {
      borderRadius: 4,
      minHeight:120
    },
    field: {
      fontSize: 14,
      lineHeight: 22,
    },
  } as Partial<ITextFieldStyles>,

  readOnlyField: {
    root: {
      // marginBottom: 16,
      selectors: commonLabelSelector,
    },
    fieldGroup: {
      //minHeight: 40,
      background: "#f8f8f8",
      border: "1px solid #605e5c",
      borderRadius: 4,
    },
    field: {
      fontSize: 14,
      fontWeight: "400",
      color: "#323130",
      fontFamily:
        '"Segoe UI", SegoeUI, "Helvetica Neue", Helvetica, Arial, sans-serif',
    },
  } as Partial<ITextFieldStyles>,
  readOnlyMultiLineField: {
    root: {
      // marginBottom: 16,
      selectors: commonLabelSelector,
    },
    fieldGroup: {
      minHeight: 120,
      background: "#f8f8f8",
      border: "1px solid #605e5c",
      borderRadius: 4,
    },
    field: {
      fontSize: 14,
      fontWeight: "400",
      color: "#323130",
      fontFamily:
        '"Segoe UI", SegoeUI, "Helvetica Neue", Helvetica, Arial, sans-serif',
    },
  } as Partial<ITextFieldStyles>,

  readOnlyNumberField: {
    root: {
      // marginBottom: 16,
      selectors: commonLabelSelector,
    },
    fieldGroup: {
      //minHeight: 40,
      background: "#f8f8f8",
      border: "1px solid #605e5c",
      borderRadius: 4,
    },
    field: {
      fontSize: 14,
      fontWeight: "400",
      color: "#323130",
      textAlign: "right",
      fontFamily:
        '"Segoe UI", SegoeUI, "Helvetica Neue", Helvetica, Arial, sans-serif',
    },
  } as Partial<ITextFieldStyles>,

  dropdown: {
    root: {
      // marginBottom: 16,
      selectors: commonLabelSelector,
    },
    title: {
      //minHeight: 40,
      borderRadius: 4,
    },
  } as Partial<IDropdownStyles>,
  comboBox: {
  root: {
    selectors: commonLabelSelector,
  },
  fieldGroup: {
    borderRadius: 4,
  },
} as Partial<IComboBoxStyles>,

  sectionHeaderStyles: {
  root: {
    fontSize: 24,
    fontWeight: 600,
    marginTop: 12,
    color: "#002050",
    marginBottom: 24,
  },
}as Partial<ILabelStyles>
};

export const customDropdownStyles: Partial<IDropdownStyles> = {
  // 1. Container root styles
  root: {
    // marginBottom: 16,
  },
  // 2. The main visible box (Acts like 'fieldGroup' in TextField)
  title: {
    background: "#f8f8f8",
    border: "1px solid #605e5c",
    borderRadius: 4,
    fontSize: 14,
    fontWeight: "400",
    color: "#323130",
    fontFamily: '"Segoe UI", SegoeUI, "Helvetica Neue", Helvetica, Arial, sans-serif',
    
    selectors: {
      // Force custom background and border when disabled
      ':disabled, .is-disabled &': {
        background: "#f8f8f8",
        border: "1px solid #605e5c",
        color: "#323130", // Keeps text color from fading into light gray
      },
    }
  },
  // 3. The Chevron/Arrow icon
  caretDown: {
    selectors: {
      ':disabled, .is-disabled &': {
        color: "#605e5c", // Keeps the arrow dark instead of fading out
      }
    }
  }
} as Partial<IDropdownStyles>;


export const customComboBoxStyles: Partial<IComboBoxStyles> = {
  // Target the container element directly (.ms-ComboBox)
  root: {
    background: "#f8f8f8",
    borderRadius: 4,
    borderWidth: 1,
    borderStyle: 'solid',
    borderColor: '#605e5c',
    display: 'inline-flex', // Ensures it wraps around input and button properly
    width: '100%',         // Adjust if you want a fixed size
    boxSizing: 'border-box',

    selectors: {
      // 1. Lock down standard focus, hover, and error footprint wrappers
      '&:hover, &:focus, &:focus-within, &.is-open': {
        borderColor: '#605e5c !important',
      },
      // 2. Handle the specific disabled state visible in your DOM (.is-disabled)
      '&.is-disabled, &:disabled': {
        background: "#f8f8f8 !important",
        borderColor: '#605e5c !important',
      },
      // 3. Clear out Fluent's hidden absolute pseudo-element rings
      '::after': {
        border: 'none !important',
        outline: 'none !important'
      }
    }
  },
  // Keep text consistent inside the inner input node (.ms-ComboBox-Input)
  input: {
    fontSize: 14,
    fontWeight: "400",
    color: "#323130 !important",
    fontFamily: '"Segoe UI", SegoeUI, "Helvetica Neue", Helvetica, Arial, sans-serif !important',
    backgroundColor: 'transparent',
    border: 'none !important', // Ensure inner input doesn't have an independent border
    outline: 'none !important',
    
    selectors: {
      '&:disabled, &.is-disabled, .is-disabled &': {
        color: "#323130 !important", 
        WebkitTextFillColor: "#323130 !important", // Fix for Safari/iOS disabled dimming
      },
    }
  },
  // Ensure the caret button matches the container box footprint
  button: {
    backgroundColor: 'transparent',
    border: 'none !important',
    
    selectors: {
      '&:disabled, &.is-disabled, .is-disabled &': {
        color: "#605e5c !important",
        backgroundColor: 'transparent !important',
      }
    }
  }
} as Partial<IComboBoxStyles>;


import { IBasePickerStyles } from '@fluentui/react';

export const customPickerStyles: Partial<IBasePickerStyles> = {
  root: {
    // marginBottom: 16,
  },
  text: {
    // 1. Establish the clean baseline variables explicitly
    background: "#f8f8f8 !important",
    border: "1px solid #605e5c !important", 
    borderRadius: 4,
    minHeight: 32,
    boxSizing: 'border-box',
    
    selectors: {
      // 2. Strict component-specific chaining (using 'ms-BasePicker' scopes it strictly)
      '.ms-BasePicker.is-disabled &, &:disabled, &.is-disabled': {
        background: "#f8f8f8 !important", 
        borderColor: "#605e5c !important", // Hardcoded color isolates it from currentColor leaks
        borderStyle: "solid !important",
        borderWidth: "1px !important",
        outline: "none !important",
      },
      // 3. Prevent Fluent UI from placing an overlay block over your custom frame
      '::after': {
        border: 'none !important',
        content: 'none !important',
      },
      // 4. Scope text rules safely only to this input element node nested inside the picker wrapper
      '& input.ms-BasePicker-input': {
        fontFamily: '"Segoe UI", SegoeUI, "Helvetica Neue", Helvetica, Arial, sans-serif',
        fontSize: 14,
        fontWeight: "400",
        color: "#323130 !important",
        selectors: {
          ':disabled': { 
            color: "#323130 !important" 
          }
        }
      }
    }
  }
};



