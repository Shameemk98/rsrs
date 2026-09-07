import { ITextFieldStyles } from '@fluentui/react';
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
  