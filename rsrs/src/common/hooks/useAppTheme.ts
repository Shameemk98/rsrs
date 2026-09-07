import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { useSPTheme } from './useSPTheme';
import { getTokens } from '../theme/getTokens';

export const useAppTheme = (context: WebPartContext) => {
  const theme = useSPTheme(context);

  const tokens = React.useMemo(() => {
    return getTokens(theme);
  }, [theme]);

  return { theme, tokens };
};
