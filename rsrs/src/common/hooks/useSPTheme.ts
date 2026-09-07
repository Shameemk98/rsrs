import * as React from 'react';
import {
  ThemeProvider,
  ThemeChangedEventArgs,
  IReadonlyTheme
} from '@microsoft/sp-component-base';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export const useSPTheme = (context: WebPartContext) => {
  const [theme, setTheme] = React.useState<IReadonlyTheme | undefined>();

  React.useEffect(() => {
    const themeProvider = context.serviceScope.consume(
      ThemeProvider.serviceKey
    );

    // Initial theme
    setTheme(themeProvider.tryGetTheme());

    const handleThemeChange = (args: ThemeChangedEventArgs) => {
      setTheme(args.theme);
    };

    themeProvider.themeChangedEvent.add({} as any, handleThemeChange);

    return () => {
      themeProvider.themeChangedEvent.remove({} as any, handleThemeChange);
    };
  }, [context]);

  return theme;
};
