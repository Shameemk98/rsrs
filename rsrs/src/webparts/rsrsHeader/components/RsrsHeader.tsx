import * as React from 'react';
import { useEffect, useState } from 'react';

import {
  Stack,
  Text,
  Persona,
  PersonaSize,
  IconButton,
  ThemeProvider
} from '@fluentui/react';

import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

import { IRsrsHeaderProps } from './IRsrsHeaderProps';
import '../styles/RSRSHeader.css';

import { RsrsTheme } from '../../../common/theme/RsrsTheme';
import { SpService } from '../../../services/spService';
import { getRsrsHeaderStyles } from './RsrsHeaderStyles';
import { getTokens } from '../../../common/theme/getTokens';

const RSRSHeader: React.FC<IRsrsHeaderProps> = (props) => {
  const tokens = getTokens(RsrsTheme);

  const styles = React.useMemo(() => {
    return getRsrsHeaderStyles(tokens);
  }, []);

  const [quarter, setQuarter] = useState('');
  const [userPhoto, setUserPhoto] = useState('');
  const [siteTitle, setSiteTitle] = useState('');
  const [siteLogo, setSiteLogo] = useState('');

  const user = props.context.pageContext.user;

  const spService = React.useMemo(
    () => new SpService(props.context),
    [props.context]
  );

  const getCurrentQuarter = (): string => {
    const today = new Date();
    const month = today.getMonth() + 1;
    const year = today.getFullYear();

    let quarter = '';

    if (month <= 3) {
      quarter = 'Q1';
    } else if (month <= 6) {
      quarter = 'Q2';
    } else if (month <= 9) {
      quarter = 'Q3';
    } else {
      quarter = 'Q4';
    }

    return `${year} ${quarter}`;
  };
  const loadUserPhoto = (): void => {
    const photoUrl =
      `${props.context.pageContext.web.absoluteUrl}` +
      `/_layouts/15/userphoto.aspx?size=L&accountname=${user.email}`;

    setUserPhoto(photoUrl);
  };

  useEffect(() => {
    const loadHeader = async (): Promise<void> => {
      setQuarter(getCurrentQuarter());
      loadUserPhoto();

      const siteInfo = await spService.getSiteInfo();

      setSiteTitle(siteInfo.title);
      setSiteLogo(siteInfo.logoUrl);
    };

    void loadHeader();
  }, [spService]);


  return (
    <>
      <ThemeProvider theme={RsrsTheme}>
        <header className={styles.root}>
          <div className={styles.outer}>
            <div className={styles.inner}>
              <div className={styles.left}>
                <div className={styles.brand}>
                  {siteLogo && (
                    <a
                      href={props.context.pageContext.web.absoluteUrl}
                      aria-label="Home"
                    >
                      <img
                        src={siteLogo}
                        alt="Site logo"
                        className={styles.logo}
                      />
                    </a>
                  )}</div>
                <div className={styles.titleWrapper} >
                  <div>
                    <Text className={styles.title}>{siteTitle}</Text>
                  </div>
                  <div>
                    <Text className={styles.quarterText}>
                      Update Quarter: {quarter}
                    </Text>
                  </div>
                </div>

              </div>

              {/* RIGHT: User */}
              <div className={styles.right}>
                <Persona
                  imageUrl={userPhoto}
                  text={user.displayName}
                  secondaryText=""
                  size={PersonaSize.size48}
                  styles={{
                    primaryText: {
                      color: 'white !important',
                      fontWeight: '600'
                    },
                    secondaryText: {
                      color: '#d9e8ff !important'
                    }
                  }}
                />

                <IconButton
                  iconProps={{ iconName: 'Settings' }}
                  title="Settings"
                  className={styles.headerIcon}
                />

              </div>

            </div>
          </div>
        </header>
      </ThemeProvider>
    </>
  );
};

export default RSRSHeader;