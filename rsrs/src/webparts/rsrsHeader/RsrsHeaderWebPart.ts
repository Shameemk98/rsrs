import * as React from 'react';
import * as ReactDom from 'react-dom';

import {
  BaseClientSideWebPart
} from '@microsoft/sp-webpart-base';

import RSRSHeader from './components/RsrsHeader';

export default class RsrsHeaderWebPart
  extends BaseClientSideWebPart<{}> {

  public render(): void {

    const element = React.createElement(
      RSRSHeader,
      {
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }
}