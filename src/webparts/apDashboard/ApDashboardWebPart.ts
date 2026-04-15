import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { ApDashboard } from './components/ApDashboard';

export default class ApDashboardWebPart extends BaseClientSideWebPart<{}> {

  public render(): void {
    const element = React.createElement(ApDashboard, {
      context: this.context
    });
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}
