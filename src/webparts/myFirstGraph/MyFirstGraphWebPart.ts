import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './MyFirstGraphWebPart.module.scss';
import * as strings from 'MyFirstGraphWebPartStrings';
import { MSGraphClientV3 } from '@microsoft/sp-http';
//import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
//import { spfi, SPFx as spSPFx} from "@pnp/sp"; 
//import { graphfi, SPFx as graphSPFx, SPFxToken } from "@pnp/graph";
//import "@pnp/graph/groups";
//import "@pnp/graph/members";
//import "@pnp/graph/users";

export interface IMyFirstGraphWebPartProps {
  description: string;
}

export default class MyFirstGraphWebPart extends BaseClientSideWebPart<IMyFirstGraphWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public async render(): Promise<void> {
    //const sp = spfi().using(spSPFx(this.context));
  
        this.domElement.innerHTML = `
        <section class="${styles.myFirstGraph} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
          <div class="${styles.welcome}">
            <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
            <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
            <div>${this._environmentMessage}</div>
            <div>Web part property value: <strong>${escape(this.properties.description)}</strong></div>
          </div>
          <div class="${styles.myFirstGraph}">
          <div id="spListContainer" />
          </div>  
          <div>
            <h3>Welcome to SharePoint Framework!</h3>
            <p>
            The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It's the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
            </p>
            <h4>Learn more about SPFx development:</h4>
              <ul class="${styles.links}">
                <li><a href="https://aka.ms/spfx" target="_blank">SharePoint Framework Overview</a></li>
                <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank">Use Microsoft Graph in your solution</a></li>
                <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank">Build for Microsoft Teams using SharePoint Framework</a></li>
                <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
                <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank">Publish SharePoint Framework applications to the marketplace</a></li>
                <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank">SharePoint Framework API reference</a></li>
                <li><a href="https://aka.ms/m365pnp" target="_blank">Microsoft 365 Developer Community</a></li>
              </ul>
          </div>
        </section>`;
        // List the latest emails based on what we got from the Graph
        let mgrflag = await this.getPermissions();
        if(mgrflag){
          alert(mgrflag);
        }
  }

  private async getPermissions(): Promise<boolean> {
    let mgrFlag = false;
    let html : string = "";
    let graphClient: MSGraphClientV3 = (await this.context.msGraphClientFactory.getClient("3"));//.api("/me").get();        
    //const graph = graphfi().using(graphSPFx(this.context), SPFxToken(this.context));

    try {

      graphClient = await this.context.msGraphClientFactory.getClient("3");
      const members : any = await graphClient.api("/groups/84278918-192b-4cae-a4a9-dc962f035575/members").top(999).get();
      const myDetails = await graphClient.api("/me").get();
      
      console.log("group members",members);
      console.log((await myDetails).id);
      console.log("total members",members.value.length);

      for (let x = 0; x < members.value.length; x++) {
        console.log("member id",members.value[x].id);
        let groupMemberID = members.value[x].id;
        if (groupMemberID === (await myDetails).id) {
          mgrFlag = true;
          html = "<p>User is a Manager</p>";
          break;
        }
      }

      const listContainer: Element | null = this.domElement.querySelector('#spListContainer');
      if(listContainer){listContainer.innerHTML = html;}
      
    } catch (err) {
      //this.addError(this.properties.siteName,"checkManagerPermission",err.message);
    }
    return mgrFlag;
  }

  public onInit(): Promise<void> {

    return super.onInit().then(_ => {
      this._getEnvironmentMessage().then(message => {
        this._environmentMessage = message;
      });

    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
