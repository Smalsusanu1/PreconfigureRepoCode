import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'ComponentProfileWebPartStrings';
import ComponentProfile from './components/ComponentProfile';
import { IComponentProfileProps } from './components/IComponentProfileProps';
import { Web } from 'sp-pnp-js';
import { spfi, SPFx } from '@pnp/sp';

export interface IComponentProfileWebPartProps {
  description: string;
  MasterTaskListID: 'ec34b38f-0669-480a-910c-f84e92e58adf';
  TaskUserListID: 'b318ba84-e21d-4876-8851-88b94b9dc300';
  DocumentsListID:'d0f88b8f-d96d-4e12-b612-2706ba40fb08';
  SmartInformationListID:"edf0a6fb-f80e-4772-ab1e-666af03f7ccd";
  SmartMetadataListID: '01a34938-8c7e-4ea6-a003-cee649e8c67a';
  SmartHelpListID:'9cf872fc-afcd-42a5-87c0-aab0c80c5457';
  TaskTypeID:'21b55c7b-5748-483a-905a-62ef663972dc';
  PortFolioTypeID: "c21ab0e4-4984-4ef7-81b5-805efaa3752e";
  TimeEntry:any;
  SiteCompostion:any;
  dropdownvalue:string,
}

export default class ComponentProfileWebPart extends BaseClientSideWebPart<IComponentProfileWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IComponentProfileProps> = React.createElement(
      ComponentProfile,
      
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        MasterTaskListID: this.properties.MasterTaskListID,
        TaskUserListID: this.properties.TaskUserListID,
        TaskTypeID:this.properties.TaskTypeID,
        DocumentsListID:this.properties.DocumentsListID,
        SmartHelpListID:this.properties.SmartHelpListID,
        SmartMetadataListID: this.properties.SmartMetadataListID,
        PortFolioTypeID:this.properties.PortFolioTypeID,
        Context: this.context,
        SmartInformationListID: this.properties.SmartInformationListID,
        TimeEntry:this.properties.TimeEntry,
        SiteCompostion:this.properties.SiteCompostion,
        dropdownvalue:this.properties.dropdownvalue,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    await super.onInit();

    // Setup PnP SP with the SPFx context
    spfi().using(SPFx(this.context)); // Initialize SP context

    // Fetch default values from SharePoint list
    const defaultValues = await this._fetchDefaultValues();
    this.properties.SmartHelpListID= defaultValues.SmartHelpListID || this.properties.SmartHelpListID || '';
    this.properties.MasterTaskListID = defaultValues.MasterTaskListID || this.properties.MasterTaskListID || '';
    this.properties.TaskUserListID = defaultValues.TaskUserListID || this.properties.TaskUserListID || '';
    this.properties.SmartMetadataListID = defaultValues.SmartMetadataListID || this.properties.SmartMetadataListID || '';
    this.properties.SmartInformationListID = defaultValues.SmartInformationListID || this.properties.SmartInformationListID || '';
    this.properties.DocumentsListID = defaultValues.DocumentsListID || this.properties.DocumentsListID || '';
    this.properties.TaskTypeID = defaultValues.TaskTypeID || this.properties.TaskTypeID || '';
    this.properties.PortFolioTypeID = defaultValues.PortFolioTypeID || this.properties.PortFolioTypeID || '';
    this.properties.TimeEntry = defaultValues.TimeEntry || this.properties.TimeEntry || '';
    this.properties.SiteCompostion = defaultValues.SiteCompostion || this.properties.SiteCompostion || '';
    this._environmentMessage = this._getEnvironmentMessage();
    
    return super.onInit();
  }
  private async _fetchDefaultValues(): Promise<any> {
    try {
      const web = new Web(this.context.pageContext.web.absoluteUrl)
        
      const items = await web.lists.getByTitle('SmartMetadata')
      .items.select("Id", "Title", "IsVisible", "Configurations", "SmartSuggestions", "Color_x0020_Tag", "TaxType", "Description1", "Item_x005F_x0020_Cover", "listId", "siteName", "siteUrl", "SortOrder", "SmartFilters", "Selectable", "Parent/Id", "Parent/Title")
      .filter("TaxType eq 'DynamicListId'").expand('Parent').get();
      const item1 = items[0];
      const  item2=JSON.parse(item1?.Configurations)
      const item = item2[0];
      return {
        SmartHelpListID:item.SmartHelpListID || '',
        MasterTaskListID: item.MasterTaskListID || '',
        TaskUserListID: item.TaskUserListID || '',
        SmartMetadataListID: item.SmartMetadataListID || '',
        SmartInformationListID: item.SmartInformationListID || '',
        DocumentsListID: item.DocumentsListID || '',
        TaskTimeSheetListID: item.TaskTimeSheetListID || '',
        TaskTypeID: item.TaskTypeID || '',
        PortFolioTypeID: item.PortFolioTypeID || '',
        TimeEntry: item.TimeEntry || '',
        SiteCompostion: item.SiteCompostion || '',
        SmalsusLeaveCalendar:item.SmalsusLeaveCalendar || '',
        NotificationsConfigrationListID:item.NotificationsConfigrationListID || '',
        ListConfigurationListID:item.ListConfigurationListID || '',
        SitePagesList:item.SitePagesList || '',
        AdminconfigrationID: item.AdminconfigrationID || '',
        TopNavigationListID: item.TopNavigationListID || '',
        TableConfrigrationListId:item.TableConfrigrationListId || '',
        SPBackupConfigListID:item.SPBackupConfigListID || '',
        SkillsPortfolioListID:item.SkillsPortfolioListID || '',
        InterviewFeedbackFormListId:item.InterviewFeedbackFormListId || '',
        TilesManagementListID:item.TilesManagementListID || '', 
        smartMetadata:item.smartMetadata || '', 
        userlistId:item.userlistId || '', 
        ContractListID:item.ContractListID || '',
        HR_EMPLOYEE_DETAILS_LIST_ID:item.HR_EMPLOYEE_DETAILS_LIST_ID || '',
        HHHHContactListId:item.HHHHContactListId || '',
        HHHHInstitutionListId:item.HHHHInstitutionListId || '',
        MAIN_SMARTMETADATA_LISTID:item.MAIN_SMARTMETADATA_LISTID || '',
        MAIN_HR_LISTID:item.MAIN_HR_LISTID || '',
        GMBH_CONTACT_SEARCH_LISTID:item.GMBH_CONTACT_SEARCH_LISTID || '',
        EventListId:item.EventListId || '',
        NewsListId:item.NewsListId || '',
        AnnouncementsListId:item.AnnouncementsListId || '',
        SPSiteConfigListID:item.SPSiteConfigListID || '',
        SPTopNavigationListID:item.SPTopNavigationListID || '',
        TeamContactSearchlistIds:item.TeamContactSearchlistIds || '',
        TeamSmartMetadatalistIds:item.TeamSmartMetadatalistIds || '',
        MasterTaskId:item.MasterTaskId || '',
        UpComingBirthdayId:item.UpComingBirthdayId || '',

        
      };
    } catch (error) {
      console.error("Error fetching default values", error);
      return {};
    }
  }
  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
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

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
                  label: strings.DescriptionFieldLabel,
                  value: this.properties.description
                }),
                PropertyPaneTextField('MasterTaskListID', {
                  label: "MasterTaskListID",
                  value: this.properties.MasterTaskListID
                }),
                PropertyPaneTextField('TaskUserListID', {
                  label: "TaskUserListID",
                  value: this.properties.TaskUserListID
                }),
                PropertyPaneTextField('SmartHelpListID', {
                  label: "SmartHelpListID",
                  value: this.properties.SmartHelpListID
                }),
                PropertyPaneTextField('SmartMetadataListID', {
                  label: "SmartMetadataListID",
                  value: this.properties.SmartMetadataListID
                }),
                PropertyPaneTextField('SmartInformationListID', {
                  label: 'SmartInformationListID',
                  value: this.properties.SmartInformationListID
                }),
                PropertyPaneTextField('DocumentsListID', {
                  label: "DocumentsListID",
                  value: this.properties.DocumentsListID
                }),
                PropertyPaneTextField('TaskTypeID', {
                  label: "TaskTypeID",
                  value: this.properties.TaskTypeID
                }),
                PropertyPaneTextField('PortFolioTypeID', {
                  label: "PortFolioTypeID",
                  value: this.properties.PortFolioTypeID
                }),
                PropertyPaneTextField('TimeEntry', {
                  label: "TimeEntry",
                  value: this.properties.TimeEntry
                }),
                PropertyPaneTextField('SiteCompostion', {
                  label: "SiteCompostion",
                  value: this.properties.SiteCompostion
                })
              ]
        }
      ] 
        }
      ]
    };
  }
}
