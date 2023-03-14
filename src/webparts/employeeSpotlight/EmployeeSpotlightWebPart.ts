import { Version } from '@microsoft/sp-core-library';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './EmployeeSpotlightWebPart.module.scss';
// import * as strings from 'EmployeeSpotlightWebPartStrings';

import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  PropertyPaneSlider
} from '@microsoft/sp-webpart-base';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { SPHttpClient } from '@microsoft/sp-http';
import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';
import { PropertyFieldColorPickerMini } from 'sp-client-custom-fields/lib/PropertyFieldColorPickerMini';
import { PropertyFieldColorPicker } from 'sp-client-custom-fields/lib/PropertyFieldColorPicker';

import * as jQuery from 'jquery';
import * as _ from "lodash";
import * as strings from 'EmployeeSpotlightWebPartStrings';
import { IEmployeeSpotlightWebPartProps } from './IEmployeeSpotlightWebPartProps';
import { SliderHelper } from './Helper';



// export interface IEmployeeSpotlightWebPartProps {
//   description: string;
// }

export interface ResponceDetails {
  title: string;
  id: string;
}

/**
 *  An interface to hold the ResponceDetails collection.
 */
export interface ResponceCollection {
  value: ResponceDetails[];
}

/**
 * An interface to hold the SpotlightDetails.
 */
export interface SpotlightDetails {
  userDisplayName: string;
  userEmail: string;
  userProfilePic: string;
  description: string;
  designation?: string;
  role?: string;
  rewardTitle?: string;
  rewardDescription?: string;
  userImage?:any;
}


export default class EmployeeSpotlightWebPart extends BaseClientSideWebPart<IEmployeeSpotlightWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private spotlightListFieldOptions: IPropertyPaneDropdownOption[] = [];
  private spotlightListOptions: IPropertyPaneDropdownOption[] = [];
  private siteOptions: IPropertyPaneDropdownOption[] = [];
  private defaultProfileImageUrl: string = "/_layouts/15/userphoto.aspx?size=L";
  private helper: SliderHelper = new SliderHelper();
  private sliderControl: any = null;


  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  public constructor() {
    super();
    // alert("Page Load 5");

    SPComponentLoader.loadScript('https://cdnjs.cloudflare.com/ajax/libs/jquery/3.2.1/jquery.min.js', { globalExportsName: 'jQuery' });
    // Next button functionality
    //debugger;
    this.helper.moveSlides(1);
    jQuery(document).on("click", "." + styles.next, (event) => {
      event.preventDefault();    //prevent default action of <a>
      this.helper.moveSlides(1);
    });

    // Previous button functionality
    jQuery(document).on("click", "." + styles.prev, (event) => {
      event.preventDefault();    //prevent default action of <a>
      this.helper.moveSlides(-1);
    });

    // start and stop slider on hover
    jQuery(document).ready(() => {
      jQuery(document).on('mouseenter', '.' + styles.containers, () => {
        if (this.properties.enabledSpotlightAutoPlay)
          clearInterval(this.sliderControl);
      }).on('mouseleave', '.' + styles.containers, () => {
        var carouselSpeed: number = this.properties.spotlightSliderSpeed * 1000;
        if (carouselSpeed && this.properties.enabledSpotlightAutoPlay) {
          //debugger;
          // alert("Auto Play : " +this.helper.startAutoPlay + " ,, carouselSpeed : " + carouselSpeed);
          this.sliderControl = setInterval(this.helper.startAutoPlay, carouselSpeed);
        }

      });
    });
  }

  public render(): void {
    this.domElement.innerHTML = `<div id="spListContainer" />`;
    this._renderSpotlightTemplateAsync();
    this._renderSpotlightDataAsync();
  }

  private _renderSpotlightTemplateAsync(): void {
    debugger;
    if (Environment.type == EnvironmentType.SharePoint || Environment.type == EnvironmentType.ClassicSharePoint) {
      this._getSiteCollectionRootWeb().then((response) => {
        this.properties.spotlightSiteCollectionURL = response['Url'];
      });
      if (this.properties.spotlightSiteURL && this.properties.spotlightListName && this.properties.spotlightEmployeeEmailColumn && this.properties.spotlightDescriptionColumn) {
        let spotlightDataCollection: SpotlightDetails[] = [];
        this._getSpotlightListData(this.properties.spotlightSiteURL, this.properties.spotlightListName, this.properties.spotlightEmployeeExpirationDateColumn, this.properties.spotlightEmployeeEmailColumn, this.properties.spotlightDescriptionColumn,
          this.properties.spotlightRoleColumn,
          this.properties.spotlightRewardTitleColumn,
          this.properties.spotlightRewardDetailsColumn,
          this.properties.spotlightUserImageColumn
          )
          .then((listDataResponse) => {
            console.log(listDataResponse);
            var spotlightListData = listDataResponse.value;
            if (spotlightListData) {
              //debugger;
              for (var key in listDataResponse.value) {
                var email = listDataResponse.value[key][this.properties.spotlightEmployeeEmailColumn]["EMail"];
                var id = listDataResponse.value[key]["ID"];
                this._getUserImage(email)
                  .then((response) => {
                    spotlightListData.forEach((item: ResponceDetails) => {
                      let userSpotlightDetails: SpotlightDetails = { userDisplayName: "", userEmail: "", userProfilePic: "", description: "", role: "", rewardTitle: "", rewardDescription: "" };
                      if (item[this.properties.spotlightEmployeeEmailColumn]["EMail"] == response["Email"]) {
                        var userName = item[this.properties.spotlightEmployeeEmailColumn];
                        var description = item[this.properties.spotlightDescriptionColumn];
                        var role = item[this.properties.spotlightRoleColumn];
                        var rewardTitle = item[this.properties.spotlightRewardTitleColumn];
                        var rewardDescription = item[this.properties.spotlightRewardDetailsColumn];
                        var userImageDetails=item[this.properties.spotlightUserImageColumn];


                        var userDescription = "";
                        var userRewardDescription = "";

                        try {
                          userDescription = jQuery(description).text();
                        }
                        catch (err) {
                          userDescription = description;
                        }

                        try {
                          userRewardDescription = jQuery(rewardDescription).text();
                        }
                        catch (err) {
                          userRewardDescription = rewardDescription;
                        }

                        // if (userDescription.length > 140) {
                        //   var displayFormUrl = this.properties.spotlightSiteURL + '/Lists/' + this.properties.spotlightListName + '/DispForm.aspx?ID=' + id;
                        //   userDescription = userDescription.substring(0, 140) + `&nbsp; <a href="${displayFormUrl}">ReadMore...</a>`;
                        // }
                        var displayName = response["DisplayName"];
                        var designationProperty = _.filter(response["UserProfileProperties"], { Key: "SPS-JobTitle" })[0];
                        var designation = designationProperty["Value"] ? designationProperty["Value"] : "";
                        // uses default image if user image not exist 
                        //debugger;
                        var profilePicture = response["PictureUrl"] != null && response["PictureUrl"] != undefined ? (<string>response["PictureUrl"]).replace("MThumb", "LThumb") : this.defaultProfileImageUrl;
                        // var profilePicture = response["PictureUrl"] != null && response["PictureUrl"] != undefined ? (<string>response["PictureUrl"]) : this.defaultProfileImageUrl;
                        //profilePicture = '/_layouts/15/userphoto.aspx?accountname=' + displayName + '&size=M&url=' + profilePicture.split("?")[0];
                        profilePicture = "/_layouts/15/userphoto.aspx?size=L&username="+response["Email"];

                        if(userImageDetails)
                        {
                          profilePicture=JSON.parse(userImageDetails).serverUrl+JSON.parse(userImageDetails).serverRelativeUrl
                        }

                        userSpotlightDetails = {
                          userDisplayName: response["DisplayName"],
                          userEmail: response["Email"],
                          userProfilePic: profilePicture,
                          description: userDescription,
                          designation: designation,
                          role: role,
                          rewardTitle: rewardTitle,
                          rewardDescription: userRewardDescription,

                        };
                        spotlightDataCollection.push(userSpotlightDetails);
                      }
                    });
                    this._addSpotlightTemplateContent(spotlightDataCollection);
                    if (this.sliderControl == null && this.properties && this.properties.enabledSpotlightAutoPlay) {
                      setTimeout(this.helper.moveSlides, 2000);
                      debugger;
                      // alert("Auto Play : " +this.helper.startAutoPlay + ", carouselSpeed : " + this.properties.spotlightSliderSpeed * 1000);
                      // this.sliderControl = setInterval(this.helper.startAutoPlay, this.properties.spotlightSliderSpeed * 1000);
                      this.sliderControl = setInterval(this.helper.startAutoPlay, 4000);

                    }
                  });
              }
            }
          });
      }
    }
  }

  private _addSpotlightTemplateContent(spotlightDetails: SpotlightDetails[]): void {
    this.domElement.innerHTML = '';
    var innerContent: string = '';
    for (let i: number = 0; i < spotlightDetails.length; i++) {
      innerContent += ` 
                  <div class="${styles.mySlides}">
                    <div style="width:100%; font-family: 'Avenir', sans-serif;">
                          <div style="float:left; display: flex;justify-content: center;margin: auto;height: 250px;padding-top: 10px;">
                            <img style="margin-left: 25px !important;border-radius:2%; height: 200px; width: 200px; margin: auto;" src="${spotlightDetails[i].userProfilePic}" />
                          </div>
                          <div style="width:56%;float:left;text-align:left; padding:10px; height: 250px;">
                              <div style="margin-bottom:0; padding:10px !important; background-color: "[theme: themePrimary, default: #0078d7]"; height: 232px;">
                              <h2 style="margin-bottom:0; margin top : 12px">${spotlightDetails[i].rewardTitle}</h2>
                              <h4 style="margin-top:0">${spotlightDetails[i].rewardDescription}</h4>
                              <h3 style="margin-bottom:0; text-transform: uppercase;">${spotlightDetails[i].userDisplayName}</h3>
                              <h4 style="margin-top:0">${spotlightDetails[i].role}</h4>
                              <p style="height: 125px !important; overflow: hidden;">${spotlightDetails[i].description}</p>                             
                              </div>
                               
                          </div>
                      </div>
                  </div>`;
    }
    this.domElement.innerHTML +=
      `<div class="${styles.containers}" id="slideshow" style="background-color: ${this.properties.spotlightBGColor}; cursor:pointer; width: 100%!important; padding: 5px;border-radius: 15px;box-shadow: rgba(0,0,0,0.25) 0 0 20px 0;text-align:center;color:${this.properties.spotlightFontColor};">
                     ` + innerContent + `
       <a  class="${styles.prev}">&#10094;</a>
       <a  class="${styles.next}">&#10095;</a>
     </div>`;
  }

  private _callAPI(url: string): Promise<ResponceCollection> {
    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1).then((response) => {
      return response.json();
    });
  }
  private _getUserImage(email: string): Promise<ResponceCollection> {
    return this._callAPI(this.properties.spotlightSiteCollectionURL + "/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='i:0%23.f|membership|" + email + "'");
  }

  private _getSiteCollectionRootWeb(): Promise<ResponceCollection> {
    return this._callAPI(this.context.pageContext.web.absoluteUrl + `/_api/Site/RootWeb?$select=Title,Url`);
  }

  private _getAllSubsites(spotlightSiteCollectionURL: string): Promise<ResponceCollection> {
    return this._callAPI(spotlightSiteCollectionURL + `/_api/web/webs?$select=Title,Url`);
  }
  private _getAllLists(siteUrl: string): Promise<ResponceCollection> {
    if (siteUrl != "" && siteUrl != undefined) {
      return this._callAPI(siteUrl + `/_api/web/lists?$orderby=Id desc&$filter=Hidden eq false and BaseTemplate eq 100`);
    }
  }
  private _getSpotlightListFields(siteUrl: string, spotlightListName: string): Promise<ResponceCollection> {
    if (siteUrl != "" && spotlightListName != "" && siteUrl != undefined && spotlightListName != undefined) {
      return this._callAPI(siteUrl + `/_api/web/lists/GetByTitle('${spotlightListName}')/Fields?$orderby=Id desc&$filter=Hidden eq false and ReadOnlyField eq false`);
    }
  }

  private _renderSpotlightDataAsync(): void {
    this._getSiteCollectionRootWeb()
      .then((response) => {
        this.properties.spotlightSiteCollectionURL = response['Url'];
        this._getAllSubsites(response['Url'])
          .then((sitesResponse) => {
            this.siteOptions = this._getDropDownCollection(sitesResponse, 'Url', 'Title');
            this.context.propertyPane.refresh();
            if (this.properties.spotlightSiteURL != "") {
              this._loadAllListsDropDown(this.properties.spotlightSiteURL);
            }
            if (this.properties.spotlightListName != "") {
              this._loadSpotlightListFieldsDropDown(this.properties.spotlightSiteURL, this.properties.spotlightListName);
            }
          });
      });
  }

  private _loadAllListsDropDown(siteUrl: string): void {
    if (siteUrl != "") {
      this._getAllLists(siteUrl)
        .then((response) => {
          this.spotlightListOptions = this._getDropDownCollection(response, 'Title', 'Title');
          this.context.propertyPane.refresh();
        });
    }
  }

  private _loadSpotlightListFieldsDropDown(siteUrl: string, spotlightListName: string): void {
    this._getSpotlightListFields(siteUrl, spotlightListName)
      .then((response) => {
        this.spotlightListFieldOptions = this._getDropDownCollection(response, 'Title', 'Title');
        this.context.propertyPane.refresh();
      });
  }

  private _getDropDownCollection(response: ResponceCollection, key: string, text: string): IPropertyPaneDropdownOption[] {
    var dropdownOptions: IPropertyPaneDropdownOption[] = [];
    if (key == 'Url')
      dropdownOptions.push({ key: this.context.pageContext.web.absoluteUrl, text: 'This Site' });
    for (var itemKey in response.value) {
      dropdownOptions.push({ key: response.value[itemKey][key], text: response.value[itemKey][text] });
    }
    return dropdownOptions;
  }

  private _getSpotlightListData(siteUrl: string, spotlightListName: string, expiryDateColumn: string, emailColumn: string,
    descriptionColumn: string,
    RoleColumn: string,
    RewardTitleColumn: string,
    RewardDetailsColumn: string,
    UserImageColumn:any,
    ): Promise<ResponceCollection> {
    if (siteUrl != "" && spotlightListName != "") {
      var today: Date = new Date();
      var dd: any = today.getDate();
      var mm: any = today.getMonth() + 1; //January is 0!
      var yyyy: any = today.getFullYear();
      dd = (dd < 10) ? '0' + dd : dd;
      mm = (mm < 10) ? '0' + mm : mm;
      var dateString: string = `${yyyy}-${mm}-${dd}`;
      debugger;
      emailColumn = emailColumn.replace(" ", "_x0020_");
      descriptionColumn = descriptionColumn.replace(" ", "_x0020_");
      expiryDateColumn = expiryDateColumn.replace(" ", "_x0020_");
      RoleColumn = RoleColumn.replace(" ", "_x0020_");
      RewardTitleColumn = RewardTitleColumn.replace(" ", "_x0020_");
      RewardDetailsColumn = RewardDetailsColumn.replace(" ", "_x0020_");

      return this._callAPI(siteUrl + `/_api/web/lists/GetByTitle('${spotlightListName}')/items?$select=ID,${descriptionColumn},${RoleColumn},${RewardDetailsColumn},${RewardTitleColumn},${emailColumn}/EMail,${UserImageColumn}&$expand=${emailColumn}/Id&$orderby=Id desc&$filter=${expiryDateColumn} ge '${dateString}'`);
    }
  }
  private _validateFiledValue(value: string): string {
    var validationMessage: string = '';
    if (value === null || value.trim().length === 0) {
      validationMessage = 'Please select a value';
    }
    return validationMessage;
  }
  protected onPropertyPaneConfigurationStart(): void {
    // Stops execution, if the list values already exists
    if (this.spotlightListOptions.length > 0) return;
    // Calls function to append the list names to dropdown
    this._renderSpotlightTemplateAsync();
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    switch (propertyPath) {
      case "spotlightSiteURL":
        this.properties.spotlightListName = "";
        this._loadAllListsDropDown(this.properties.spotlightSiteURL);
        break;
      case "spotlightListName":
        this.properties.spotlightEmployeeEmailColumn = "";
        this.properties.spotlightDescriptionColumn = "";
        this._loadSpotlightListFieldsDropDown(this.properties.spotlightSiteURL, this.properties.spotlightListName);
        break;
      default:
        break;
    }
  }




  // public render(): void 
  // {
  //   this.domElement.innerHTML = `
  //   <section class="${styles.employeeSpotlight} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
  //     <div class="${styles.welcome}">
  //       <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
  //       <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
  //       <div>${this._environmentMessage}</div>
  //       <div>Web part property value: <strong>${escape(this.properties.description)}</strong></div>
  //     </div>
  //     <div>
  //       <h3>Welcome to SharePoint Framework!</h3>
  //       <p>
  //       The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It's the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
  //       </p>
  //       <h4>Learn more about SPFx development:</h4>
  //         <ul class="${styles.links}">
  //           <li><a href="https://aka.ms/spfx" target="_blank">SharePoint Framework Overview</a></li>
  //           <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank">Use Microsoft Graph in your solution</a></li>
  //           <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank">Build for Microsoft Teams using SharePoint Framework</a></li>
  //           <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
  //           <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank">Publish SharePoint Framework applications to the marketplace</a></li>
  //           <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank">SharePoint Framework API reference</a></li>
  //           <li><a href="https://aka.ms/m365pnp" target="_blank">Microsoft 365 Developer Community</a></li>
  //         </ul>
  //     </div>
  //   </section>`;
  // }

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
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);

  }
  // @ts-ignore
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupName: strings.propertyPaneHeading,
              groupFields: [
                PropertyPaneDropdown('spotlightSiteURL', {
                  label: "Site",
                  options: this.siteOptions,
                  selectedKey: this._validateFiledValue.bind(this)
                }),
                PropertyPaneDropdown('spotlightListName', {
                  label: "List",
                  options: this.spotlightListOptions,
                  selectedKey: this._validateFiledValue.bind(this)
                }),
                PropertyPaneDropdown('spotlightEmployeeEmailColumn', {
                  label: "Employee",
                  options: this.spotlightListFieldOptions,
                  selectedKey: this._validateFiledValue.bind(this)
                }),
                PropertyPaneDropdown('spotlightDescriptionColumn', {
                  label: "Description",
                  options: this.spotlightListFieldOptions,
                  selectedKey: this._validateFiledValue.bind(this)
                }),
                PropertyPaneDropdown('spotlightEmployeeExpirationDateColumn', {
                  label: "Expiry Date",
                  options: this.spotlightListFieldOptions,
                  selectedKey: this._validateFiledValue.bind(this)
                }),
                PropertyPaneDropdown('spotlightRoleColumn', {
                  label: "Role",
                  options: this.spotlightListFieldOptions,
                  selectedKey: this._validateFiledValue.bind(this)
                }),
                PropertyPaneDropdown('spotlightRewardTitleColumn', {
                  label: "Reward Title",
                  options: this.spotlightListFieldOptions,
                  selectedKey: this._validateFiledValue.bind(this)
                }),
                PropertyPaneDropdown('spotlightRewardDetailsColumn', {
                  label: "Reward Details",
                  options: this.spotlightListFieldOptions,
                  selectedKey: this._validateFiledValue.bind(this)
                }),
                PropertyPaneDropdown('spotlightUserImageColumn', {
                  label: "Image Details",
                  options: this.spotlightListFieldOptions,
                  selectedKey: this._validateFiledValue.bind(this)
                })


              ]
            },
            {
              groupName: strings.effectsGroupName,
              groupFields: [
                // PropertyFieldColorPickerMini('spotlightBGColor', {
                //   label: strings.spotlightBGColorLableMessage,
                //   initialColor: this.properties.spotlightBGColor,
                //   disabled: false,
                //   onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                //   properties: this.properties,
                //   onGetErrorMessage: null,
                //   deferredValidationTime: 0,
                //   key: 'spotlightBGColorFieldId',
                //   render: function (): void {
                //     throw new Error('Function not implemented.');
                //   }
                // }),
                PropertyFieldColorPickerMini('miniColor', {
                  label: 'Select background color',
                  initialColor: this.properties.spotlightBGColor,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  render: this.render.bind(this),
                  disableReactivePropertyChanges: this.disableReactivePropertyChanges,
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'spotlightBGColorFieldId'
                }),

                //   PropertyFieldColorPicker('spotlightBGColor', {
                //     label: "background color",
                //     initialColor: this.properties.spotlightBGColor,
                //     onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                //     render: this.render.bind(this),
                //     disableReactivePropertyChanges: this.disableReactivePropertyChanges,
                //     properties: this.properties,
                //     onGetErrorMessage: null,
                //     deferredValidationTime: 0,
                //     key: 'spotlightBGColorFieldId',
                //  }),


                // PropertyFieldColorPickerMini('spotlightFontColor', {
                //   label: strings.spotlightFontColorLableMessage,
                //   initialColor: this.properties.spotlightFontColor,
                //   disabled: false,
                //   onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                //   properties: this.properties,
                //   onGetErrorMessage: null,
                //   deferredValidationTime: 0,
                //   key: 'spotlightFontColorFieldId',
                //   render: function (): void {
                //     throw new Error('Function not implemented.');
                //   }
                // }),
                PropertyFieldColorPicker('spotlightFontColor', {
                  label: "Font color",
                  initialColor: this.properties.spotlightFontColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  render: this.render.bind(this),
                  disableReactivePropertyChanges: this.disableReactivePropertyChanges,
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'spotlightFontColorFieldId'
                }),
                PropertyPaneToggle('enabledSpotlightAutoPlay', {
                  label: "Enable Auto Slider"
                }),
                PropertyPaneSlider('spotlightSliderSpeed', {
                  label: "Slider Speed",
                  min: 0,
                  max: 7,
                  value: 3,
                  showValue: true,
                  step: 0.5
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
