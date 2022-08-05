
import { } from 'jquery';

import PnPTelemetry from "@pnp/telemetry-js";
const telemetry = PnPTelemetry.getInstance();
telemetry.optOut();

//import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { PropertyFieldPeoplePicker, PrincipalType } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';

//import { people } from 'TileImageWebPartStrings';
import { IPropertyFieldGroupOrPerson } from "@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker";
import { PropertyFieldMessage } from '@pnp/spfx-property-controls/lib/PropertyFieldMessage';

import { SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration } from '@microsoft/sp-http';

import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneCheckbox,
  PropertyPaneDynamicField,
  PropertyPaneLink,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './AtlasTilesConnectWebPart.module.scss';
import * as strings from 'AtlasTilesConnectWebPartStrings';
// import { PropertyFieldFilePicker, IPropertyFieldFilePickerProps, IFilePickerResult } from "@pnp/spfx-property-controls/lib/PropertyFieldFilePicker";


import "@pnp/sp/webs";
import "@pnp/sp/items";
import "@pnp/sp/folders";
import "@pnp/sp/lists";
import "@pnp/sp/webs";
import "@pnp/sp/lists";

import { SPComponentLoader } from '@microsoft/sp-loader';
import { sp } from '@pnp/sp';

import "isomorphic-fetch"; // or import the fetch polyfill you installed
import { Client } from "@microsoft/microsoft-graph-client";
import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

export interface IAtlasTilesConnectWebPartProps {
  LangEnglish: any;
  LangChinese: any;
  LangGerman: any;
  LangSpanish: any;
  LangFrench: any;
  LangPolish: any;
  LangJapanese: any;
  LangPortuguese: any;
  LangRussian: any;

  EnglishText: any;
  ChineseText: any;
  GermanText: any;
  SpanishText: any;
  FrenchText: any;
  PolishText: any;
  JapaneseText: any;
  PortugueseText: any;
  RussianText: any;


  description: string;
  ImageURL: string;
  Hyperlink: string;
  TargetAudience: string;
  people: IPropertyFieldGroupOrPerson[];
  context: WebPartContext;
  currUserLang: string;
}

export default class AtlasTilesConnectWebPart extends BaseClientSideWebPart<IAtlasTilesConnectWebPartProps> {


  public async render(): Promise<void> {


    var siteUrl = this.context.pageContext.web.absoluteUrl ///Get Site Url
    // console.log(siteUrl)

    const myArray = siteUrl.split("/");
    var siteName = myArray[myArray.length - 1].split(".")[0]; ///Get Site Name
    // console.log(siteName)
    var testuser = this.context.pageContext.user;
    console.log(testuser)

    var userEmail = this.context.pageContext.user.email;
    this.context.spHttpClient.get(`${siteUrl}/_api/web/lists/getbytitle('Preference')/Items?&$filter=Title eq '${userEmail}'`,
      SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        response.json().then((responseJSON: any) => {
          console.log(responseJSON.value);
          var prefLanguage = responseJSON.value.map(function (obj: { Language: any; }) {
            return obj.Language;
          });
          console.log(prefLanguage)



          if (!this.renderedOnce) {
            console.log("SCRIPT LOADED...");
            SPComponentLoader.loadCss('https://use.fontawesome.com/releases/v5.0.9/css/all.css');
            SPComponentLoader.loadScript('https://code.jquery.com/jquery-1.7.1.min.js');
          }

          // this.context.spHttpClient.get(`${siteUrl}/_api/Web/CurrentUser/Groups`,
          //   SPHttpClient.configurations.v1)
          //   .then((response: SPHttpClientResponse) => {
          //     response.json().then((responseJSON: any) => {
          //       console.log(responseJSON.value);
          //       var finalArray = responseJSON.value.map(function (obj: { Title: any; }) {
          //         return obj.Title;
          //       });
          //       console.log(finalArray);

          this.context.msGraphClientFactory.getClient().then(async (client: MSGraphClient): Promise<void> => {
            let group = await client.api('/me/memberOf/$/microsoft.graph.group')
              .filter('groupTypes/any(a:a eq \'unified\')')
              .get();
            console.log(group.value)
            var finalArray = group.value.map(function (obj: { displayName: any; }) {
              return obj.displayName;
            });
            console.log(finalArray)
            // finalArray =group1






            //Array Retrieved from Current users Groups.....
            if (this.properties.people && this.properties.people.length > 0) {
              ///console.log(JSON.stringify(this.properties.people));
              console.log(this.properties.people)
              // const GroupArray = this.properties.people.map((obj: { fullName: any; }) => {
              //   return obj.fullName;
              // });
            var tempPeopleArray = this.properties.people
            const GroupArray = tempPeopleArray.map(element => element.description);
            console.log(GroupArray)


          var usrFullname = this.context.pageContext.user.displayName;
              var Groupintersections = finalArray.filter(e => GroupArray.indexOf(e) !== -1);
              console.log(Groupintersections)

              ///console.log(GroupArray);//Array Of Group in property pane
              if (GroupArray.includes(usrFullname) || Groupintersections.length > 0) {
                // console.log("Current User Present In The Group");
                this.domElement.innerHTML = `
             <head>
             <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js"></script>
             <link rel="preconnect" href="https://fonts.googleapis.com">
              <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
              <link href="https://fonts.googleapis.com/css2?family=Oswald&display=swap"     rel="stylesheet">
              </head>
					   <script src='https://kit.fontawesome.com/a076d05399.js' crossorigin='anonymous'></script>
					   <div class="ms-rte-embedcode ms-rte-embedwp">
					   <div class="${styles.MainContainer}"
					   style="background-image: url(${escape(this.properties.ImageURL)});
             box-shadow: 0 4px 8px 0 rgb(0 0 0 / 20%), 0 6px 20px 0 rgb(0 0 0 / 19%);
					   background-repeat: no-repeat;width:100%;height:200px;
					   background-size:cover;
					   background-position: center;">
		   
					   <a  class="${styles.callToAction}" onMouseOver="this.style.color='#CC0A0A'; style.backgroundColor='rgba(255, 255, 255, 0.8)'" onMouseOut="this.style.color='#424242'; style.backgroundColor='rgba(255, 255, 255, 0.7)'" 
					   
					   style="
					   display: block;
					   float: left;
					   background: rgba(255, 255, 255, 0.7);
					   margin-top: 2.25em;
					   //vertical-align: middle;
					   text-align: left;
					   font-family: 'Oswald' !important;
					   text-decoration: none;
					   font-size: 3em;
					   padding: 0.25em 0.5em 0.25em calc(2% + 0em);
					   color: #424242;
					   text-transform: uppercase;" href="${escape(this.properties.Hyperlink)}" target="_blank" unselectable="on" >
            ${prefLanguage[0].includes("English") && this.properties.LangEnglish == true ?
                    this.properties.EnglishText :
                    prefLanguage[0].includes("Chinese") && this.properties.LangChinese == true ?
                      this.properties.ChineseText :
                      prefLanguage[0].includes("German") && this.properties.LangGerman == true ?
                        this.properties.GermanText :
                        prefLanguage[0].includes("Spanish") && this.properties.LangSpanish == true ?
                          this.properties.SpanishText :
                          prefLanguage[0].includes("French") && this.properties.LangFrench == true ?
                            this.properties.FrenchText :
                            prefLanguage[0].includes("Polish") && this.properties.LangPolish == true ?
                              this.properties.PolishText :
                              prefLanguage[0].includes("Japanese") && this.properties.LangJapanese == true ?
                                this.properties.JapaneseText :
                                prefLanguage[0].includes("Portuguese") && this.properties.LangPortuguese == true ?
                                  this.properties.PortugueseText :
                                  prefLanguage[0].includes("Russian") && this.properties.LangRussian == true ?
                                    this.properties.RussianText :
                                    `${escape(this.properties.description)}`

                  }
					   
		   
		   
					   <i style="
						 border: solid #CC0A0A;
						 font-color: #CC0A0A;    
						 border-width: 0 4px 4px 0;
						 display: inline-block;
						 padding: 10px;
						 height:10px; width:10px;
						 transform: rotate(-45deg);
						 -webkit-transform: rotate(-45deg);">
					   </i></a>
					   
					 </div></div>
					 `;
              }
              else {
                this.domElement.innerHTML = `
                
              `;

              }
            }
            else {
              this.domElement.innerHTML = `
                  <head>
                  <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js"></script>
                  <link rel="preconnect" href="https://fonts.googleapis.com">
                   <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
                   <link href="https://fonts.googleapis.com/css2?family=Oswald&display=swap"     rel="stylesheet">
                   </head>
                  <script src='https://kit.fontawesome.com/a076d05399.js' crossorigin='anonymous'></script>
                  <div class="ms-rte-embedcode ms-rte-embedwp">
                  <div class="${styles.MainContainer}"
                  style="background-image: url(${escape(this.properties.ImageURL)});
                  box-shadow: 0 4px 8px 0 rgb(0 0 0 / 20%), 0 6px 20px 0 rgb(0 0 0 / 19%);
                  background-repeat: no-repeat;width:100%;height:200px;
                  background-size:cover;
                  background-position: center;">
            
                  <a  class="${styles.callToAction}" onMouseOver="this.style.color='#CC0A0A'; style.backgroundColor='rgba(255, 255, 255, 0.8)'" onMouseOut="this.style.color='#424242'; style.backgroundColor='rgba(255, 255, 255, 0.7)'" 
                  
                  style="
                  display: block;
                  float: left;
                  background: rgba(255, 255, 255, 0.7);
                  margin-top: 2.25em;
                  //vertical-align: middle;
                  text-align: left;
                  font-family: 'Oswald' !important;
                  text-decoration: none;
                  font-size: 3em;
                  padding: 0.25em 0.5em 0.25em calc(2% + 0em);
                  color: #424242;
                  text-transform: uppercase;" href="${escape(this.properties.Hyperlink)}" target="_blank" unselectable="on" >
                 ${prefLanguage[0].includes("English") && this.properties.LangEnglish == true ?
                  this.properties.EnglishText :
                  prefLanguage[0].includes("Chinese") && this.properties.LangChinese == true ?
                    this.properties.ChineseText :
                    prefLanguage[0].includes("German") && this.properties.LangGerman == true ?
                      this.properties.GermanText :
                      prefLanguage[0].includes("Spanish") && this.properties.LangSpanish == true ?
                        this.properties.SpanishText :
                        prefLanguage[0].includes("French") && this.properties.LangFrench == true ?
                          this.properties.FrenchText :
                          prefLanguage[0].includes("Polish") && this.properties.LangPolish == true ?
                            this.properties.PolishText :
                            prefLanguage[0].includes("Japanese") && this.properties.LangJapanese == true ?
                              this.properties.JapaneseText :
                              prefLanguage[0].includes("Portuguese") && this.properties.LangPortuguese == true ?
                                this.properties.PortugueseText :
                                prefLanguage[0].includes("Russian") && this.properties.LangRussian == true ?
                                  this.properties.RussianText :
                                  `${escape(this.properties.description)}`

                }
                  
            
            
                  <i style="
                  border: solid #CC0A0A;
                  font-color: #CC0A0A;    
                  border-width: 0 4px 4px 0;
                  display: inline-block;
                  padding: 10px;
                  height:10px; width:10px;
                  transform: rotate(-45deg);
                  -webkit-transform: rotate(-45deg);">
                  </i></a>
                  
                </div></div>
                `;

            }

          });


          //   });

          // });

        })
      });


  }



  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {

    let EnglishProperty: any;
    let ChineseProperty: any;
    let GermanProperty: any;
    let SpanishProperty: any;
    let FrenchProperty: any;
    let PolishProperty: any;
    let JapaneseProperty: any;
    let PortugueseProperty: any;
    let RussianProperty: any;

    if (this.properties.LangEnglish == true) {
      EnglishProperty = PropertyPaneTextField('EnglishText', {
        label: "",
        value: this.properties.EnglishText
      })
    }
    else {
      EnglishProperty = ""
    };
    ////////////////////////////////////////////////////////////
    if (this.properties.LangChinese == true) {
      ChineseProperty = PropertyPaneTextField('ChineseText', {
        label: "",
        value: this.properties.ChineseText
      })
    }
    else {
      ChineseProperty = ""
    };
    /////////////////////////////////////////////////////////////
    if (this.properties.LangGerman == true) {
      GermanProperty = PropertyPaneTextField('GermanText', {
        label: "",
        value: this.properties.GermanText
      })
    }
    else {
      GermanProperty = ""
    };
    ////////////////////////////////////////////////////////////
    if (this.properties.LangSpanish == true) {
      SpanishProperty = PropertyPaneTextField('SpanishText', {
        label: "",
        value: this.properties.SpanishText
      })
    }
    else {
      SpanishProperty = ""
    };
    ////////////////////////////////////////////////////////////
    if (this.properties.LangFrench == true) {
      FrenchProperty = PropertyPaneTextField('FrenchText', {
        label: "",
        value: this.properties.FrenchText
      })
    }
    else {
      FrenchProperty = ""
    };
    ///////////////////////////////////////////////////////////////
    if (this.properties.LangPolish == true) {
      PolishProperty = PropertyPaneTextField('PolishText', {
        label: "",
        value: this.properties.PolishText
      })
    }
    else {
      PolishProperty = ""
    };
    //////////////////////////////////////////////////////////////
    if (this.properties.LangJapanese == true) {
      JapaneseProperty = PropertyPaneTextField('JapaneseText', {
        label: "",
        value: this.properties.JapaneseText
      })
    }
    else {
      JapaneseProperty = ""
    };
    /////////////////////////////////////////////////////////////
    if (this.properties.LangPortuguese == true) {
      PortugueseProperty = PropertyPaneTextField('PortugueseText', {
        label: "",
        value: this.properties.PortugueseText
      })
    }
    else {
      PortugueseProperty = ""
    };
    //////////////////////////////////////////////////////////
    if (this.properties.LangRussian == true) {
      RussianProperty = PropertyPaneTextField('RussianText', {
        label: "",
        value: this.properties.RussianText
      })
    }
    else {
      RussianProperty = ""
    };
    ///////////////////////////////////////////////////////////

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
                }),
                PropertyPaneTextField('ImageURL', {
                  label: strings.ImageURLFieldLabel,
                  value: ""
                }),
                PropertyPaneTextField('Hyperlink', {
                  label: strings.Hyperlinklabel,
                  value: ""
                }),
                // PropertyPaneTextField('TargetAudience', {
                //   label: "Target Audience",
                //   value: "Americas Users"
                // }),
                PropertyFieldPeoplePicker('people', {
                  label: 'Target Audience',
                  initialData: this.properties.people,
                  allowDuplicate: false,
                  principalType: [PrincipalType.Users, PrincipalType.SharePoint, PrincipalType.Security],
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  context: this.context as any,
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'peopleFieldId'

                }),

                PropertyPaneCheckbox('LangEnglish', {
                  text: "English",
                  checked: false,
                  disabled: false
                }),
                EnglishProperty,
                PropertyPaneCheckbox('LangChinese', {
                  text: "Chinese",
                  checked: false,
                  disabled: false
                }),
                ChineseProperty,
                PropertyPaneCheckbox('LangGerman', {
                  text: "German",
                  checked: false,
                  disabled: false
                }),
                GermanProperty,
                PropertyPaneCheckbox('LangSpanish', {
                  text: "Spanish",
                  checked: false,
                  disabled: false
                }),
                SpanishProperty,
                PropertyPaneCheckbox('LangFrench', {
                  text: "French",
                  checked: false,
                  disabled: false
                }),
                FrenchProperty,
                PropertyPaneCheckbox('LangPolish', {
                  text: "Polish",
                  checked: false,
                  disabled: false
                }),
                PolishProperty,
                PropertyPaneCheckbox('LangJapanese', {
                  text: "Japanese",
                  checked: false,
                  disabled: false
                }),
                JapaneseProperty,
                PropertyPaneCheckbox('LangPortuguese', {
                  text: "Portuguese",
                  checked: false,
                  disabled: false
                }),
                PortugueseProperty,
                PropertyPaneCheckbox('LangRussian', {
                  text: "Russian",
                  checked: false,
                  disabled: false
                }),
                RussianProperty

              ]

            }


          ]

        }
      ]
    };
  }
}
