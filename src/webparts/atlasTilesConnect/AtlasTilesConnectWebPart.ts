
import {} from 'jquery';

import PnPTelemetry from "@pnp/telemetry-js";
const telemetry = PnPTelemetry.getInstance();
telemetry.optOut();

//import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { PropertyFieldPeoplePicker, PrincipalType } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';

//import { people } from 'TileImageWebPartStrings';
import { IPropertyFieldGroupOrPerson } from "@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker";

import { SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration } from '@microsoft/sp-http';

import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneDynamicField,
  PropertyPaneLink,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './AtlasTilesConnectWebPart.module.scss';
import * as strings from 'AtlasTilesConnectWebPartStrings';
// import { PropertyFieldFilePicker, IPropertyFieldFilePickerProps, IFilePickerResult } from "@pnp/spfx-property-controls/lib/PropertyFieldFilePicker";


// import "@pnp/sp/webs";
// import "@pnp/sp/items";
// import "@pnp/sp/folders";
// import "@pnp/sp/lists";
// import "@pnp/sp/webs";
// import "@pnp/sp/lists";

import { SPComponentLoader } from '@microsoft/sp-loader';


export interface IAtlasTilesConnectWebPartProps {
  description: string;
  ImageURL: string;
  Hyperlink: string;
  TargetAudience: string;
  people: IPropertyFieldGroupOrPerson[];
  context: WebPartContext;
}

export default class AtlasTilesConnectWebPart extends BaseClientSideWebPart<IAtlasTilesConnectWebPartProps> {

  public render(): void {
    if (!this.renderedOnce){
      console.log("SCRIPT LOADED...");
    SPComponentLoader.loadCss('https://use.fontawesome.com/releases/v5.0.9/css/all.css'); 
    SPComponentLoader.loadScript('https://code.jquery.com/jquery-1.7.1.min.js');
    }

    this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/Web/CurrentUser/Groups`,
      SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        response.json().then((responseJSON: any) => {
          // console.log(responseJSON.value);
          var finalArray = responseJSON.value.map(function (obj: { Title: any; }) {
            return obj.Title;
          });
          ///console.log(finalArray);//Array Retrieved from Current users Groups.....

          if (this.properties.people && this.properties.people.length > 0) {
            ///console.log(JSON.stringify(this.properties.people));

            const GroupArray = this.properties.people.map((obj: { fullName: any; }) => {
              return obj.fullName;
            });
            ///console.log(GroupArray);//Array Of Group in property pane
            console.log("Current User Present In The Group");
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
					   background-repeat: no-repeat;width:120%;height:287px;
					   background-size:cover;
					   background-position: center;">
		   
					   <a  class="${styles.callToAction}" onMouseOver="this.style.color='#CC0A0A'; style.backgroundColor='rgba(255, 255, 255, 0.7)'" onMouseOut="this.style.color='#424242'; style.backgroundColor='rgba(255, 255, 255, 0.5)'" 
					   
					   style="
					   display: block;
					   float: left;
					   background: rgba(255, 255, 255, 0.5);
					   margin-top: 2.25em;
					   //vertical-align: middle;
					   text-align: left;
					   font-family: 'Oswald' !important;
					   text-decoration: none;
					   font-size: 3em;
					   padding: 0.25em 0.5em 0.25em calc(2% + 0em);
					   color: #424242;
					   text-transform: uppercase;" href="${escape(this.properties.Hyperlink)}" target="_blank" unselectable="on" >
					   ${escape(this.properties.description)} 
		   
		   
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
            <div>
              <i id='btnTest' class="fas fa-arrow-circle-right"></i>
              <i id='btn2' class="fas fa-arrow-circle-right" onClick=" ${this.swap()}"></i>
              <div id='a'> hover here </div>
              <div id='b'> Change color</div>
            </div>
            `;
            
          }
          

          $(document).on('mouseover', '#a', function (e) {
            $("#b").css("background-color", "yellow");         
          });

          $(document).on('mouseout', '#a', function (e) {
            $("#b").css("background-color", "red");
          //  $("#b").addClass(`${styles.blue}`);        

           });

          //  $(document).on('click', '#btnTest', function (e) {
          //   $("#b").addClass(`${styles.blue}`);        
          // });

          // $(document).ready(function(){
          //   $("button").click(function(){
          //     $("h1, h2, p").addClass("blue");
          //     $("div").addClass("important");
          //   });
          // });
         


        });

      });




    }
    
    public swap(): any {
      console.log("Triggered!");
      // document.getElementById("flip").className == "expanded" ? document.getElementById("flip").className = "collapsed": document.getElementById("flip").className = "expanded";
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

                })






              ]

            }


          ]

        }
      ]
    };
  }
}
