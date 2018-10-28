import { Version, Environment, EnvironmentType } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import{SPComponentLoader} from '@microsoft/sp-loader';
import styles from './VotingCountWebPart.module.scss';
import * as strings from 'VotingCountWebPartStrings';
import * as $ from 'jquery';
import pnp from "sp-pnp-js";
require('bootstrap');

export interface IVotingCountWebPartProps {
  description: string;
}

export default class VotingCountWebPart extends BaseClientSideWebPart<IVotingCountWebPartProps> {

  public render(): void {
    let url="https://stackpath.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css";
    SPComponentLoader.loadCss(url); 
    this.domElement.innerHTML = `
      <div class="${ styles.votingCount }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <div id="buttonid"></div>
              <button type="button" id="saveid">SAVE</button>
            </div>
          </div>
        </div>
      </div>`;
      var Url = this.context.pageContext.web.absoluteUrl;
      $(document).ready(function()
      {
        GetLocation();
        $(document).on("click","#saveid",function()
        {
         var getvalue = $('input[name=loc]:checked').attr('value');         
         SaveVote(getvalue);
        });
      });
      function SaveVote(getvalue)
      {
        if (Environment.type === EnvironmentType.Local)
          {
            this.domElement.querySelector('#saveid').innerHTML = "Sorry this does not work in local workbench";
          } 
        else 
          {
           alert("else");
            // pnp.sp.web.lists.getByTitle("Sravani_NewVotes").items.select("Location", "Lookup/Title").expand("Lookup").get().then((items: any[]) => 
            // {
            // console.log(getvalue);
            // });
            var call = jQuery.ajax(
            {
              url:Url + `/_api/web/lists/getbytitle('Sravani_NewVotes')/Items/?$top=1`,
              type: "GET",
              data: JSON.stringify(
                {
                  "__metadata": { type: "SP.Data.Sravani_NewVotesListItem" },
                  Title: "sample",
                  Votes:1,
                  Location:getvalue
                }),
                headers: 
                {
                  Accept: "application/json;odata=verbose"
                  "Content-Type": "application/json;odata=verbose",
                  "X-RequestDigest": jQuery("#__REQUESTDIGEST").val()
                }
            });
            call.done(function (data, textStatus, jqXHR) 
            {
              var value=getvalue;
              var message =  $("#saveid");
              $.each(data.d.results,function(value,element)
              {
              message.append(element.VoteId);
              });
            });
            call.fail(function (jqXHR, textStatus, errorThrown) 
            {
              var response = JSON.parse(jqXHR.responseText);
              var message = response ? response.error.message.value : textStatus;
              alert("Call failed. Error: " + message);
            });
           
          }
        }
              
         
      function GetLocation()
      {
        if (Environment.type === EnvironmentType.Local)
        {
            this.domElement.querySelector('#buttonid').innerHTML = "Sorry this does not work in local workbench";
        }
        else 
        {
            var call =  $.ajax({
              url: Url + `/_api/web/lists/getbytitle('Sravani_Location')/Items?$select=Title,ID`,
              type: "GET",
              dataType: "json",
            }),
              headers: {
                Accept: "application/json;odata=verbose",
              }
            call.done(function (data, textstatus, jqXHR) {
              var ButtonId =  $("#buttonid");
      
              $.each(data.value,function(val,element){
                //alert(element.Title);
                ButtonId.append(`<input type="radio" name="loc" id="rdbid" value="${element.Title}">${element.Title}`);
              });
                
              
            });
            call.fail(function (jqXHR, textStatus, errorThrown) {
                var response = JSON.parse(jqXHR.responseText);
                var message = response ? response.error.message.value : textStatus;
                alert("Call failed. Error: " + message);
            });          
        }
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
