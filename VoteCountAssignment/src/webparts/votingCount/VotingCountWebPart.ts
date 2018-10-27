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
           <button type="button">SAVE</button>
            </div>
            
            
          </div>
        </div>
      </div>`;

      var Url = this.context.pageContext.web.absoluteUrl;

      $(document).ready(function(){
        GetLocation();
        $(document).on("click","#saveid",function(){
          SaveVote();
        });
            
      });
      var VoteId;
      function SaveVote()
      {
        var getvalue = $('input[name=loc]:checked').val();
        VoteId=this.getvalue;
        let html: string = '';
          if (Environment.type === EnvironmentType.Local)
          {
            this.domElement.querySelector('#buttonid').innerHTML = "Sorry this does not work in local workbench";
          } 
          else 
          {
            
      
              var call = jQuery.ajax({
                url:this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('Sravani_Votes')/Items?$select=Votes`,
                  type: "POST",
                  data: JSON.stringify({
                      "__metadata": { type: "SP.Data.VotesListItem" },
                     Votes:VoteId,
                      //AssignedToId: userId,
                      
                  }),
                 
              });
              call.done(function (data, textStatus, jqXHR) {
                alert(VoteId);
                var message =  $("#buttonid");
                $.each(data.d.results,function(value,element){
                message.append(element.VoteId);
              });
            });
              call.fail(function (jqXHR, textStatus, errorThrown) {
                var response = JSON.parse(jqXHR.responseText);
                var message = response ? response.error.message.value : textStatus;
                alert("Call failed. Error: " + message);
              });
          }
            // var call = jQuery.ajax({
            //   url:this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('Sravani_Votes')/Items?$select=Title,ID&$filter=Title/ID eq '${getvalue}'`,
            //   type: "POST",
            //   data: JSON.stringify({
            //     "__metadata": { type: "SP.List" },               
            //     Votes:""
            // }),
            //   }),
            //   headers: {
            //   Accept: "application/json;odata=verbose",
            //     }
            //   call.done(function (data, textstatus, jqXHR) {
            //   var message =  $("#buttonid");
            //   $.each(data.d.results,function(value,element){
            //   message.append(element.ID);
            //   })
            //   });
            //   call.fail(function (jqXHR, textStatus, errorThrown) {
            //       var response = JSON.parse(jqXHR.responseText);
            //       var message = response ? response.error.message.value : textStatus;
            //       alert("Call failed. Error: " + message);
            //   });
         
      }
      function GetLocation()
      {
       
          let html: string = '';
          if (Environment.type === EnvironmentType.Local) {
            this.domElement.querySelector('#buttonid').innerHTML = "Sorry this does not work in local workbench";
          } else 
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
                ButtonId.append(`<input type="radio" name="loc" id="rdbid">${element.Title}`);
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
