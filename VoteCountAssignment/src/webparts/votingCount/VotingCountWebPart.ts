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
//import { IDigestCache, DigestCache } from '@microsoft/sp-http';
import * as $ from 'jquery';
import pnp from "sp-pnp-js";
import {GoogleCharts} from 'google-charts';
import { CurrentUser } from 'sp-pnp-js/lib/sharepoint/siteusers';
require('bootstrap');
var selid;

export interface IVotingCountWebPartProps {
  description: string;
}

export default class VotingCountWebPart extends BaseClientSideWebPart<IVotingCountWebPartProps> {

  // protected onInit(): Promise<void> {
  //   return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
  //     const digestCache: IDigestCache = this.context.serviceScope.consume(DigestCache.serviceKey);
  //     digestCache.fetchDigest(this.context.pageContext.web.serverRelativeUrl).then((digest: string): void => {
  //       use the digest here
      
  //          "X-RequestDigest" : jQuery("#__REQUESTDIGEST").val();
       
  //       resolve();
  //     });
  //   });
  // }
  public render(): void {
   
    let url="https://stackpath.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css";
    SPComponentLoader.loadCss(url); 
    let CurrentUser=this.context.pageContext.user.displayName;
    alert(CurrentUser);
    this.domElement.innerHTML = `
      <div class="${ styles.votingCount }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
            <div id='UserLabelid'></div>
              <div id="buttonid"></div>
              <br>
              <br>
              
              <button type="button" id="saveid" style="color:DodgerBlue;">SAVE</button>
            <div id="chart"></div>  
            </div>

          </div>
        </div>
      </div>`;
      
      var Url = this.context.pageContext.web.absoluteUrl;
     
     
      $(document).ready(function()
      {
        
        GetLocation();
                          
        $(document).on("click",".btncls",function()
        {
      
         selid=$(this).attr('id');
         
         alert("clicked button id is"+selid);         
        $(".btncls").removeClass('active').addClass('disabled');
        $('#'+selid).removeAttr('class');
        $('#'+selid).addClass('active btn btn-success'); 

        });
             
         $(document).on("click","#saveid",function()
            {                        
                SaveVote();
            });
            GetCurrentUser();
        GoogleCharts.load(drawChart);
                var actualdata=([['Location','VotePercent'],
                ['A',10],
                ['B',20]
                ]);
          function drawChart()
           {
         
            // Standard google charts functionality is available as GoogleCharts.api after load
            const data = GoogleCharts.api.visualization.arrayToDataTable(actualdata);
            const pie_1_chart = new GoogleCharts.api.visualization.PieChart(document.getElementById('chart'));
            pie_1_chart.draw(data);
          }
      });
     
       function SaveVote()
       {
          if (Environment.type === EnvironmentType.Local)
          {
            this.domElement.querySelector('#saveid').innerHTML = "Sorry this does not work in local workbench";
          } 
         else 
           {
            alert(selid);
             pnp.sp.web.lists.getByTitle("Sravani_NewVotes").items.add
             ({
               Title: selid
              });
            pnp.sp.web.lists.getByTitle("Sravani_NewVotes").items.getById(selid).update({
            Title: selid
              });
            }
          }
      function GetCurrentUser()
      {
        
          if (Environment.type === EnvironmentType.Local)
          {
            this.domElement.querySelector('#saveid').innerHTML = "Sorry this does not work in local workbench";
          } 
          else
          {
            var call = jQuery.ajax
            ({
             
               url:Url + `/_api/web/lists/getbytitle('Sravani_NewVotes')/Items?$select=LocationId,created By&$filter=Created By`,
              type: "GET",
               data: JSON,   
                
                 headers: 
                 {
                   Accept: "application/json;odata=verbose",
                    "Content-Type": "application/json;odata=verbose",
                   // "X-RequestDigest": $("#__REQUESTDIGEST").val(),
                   // "Authorization": "Bearer " + accessToken
                 }
                                
             });
           
            call.done(function (data, textStatus, jqXHR) 
            {
              var userid=$('#UserLabelid');
              var message =  $("#saveid");
              
              $.each(data.d.results,function(value,element)
              {
                userid.append(`"Already voted",${element.LocationId}`);
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
              var Button =  $("#buttonid");
      
              $.each(data.value,function(val,element){
               
                Button.append(`<button type="button" class="btncls" style="color:DodgerBlue;" id="${element.ID}">${element.Title}</button>&nbsp&nbsp`);
                //alert(`${element.ID}`);
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
