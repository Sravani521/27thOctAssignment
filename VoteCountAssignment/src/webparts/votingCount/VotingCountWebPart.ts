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
import { IDigestCache, DigestCache } from '@microsoft/sp-http';
import * as $ from 'jquery';
import pnp from "sp-pnp-js";
import {GoogleCharts} from 'google-charts';

require('bootstrap');
var CurrentUser;
var SelectedId;
var IsVoted:boolean;
var UpdatedId;
var UserId;


export interface IVotingCountWebPartProps {
  description: string;
}

export default class VotingCountWebPart extends BaseClientSideWebPart<IVotingCountWebPartProps> {

  /*******tried for error request digest********/
 
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
    CurrentUser=this.context.pageContext.user.displayName;
    //alert(CurrentUser);
    this.domElement.innerHTML = `
      <div class="${ styles.votingCount }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
            <div id='UserLabelId'></div>
              <div id=LocButtonId"></div>
              <br>
              <br>
              
              <button type="button" id="SaveBtnId" style="color:DodgerBlue;">SAVE</button>
            <div id="ChartId"></div>  
            </div>
          </div>
        </div>
      </div>`;
      
      var AbsUrl = this.context.pageContext.web.absoluteUrl;
     
     
      $(document).ready(function()
      {
        
        GetLocation();//getting the locations
         
        //on click of location button 
        $(document).on("click",".btncls",function()
        {
      
         SelectedId=$(this).attr('id');
         
         //alert("clicked button id is"+selid);         
        $(".btncls").removeClass('active').addClass('disabled');
        $('#'+SelectedId).removeAttr('class');
        $('#'+SelectedId).addClass('active btn btn-success'); 

        });
           
       //on click of the save button 
         $(document).on("click","#SaveBtnId",function()
            {                                                                     
                if(IsVoted==true)
                {
                  UpdateLoc();
                }
                else if(IsVoted==false)
                {
                  SaveVote();              
                }
          });

          GetCurrentUser();// checking the current user
          
          GoogleCharts.load(drawChart);
          var actualdata=([['Location','VotePercent'],
                  ['A',10],
                  ['B',20]
                  ]);
            function drawChart()
            {
          
              // Standard google charts functionality is available as GoogleCharts.api after load
              const data = GoogleCharts.api.visualization.arrayToDataTable(actualdata);
              const pie_1_chart = new GoogleCharts.api.visualization.PieChart(document.getElementById('ChartId'));
              pie_1_chart.draw(data);
            }
      });

      //updating the location if same user want to change the vote
     function UpdateLoc()
     {
      // alert("coming to update");
      if (Environment.type === EnvironmentType.Local)
      {
        this.domElement.querySelector('#SaveBtnId').innerHTML = "Sorry this does not work in local workbench";
      } 
      else
      {
        pnp.sp.web.lists.getByTitle("Sravani_NewVotes").items.getById(UpdatedId).update
        ({            
            LocationId:SelectedId
        });
      }
     }

     //saving the vote of the user
       function SaveVote()
       {
          if (Environment.type === EnvironmentType.Local)
          {
            this.domElement.querySelector('#SaveBtnId').innerHTML = "Sorry this does not work in local workbench";
          } 
         else 
           {
           // alert(selid);
            
             pnp.sp.web.lists.getByTitle("Sravani_NewVotes").items.add
             ({
               
               CUser:CurrentUser,
               LocationId:SelectedId
               
             });
           
            
            }
          }

        //getting the current user   
      function GetCurrentUser()
      {
        
          if (Environment.type === EnvironmentType.Local)
          {
            this.domElement.querySelector('#UserLabelId').innerHTML = "Sorry this does not work in local workbench";
          } 
          else
          {
            var call = jQuery.ajax
            ({
             
               url:AbsUrl + `/_api/web/lists/getbytitle('Sravani_NewVotes')/Items?$select=LocationId,Title,ID,CUser&$filter=CUser eq '${CurrentUser}'`,
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
              UserId=$('#UserLabelId');
              
              IsVoted=false;
              $.each(data.d.results,function(value,element)
              {
                UserId.append(` ${element.CUser} you have already voted `);
                UpdatedId=`${element.ID}`;
                $(".btncls").removeClass('active').addClass('disabled');
                $('#'+element.Title).removeAttr('class');
                $('#'+element.Title).addClass('active btn btn-success'); 
                IsVoted=true;
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
              
        //retreiving all the locations 
      function GetLocation()
      {
        if (Environment.type === EnvironmentType.Local)
        {
            this.domElement.querySelector('#ButtonId').innerHTML = "Sorry this does not work in local workbench";
        }
        else 
        {
            var call =  $.ajax({
              url: AbsUrl + `/_api/web/lists/getbytitle('Sravani_Location')/Items?$select=Title,ID`,
              type: "GET",
              dataType: "json",
            }),
              headers: {
                Accept: "application/json;odata=verbose",
              }
            call.done(function (data, textstatus, jqXHR) {
              var Button =  $("#ButtonId");

             
              $.each(data.value,function(val,element){
               
                Button.append(`<button type="button" class="btncls btncls-success" style="color:DodgerBlue;" id="${element.ID}">${element.Title}</button>&nbsp&nbsp`);
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
