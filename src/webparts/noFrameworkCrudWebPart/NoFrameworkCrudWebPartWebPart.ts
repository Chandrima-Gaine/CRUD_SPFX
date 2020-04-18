import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import { SPHttpClient, SPHttpClientResponse } from'@microsoft/sp-http'; 
import { IListItem } from'./IListItem'; 
import { SPComponentLoader } from'@microsoft/sp-loader';

import styles from './NoFrameworkCrudWebPartWebPart.module.scss';
import * as strings from 'NoFrameworkCrudWebPartWebPartStrings';

SPComponentLoader.loadCss('https://arksptraining.sharepoint.com/sites/AppCatalogSite/ChandrimaLib/CRUD/styles.css');

export interface INoFrameworkCrudWebPartWebPartProps {
  listname: string;
}

export default class NoFrameworkCrudWebPartWebPart extends BaseClientSideWebPart<INoFrameworkCrudWebPartWebPartProps> {

private listItemEntityTypeName: string = undefined; 
private Listname : string="ChandrimaJobPortal";

public render(): void { 
  this.domElement.innerHTML = ` 
  <div class="${styles.noFrameworkCrudWebPart}">
  <div class="${styles.container}">
  <div class="${styles.row}">
  <div class="${styles.column}">
  <span class="${styles.title}"></span>
  
  <div class="row">
  <h2 style="text-align:left; color:#1C6EA4; font-weight: bold;" id="statusMode">
                Chandrima's CRUD Operation Using SPFX <br>
                Finance, Banking & Sales Job Apply Portal
  </h2><br>
  <h2 style="text-align:Center; color:#1C6EA4; font-weight: bold;" id="statusMode">
                Job Application Form
  </h2>
  <h4 style="text-align:Center; font-style: italic; font-weight: normal; color:#1C6EA4;" id="statusMode">
                Thank you for your interest in working with us
  </h4>  
  <div class="row">
  <div class="col-25">
  <label for="fname">Job Title</label>
  </div>
  <div class="col-75">
  <input type="text" id="idTitle" name="Title" placeholder="Job Title..">
  </div>
  </div>
  <div class="row">
  <div class="col-25">
  <label for="fname">Name</label>
  </div>
  <div class="col-75">
  <input type="text" id="idfname" name="firstname" placeholder="Name..">
  </div>
  </div>
  <div class="row">
  <div class="col-25">
  <label for="lname">Gender</label>
  </div>
  <div class="col-75">
  <input type="text" id="idgender" name="gender" placeholder="Male/Female">
  </div>
  </div>
  <div class="row">
  <div class="col-25">
  <label for="country">Department</label>
  </div>
  <div class="col-75">
  <input type="text" id="idDepart" name="lname" placeholder="Finance/Banking/Sales..">
  </div>
  </div>
  <div class="row">
  <div class="col-25">
  <label for="subject">City</label>
  </div>
  <div class="col-75">
  <input type="text" id="idCity" name="gender" placeholder="City..">
  </div>
  </div>
  <!-- hidden controls -->
  <div style="display: none">
  <input id="recordId" />
  </div>
  <div class="row">
  <div class="ms-Grid-row ms-bgColor-themeDarkms-fontColor-white ${styles.row}">
  
  <button class="${styles.button} create-Button">
  <span class="${styles.label}">Save</span>
  </button>
  <button class="${styles.button} update-Button">
  <span class="${styles.label}">Update</span>
  </button>
  <button  class="${styles.button} read-Button">
  <span  class="${styles.label}">Clear All</span>
  </button>
  
  </div>
  </div>
  <div class="divTable blueTable">
  <div class="divTableHeading">
  <div class="divTableRow">
  <div class="divTableHead">Title</div>
  <div class="divTableHead">Name</div>
  <div class="divTableHead">Gender</div>
  <div class="divTableHead">Department</div>
  <div class="divTableHead">City</div>
  </div>
  </div>
  <div class="divTableBody" id="fileGrid">
  </div>
  </div>
  <div class="blueTable outerTableFooter"><div class="tableFootStyle"><div class="links"><a href="#">&laquo;</a><a class="active" href="#">1</a><a href="#">2</a><a href="#">3</a><a href="#">4</a><a href="#">&raquo;</a></div></div></div>
  </div>
  </div>
  </div>
  </div>`; 
  this.setButtonsEventHandlers();
  this.getAllItem();
    } 
  
  private setButtonsEventHandlers(): void { 
  const webPart: NoFrameworkCrudWebPartWebPart = this; 
  this.domElement.querySelector('button.create-Button').addEventListener('click', () => { webPart.SaveItem(); }); 
  this.domElement.querySelector('button.update-Button').addEventListener('click', () => { webPart.updateItem(); }); 
  this.domElement.querySelector('button.read-Button').addEventListener('click', () => { webPart.ClearMethod(); }); 
    }
  // Start Get All Data From SharePoint  List
  private getAllItem(){ 
  this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.Listname}')/items?$orderby=Created desc`, 
             SPHttpClient.configurations.v1, 
             { 
  headers: { 
  'Accept': 'application/json;odata=nometadata', 
  'odata-version': ''
               } 
             })
             .then((response: SPHttpClientResponse)=> { 
  return response.json(); 
             }) 
             .then((item):void => {
  debugger;
  var len = item.value.length;
  var txt = "";
  if(len>0){
  for(var i=0;i<len;i++){
  
  txt += '<div class="divTableRow" ><div class="divTableCell">'+item.value[i].Title +'</div><div class="divTableCell">'+item.value[i].NameF +'</div><div class="divTableCell">'+item.value[i].Gender+'</div>' +
  '<div class="divTableCell">'+item.value[i].Department+'</div><div class="divTableCell">'+item.value[i].City+'</div><div class="divTableCell">'+"<a id='" + item.value[i].ID + "' href='#' class='EditFileLink'>Edit</a>"+'</div><div class="divTableCell">'+"<a id='" + item.value[i].ID + "' href='#' class='DeleteLink'>Delete</a>"+'</div></div>';
                   }
  if(txt != ""){
  document.getElementById("fileGrid").innerHTML = txt;
                    }
  
  //  delete Start Bind The Event into anchor Tag
  let listItems = document.getElementsByClassName("DeleteLink");
  for(let j:number = 0; j<listItems.length; j++){
  listItems[j].addEventListener('click', (event) => {
  this.DeleteItemClicked(event);
                      });
                    }
  // End Bind The Event into anchor Tag
  
  //  Edit Start Bind The Event into anchor Tag
  let EditlistItems = document.getElementsByClassName("EditFileLink");
  for(let j:number = 0; j<EditlistItems.length; j++){
  EditlistItems[j].addEventListener('click', (event) => {
  this.UpdateItemClicked(event);
                        });
                      }
  // End Bind The Event into anchor Tag
               }
  // debugger;
  
  //console.log(item.Title) ;
  
             }, (error: any): void => { 
  //alert(error);
             }); 
        }
  // End Get All Data From SharePoint  List
  //Start Save Item in SharerPoint List
  private SaveItem(): void { 
  debugger;
  if(document.getElementById('idTitle')["value"]=="")
      {
  alert('Required the Title !!!');
  return;
      }
  if(document.getElementById('idfname')["value"]=="")
      {
  alert('Required the Name !!!');
  return;
      }
  if(document.getElementById('idgender')["value"]=="")
      {
  alert('Required the Gender !!!');
  return;
      }
  if(document.getElementById('idDepart')["value"]=="")
      {
  alert('Required the Department !!!');
  return;
      }
  if(document.getElementById('idCity')["value"]=="")
      {
  alert('Required the City !!!');
  return;
      }
  const body: string = JSON.stringify({ 
  'Title': document.getElementById('idTitle')["value"],
  'NameF': document.getElementById('idfname')["value"],
  'Gender': document.getElementById('idgender')["value"],
  'Department': document.getElementById('idDepart')["value"],
  'City': document.getElementById('idCity')["value"]    
      });  
  this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.Listname}')/items`, 
      SPHttpClient.configurations.v1, 
      { 
  headers: { 
  'Accept': 'application/json;odata=nometadata', 
  'Content-type': 'application/json;odata=nometadata', 
  'odata-version': ''
        }, 
  body: body 
      }) 
      .then((response: SPHttpClientResponse): Promise<IListItem>=> { 
  return response.json(); 
      }) 
      .then((item: IListItem): void => {
  this.ClearMethod();
  alert('Item has been successfully Saved ');
  localStorage.removeItem('ItemId');
  localStorage.clear();
  this.getAllItem();
      }, (error: any): void => { 
  alert(`${error}`); 
      }); 
    } 
  //End Save Item in SharerPoint List
  
  //Start Update Item in SharerPoint List
  private UpdateItemClicked(ev): void{
  let me:any = ev.target;
  this.getByIdItem(me.id);
    }
  private updateItem(){
  ///alert(localStorage.getItem('ItemId')) ;
  const body: string = JSON.stringify({ 
  'Title': document.getElementById('idTitle')["value"],
  'NameF': document.getElementById('idfname')["value"],
  'Gender': document.getElementById('idgender')["value"],
  'Department': document.getElementById('idDepart')["value"],
  'City': document.getElementById('idCity')["value"]    
      });  
  this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.Listname}')/items(${localStorage.getItem('ItemId')})`, 
            SPHttpClient.configurations.v1, 
            { 
  headers: { 
  'Accept': 'application/json;odata=nometadata', 
  'Content-type': 'application/json;odata=nometadata', 
  'odata-version': '', 
  'IF-MATCH': '*', 
  'X-HTTP-Method': 'MERGE'
              }, 
  body: body
            }) 
            .then((response: SPHttpClientResponse): void => { 
  alert(`Item with ID: ${localStorage.getItem('ItemId')} successfully updated`);
  this.ClearMethod();
  localStorage.removeItem('ItemId');
  localStorage.clear();
  this.getAllItem();
  
            }, (error: any): void => { 
  alert(`${error}`); 
            }); 
  
    } 
  //End Update Item in SharerPoint List
  // Delete the Items From SharePoint List
  private DeleteItemClicked(ev): void{
  let me:any = ev.target;
  //alert(me.id);
  this.deleteItem(me.id);
    }
  private deleteItem(Id: number){
  if (!window.confirm('Are you sure you want to delete the latest item?')) { 
  return; 
      } 
  let etag: string = undefined; 
  this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.Listname}')/items(${Id})`, 
            SPHttpClient.configurations.v1, 
            { 
  headers: { 
  'Accept': 'application/json;odata=nometadata', 
  'Content-type': 'application/json;odata=verbose', 
  'odata-version': '', 
  'IF-MATCH': '*', 
  'X-HTTP-Method': 'DELETE'
              } 
            })
        .then((response: SPHttpClientResponse): void => { 
  alert(`Item with ID: ${Id} successfully Deleted`);
  
  this.getAllItem();
        }, (error: any): void => { 
  alert(`${error}`); 
        }); 
    } 
  // End Delete the Items From SharePoint List
  
  // Start Get Item By Id
  private getByIdItem(Id: number){ 
  this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.Listname}')/items(${Id})`, 
       SPHttpClient.configurations.v1, 
       { 
  headers: { 
  'Accept': 'application/json;odata=nometadata', 
  'odata-version': ''
         } 
       })
       .then((response: SPHttpClientResponse)=> { 
  return response.json(); 
       }) 
       .then((item):void => {
  document.getElementById('idTitle')["value"]=item.Title;
  document.getElementById('idfname')["value"]=item.NameF;
  document.getElementById('idgender')["value"]=item.Gender;
  document.getElementById('idDepart')["value"]=item.Department;
  document.getElementById('idCity')["value"]=item.City;
  localStorage.setItem('ItemId', item.Id);
  
       }, (error: any): void => { 
  alert(error);
       }); 
  }
  // End Get Item By Id
  // start Clear Method of input type
  private ClearMethod()
  {
  document.getElementById('idTitle')["value"]="";
  document.getElementById('idfname')["value"]="";
  document.getElementById('idgender')["value"]="";
  document.getElementById('idDepart')["value"]="";
  document.getElementById('idCity')["value"]="";
  }
  // End Clear Method of input type
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
  PropertyPaneTextField('listname', { 
  label: strings.ListNameFieldLabel
                  }) 
                ] 
              } 
            ] 
          } 
        ] 
      }; 
    }

}
