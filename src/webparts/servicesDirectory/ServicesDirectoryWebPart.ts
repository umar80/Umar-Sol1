

import { Version ,Environment,
  EnvironmentType} from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider
} from '@microsoft/sp-webpart-base';

import { SPComponentLoader } from '@microsoft/sp-loader';



import { BaseDialog, IDialogConfiguration, Dialog } from '@microsoft/sp-dialog';

import { escape } from '@microsoft/sp-lodash-subset';
import {
  SPHttpClient,
  SPHttpClientResponse   
 } from '@microsoft/sp-http';

import styles from './ServicesDirectoryWebPart.module.scss';
import * as strings from 'ServicesDirectoryWebPartStrings';
import pnp, { CamlQuery, sp } from 'sp-pnp-js';
import {ServiceDirectory,ServiceDirectorys} from './ServiceDirectoryList';
import MockHttpClient from './MockHttpClient';
import RatingObj from './RatingObj';
import InjioDialog from './InjioDialog'

import * as jQuery from 'jquery';

export interface IServicesDirectoryWebPartProps {
  description: string;
  pageSize:number
}

export default class ServicesDirectoryWebPart extends BaseClientSideWebPart<IServicesDirectoryWebPartProps> {

  private _numberOfRecords=0;
  private _searchByName=false;
  private _currentQuery="";
  private _detailedPage ="";
  

  protected onInit():Promise<void>{
    return super.onInit().then(_ => {
      pnp.setup({
        spfxContext: this.context
      });
    });
  }

  private _loadSPJSOMScripts() {
 
    const siteColUrl = this.context.pageContext.web.absoluteUrl;
    console.log(siteColUrl);
    try {
      SPComponentLoader.loadScript(siteColUrl + '/_layouts/15/init.js', {
        globalExportsName: '$_global_init'
      })/*
        .then((): Promise<{}> => {
          return SPComponentLoader.loadScript(siteColUrl + '/_layouts/15/MicrosoftAjax.js', {
            globalExportsName: 'Sys'
          });
        })
        .then((): Promise<{}> => {
          return SPComponentLoader.loadScript(siteColUrl + '/_layouts/15/SP.Runtime.js', {
            globalExportsName: 'SP'
          });
        })
        .then((): Promise<{}> => {
          return SPComponentLoader.loadScript(siteColUrl + '/_layouts/15/SP.js', {
            globalExportsName: 'SP'
          });
        })
        .then((): Promise<{}> => {
          return SPComponentLoader.loadScript(siteColUrl + '/_layouts/15/reputation.js', {
            globalExportsName: 'SP'
          });
        })
        .then((): Promise<{}> => {
          return SPComponentLoader.loadScript('https://ajax.googleapis.com/ajax/libs/jquery/3.3.1/jquery.min.js', {
            globalExportsName: 'jQuery'
          });
        })
        .then((): Promise<{}> => {
          return SPComponentLoader.loadScript(siteColUrl + '/siteassets/jquery.rateit.min.js', {            
            globalExportsName: 'jQuery',            
          });
        })  */      
        /*.then((): void => {          
          SPComponentLoader.loadCss(siteColUrl + '/siteassets/rateit.css');
          jQuery('rateit').rateit();
        })*/
        .catch((reason: any) => {         
          
        });
    } catch (error) {

    }
  }

  private loadSP() : Promise<any> {
    var globalExportsName = null, p = null;
    var promise = new Promise<any>((resolve, reject) => {
      globalExportsName = '$_global_init'; p = (window[globalExportsName] ? Promise.resolve(window[globalExportsName]) : SPComponentLoader.loadScript('/_layouts/15/init.js', { globalExportsName }));
      p.catch((error) => { })
        .then(($_global_init): Promise<any> => {
          globalExportsName = 'Sys'; p = (window[globalExportsName] ? Promise.resolve(window[globalExportsName]) : SPComponentLoader.loadScript('/_layouts/15/MicrosoftAjax.js', { globalExportsName }));
          return p;
        }).catch((error) => { })
        .then((Sys): Promise<any> => {
          globalExportsName = 'Sys'; p = ((window[globalExportsName] && window[globalExportsName].ClientRuntimeContext) ? Promise.resolve(window[globalExportsName]) : SPComponentLoader.loadScript('/_layouts/15/ScriptResx.ashx?name=sp.res&culture=en-us', { globalExportsName }));
          return p;
        })/*.catch((error) => { })
        .then((Sys): Promise<any> => {
          globalExportsName = 'SP-Runtime'; p = ((window[globalExportsName] && window[globalExportsName].ClientRuntimeContext) ? Promise.resolve(window[globalExportsName]) : SPComponentLoader.loadScript('/_layouts/15/SP.Runtime.js', { globalExportsName }));
          return p;
        })*/.catch((error) => { })
        .then((SP): Promise<any> => {
          globalExportsName = 'SP'; p = ((window[globalExportsName] && window[globalExportsName].ClientContext) ? Promise.resolve(window[globalExportsName]) : SPComponentLoader.loadScript('/_layouts/15/SP.js', { globalExportsName }));
          return p;
        })
        .then((Sys): Promise<any> => {
          globalExportsName = 'SP-Init'; p = ((window[globalExportsName] && window[globalExportsName].ClientRuntimeContext) ? Promise.resolve(window[globalExportsName]) : SPComponentLoader.loadScript('/_layouts/15/SP.Init.js', { globalExportsName }));
          return p;
        }).catch((error) => { })
        .then((Sys): Promise<any> => {
          globalExportsName = 'SP-UI-Dialog'; p = ((window[globalExportsName] && window[globalExportsName].ClientRuntimeContext) ? Promise.resolve(window[globalExportsName]) : SPComponentLoader.loadScript('/_layouts/15/SP.UI.Dialog.js', { globalExportsName }));
          return p;
        }).catch((error) => { })
        .then((Sys): Promise<any> => {
          globalExportsName = 'RP'; p = ((window[globalExportsName] && window[globalExportsName].ClientRuntimeContext) ? Promise.resolve(window[globalExportsName]) : SPComponentLoader.loadScript('/_layouts/15/reputation.js', { globalExportsName }));
          return p;
        }).catch((error) => { })
        .then((Sys): Promise<any> => {
          globalExportsName = 'SP-ClientTemplates'; p = ((window[globalExportsName] && window[globalExportsName].ClientRuntimeContext) ? Promise.resolve(window[globalExportsName]) : SPComponentLoader.loadScript('/_layouts/15/clienttemplates.js', { globalExportsName }));
          return p;
        }).catch((error) => { })
        .then((Sys): Promise<any> => {
          globalExportsName = 'SP-UI-Reputation-Debug'; p = ((window[globalExportsName] && window[globalExportsName].ClientRuntimeContext) ? Promise.resolve(window[globalExportsName]) : SPComponentLoader.loadScript('/_layouts/15/SP.UI.Reputation.debug.js', { globalExportsName }));
          return p;
        }).catch((error) => { })
        .then((SP) => {
          resolve(SP);
        });
    });
    return promise;
  }

  protected setupLookups():Promise<void>{
    return pnp.sp.web.lists.getByTitle("Services Directory").fields.get().
    then(
      (response)=> {
      return response;
    });
  }
  private _getMockListData(): Promise<ServiceDirectorys> {
    return MockHttpClient.get()
      .then((data: ServiceDirectory[]) => {
        var listData: ServiceDirectorys={ value: data };
        return listData;
      }) as Promise<ServiceDirectorys>;
  }

  public render(): void {

    if (Environment.type === EnvironmentType.Local) {
      this._getMockListData().then((response) => {
        this._renderList(response.value);
      });
    }
    else if (Environment.type == EnvironmentType.SharePoint || 
              Environment.type == EnvironmentType.ClassicSharePoint) {
      this.loadSP();
      this.setupLookups().then(
          (response) => {      
         // this._getFilterItems("");     
          this._renderList(response);
          //this._loadSPJSOMScripts();
        });
    }
  }

  protected _renderList(respose){

    let alphabetHTML=`<div title="{0}" href="/sites/IngiyoLight/StaffDirectory/Pages/Search.aspx?k=firstname:{0}*" style="font-size: 25px;">{0}</div>`;
    let alphabet:string="A";
    let i=0;
    let innerHTML=""; 
    
    console.log(respose);
    let countryList=[];
    let stateList=[];
    let regionList=[];
    let centreList=[];
    let serviceTypeList=[];

    let countryListHTML="";
    let stateListHTML="";
    let regionListHTML="";
    let centreListHTML="";
    let serviceTypeListHTML="";

    for(i=0;i<respose.length;i++){  
      let j=0;
      if(respose[i].StaticName=="Country"){
        for(j=0;j<respose[i].Choices.length;j++){
          countryList.push(respose[i].Choices[j]);
         }
      }
      if(respose[i].StaticName=="State"){
        for(j=0;j<respose[i].Choices.length;j++){
          stateList.push(respose[i].Choices[j]);
        }
      }
      if(respose[i].StaticName=="Region"){
        for(j=0;j<respose[i].Choices.length;j++){
          regionList.push(respose[i].Choices[j]);
        }
      }
      if(respose[i].StaticName=="Centre"){
        for(j=0;j<respose[i].Choices.length;j++){
          centreList.push(respose[i].Choices[j]);
        }
      }
      if(respose[i].StaticName=="ServiceType"){
        for(j=0;j<respose[i].Choices.length;j++){
          serviceTypeList.push(respose[i].Choices[j]);
        }
      }
    }

    for(i=0;i<countryList.length;i++){      
      countryListHTML+=`<option>${countryList[i]}</option>`;
    }
    for(i=0;i<stateList.length;i++){      
      stateListHTML+=`<option>${stateList[i]}</option>`;
    }
    for(i=0;i<regionList.length;i++){      
      regionListHTML+=`<option>${regionList[i]}</option>`;
    }
    for(i=0;i<centreList.length;i++){      
      centreListHTML+=`<option>${centreList[i]}</option>`;
    }
    for(i=0;i<serviceTypeList.length;i++){      
      serviceTypeListHTML+=`<option>${serviceTypeList[i]}</option>`;
    }
    for(i=0;i<26;i++){
      let character=String.fromCharCode(65+i);
      innerHTML+=`<a
          title="${character}" 
          class="${styles.alphabet} alphabet"
          href="#">
          ${character}
          </a>`;
    }
    this.domElement.innerHTML = `
    <div class=${styles.servicesDirectory}>      
      <div class=${styles.servicesFilter}>
        <select class="country">
          <option value="0">Country</option>
          ${countryListHTML}
        </select> 
        <select class="state">
          <option value="0">State</option>
          ${stateListHTML}
        </select> 
        <select class="region">
          <option value="0">Region</option>
          ${regionListHTML}
        </select> 
        <select class="centre">
          <option value="0">Centre</option>
          ${centreListHTML}
        </select> 
        <select class="serviceType">
          <option value="0">Service Type</option>
          ${serviceTypeListHTML}
        </select> 
        <input type="text" value="" class="searchByName"/>
        <input type="button" class="${styles.Filter}" value="Filter" id="btnFilter" />
        <input type="button" class="${styles.clearFilter}" value="Clear Filter" id="btnClearFilter"/>
      </div>
      <div class="${styles.AlphabetSet}">
        ${innerHTML}
      </div>
      <div class="${styles.backing2}">
        <select>
          <option value="1">Bad</option>
          <option value="2">OK</option>
          <option value="3">Great</option>
          <option value="4">Excellent</option>
        </select>
      </div>
      <div class="${styles.results}" id="divResults">
      </div>        

      <div class="rateit" data-rateit-backingfld="#backing2" data-rateit-min="0"></div>
    </div>    
    `;

    this.domElement.querySelector('#btnFilter').addEventListener('click', () => {       
      this._numberOfRecords=0; 
      
      this._getFilterItems("");                  
    });  
      
    const alphabetElements= this.domElement.querySelectorAll('.alphabet');
    let   _lastSelectedAlphabet:any  ;
    for(let i=0;i<alphabetElements.length;i++)
    {
      alphabetElements[i].addEventListener('click', (event) => {         
        this._numberOfRecords=0;   
          
        console.log(_lastSelectedAlphabet);
        alphabetElements[i].innerHTML = `<a title="${alphabetElements[i].attributes["title"].value}" class="${styles.alphabetHover}" href="#">`+alphabetElements[i].attributes["title"].value+`</a>`;
        /*
           if(_lastSelectedAlphabet != undefined) 
           {
            if(_lastSelectedAlphabet != alphabetElements[i])
              {
                alphabetElements[i].innerHTML = `<a title="${_lastSelectedAlphabet.attributes["title"].value}" class="${styles.alphabet}" href="#">`+_lastSelectedAlphabet.attributes["title"].value+`</a>`;

              }
              else{
              alphabetElements[i].innerHTML = `<a title="${alphabetElements[i].attributes["title"].value}" class="${styles.alphabetHover}" href="#">`+alphabetElements[i].attributes["title"].value+`</a>`;
              }
          }
          else{alphabetElements[i].innerHTML = `<a title="${alphabetElements[i].attributes["title"].value}" class="${styles.alphabetHover}" href="#">`+alphabetElements[i].attributes["title"].value+`</a>`;}
          _lastSelectedAlphabet = alphabetElements[i];
          */

        this._getFilterItems(event.srcElement.attributes["title"].value); 
      });
    }
    
       this.domElement.querySelector('#btnClearFilter').addEventListener('click', () => { this._openDialog(); });   



    this._getFilterItems("");     
  }

  protected _openDialog():void{    
   // let dialog=new InjioDialog({isBlocking:true});    
    //dialog.show();
      console.log("clear clicked");
  }




  protected _getFilterItems(alphabet):void{
    if (Environment.type === EnvironmentType.Local) {
      this._getMockListData().then((response) => {
        this.renderServiceDirectoryItems(response.value);
      });
    }
    else if (Environment.type == EnvironmentType.SharePoint || 
              Environment.type == EnvironmentType.ClassicSharePoint) {
      let whereClause="";
      let _query="";
      if( alphabet==""){
        let queryArray=[];
        if(jQuery('.country').val()!="0"){
          queryArray.push(`<Eq><FieldRef Name='Country'/><Value Type='Text'>${jQuery('.country').val()}</Value></Eq>`);
        }
        if(jQuery('.state').val()!="0"){
          queryArray.push(`<Eq><FieldRef Name='State'/><Value Type='Text'>${jQuery('.state').val()}</Value></Eq>`);          
        }
        if(jQuery('.region').val()!="0"){
          queryArray.push(`<Eq><FieldRef Name='Region'/><Value Type='Text'>${jQuery('.region').val()}</Value></Eq>`);
        }
        if(jQuery('.centre').val()!="0"){
          queryArray.push(`<Eq><FieldRef Name='Centre'/><Value Type='Text'>${jQuery('.centre').val()}</Value></Eq>`);
        }
        if(jQuery('.serviceType').val()!="0"){
          queryArray.push(`<Eq><FieldRef Name='ServiceType'/><Value Type='Text'>${jQuery('.serviceType').val()}</Value></Eq>`);
        }      
        if(jQuery('.searchByName').val()!=""){
          queryArray.push(`<Contains><FieldRef Name='Title'/><Value Type='Text'>${jQuery('.searchByName').val()}</Value></Contains>`);
        }
        
       
        if(queryArray.length==1){
          whereClause =queryArray[0];
        }
        else{		
          for(let i=0;i<	queryArray.length; i++){
            if(i==0){				
              whereClause="<And>"+queryArray[i]+queryArray[i+1]+"</And>";
              i++;
            }
            else{
              whereClause="<And>"+queryArray[i]+whereClause+"</And>";
            }			
          }        
        }
      }
      else{
        whereClause=`<BeginsWith><FieldRef Name='Title'/><Value Type='Text'>${alphabet}</Value></BeginsWith>`;
      }
      this._currentQuery=alphabet;

      _query=`
        <View>
          <ViewFields>
                <FieldRef Name='ID' />
                <FieldRef Name='Title' />
                <FieldRef Name='Description' />
                <FieldRef Name='LocationMap'/>
                <FieldRef Name='ServiceType'/>
                <FieldRef Name='Website'/>
                <FieldRef Name='AverageRating'/>
                <FieldRef Name='Phone'/>
                <FieldRef Name='Logo'/>
                <FieldRef Name='Contact'/>
          </ViewFields>
          <Query>
            <Where>${whereClause}</Where>
          <OrderBy><FieldRef Name='Title' /></OrderBy></Query>
        <RowLimit>${this.properties.pageSize+this._numberOfRecords}</RowLimit></View>`;
        
      this.getFilterItems(_query)
        .then((response) => {          
         
          this.renderServiceDirectoryItems(response);          
        });
    }
  } 

  protected getFilterItems(xml:string):Promise<ServiceDirectory[]>{
    const q:CamlQuery={
      ViewXml:xml
    };
    return pnp.sp.web.lists.getByTitle("Services Directory").getItemsByCAMLQuery(q).then((response)=>{      
      console.log("Query Response");
      console.log(response);
      return response;    
    });
  };

  protected renderServiceDirectoryItems(response){
    
    let serviceDirectoryHTML="";
    let i=0;
    serviceDirectoryHTML+=`<div class="${styles.resultDiv}">`;
    console.log(response);

    
    for(i=0;i<response.length;i++){      
      serviceDirectoryHTML+=`<div class="${styles.item}">

        <div class="${styles.contentOuterDiv}">
          <div class="${styles.logo}">${response[i].Logo}</div>
          <div class="${styles.servicetitle}"><a href="https://webvine.sharepoint.com/sites/IngiyoLight/ServicesDirectory/Pages/DetailedPage.aspx?SID=${response[i].ID}" target="_Blank">${response[i].Title}</a></div>
          <br/>
          <div>
            <div class="${styles.serviceType}">Service Type: ${response[i].ServiceType}</div>
          </div>
          <div>
            <div class="${styles.serviceWebSite}">Website: ${response[i].Website?'<a href="'+response[i].Website.Url+'">'+response[i].Website.Description+'</a>':''} </div>
          </div>
          <div>
            <div class="${styles.serviceRating}">Rating: ${response[i].AverageRating} <input type="button"  value="R" id="btnRating" /></input></div>
          </div>          
        </div>
        <hr/>
        <div class="${styles.contentOuterDiv}">
          <div class="${styles.serviceContact}">Contact: ${response[i].Contact}</div>
          <div class="${styles.serviceEmail}">Email: ${response[i].Email != null || response[i].Email !=undefined ? response[i].Email : "-" }</div>
          <div class="${styles.servicePhone}">Phone: ${response[i].Phone != null?response[i].Phone:"-"}</div>
        </div>      
        </div>`;
    }
    if(response.length>this._numberOfRecords && (response.length%this.properties.pageSize==0)){
      serviceDirectoryHTML+=`<div><a class="${styles.showMore} showMore" href="#">Show more</a></div>`;
    }

    serviceDirectoryHTML+=`</div>`;
    this.domElement.querySelector('#divResults').innerHTML=`${serviceDirectoryHTML}`;
    this.domElement.querySelector('#btnRating').addEventListener('click', () => {this.myRating(1,2);});
    if(this.domElement.querySelector('a.showMore'))
      this.domElement.querySelector('a.showMore').addEventListener('click',()=>{
      this._getFilterItems(this._currentQuery);
    });
    this._numberOfRecords=response.length;
    //console.log(serviceDirectoryHTML);
  }

  protected _getRating():void{
    console.log("GET RATING");
}

protected myRating(itemId, value)
{
  console.log("Set Rating Functionality is Called");
  let spCurrentContext =  SP.ClientContext.get_current();
    
  console.log(spCurrentContext);

  let _listId:any = "E1740B60-6B6D-410F-803B-5EFDE083103B";
    console.log(spCurrentContext);
    console.log(this.context);
    console.log(_listId);
    console.log(itemId);
    console.log(value);
    
    //var rating = setRating(spCurrentContext,_listId,itemId,value);
 //   var rep = new Microsoft.Office.Server.ReputationModel.Reputation();

   //console.log(rep);
    //var rating = Microsoft.Office.Server.ReputationModel.Reputation.setLike(spCurrentContext,_listId,itemId,value);
    Microsoft.Office.Server.ReputationModel.Reputation.setLike(spCurrentContext,_listId,itemId,true);
    
    //Microsoft.Office.Server.ReputationModel.Reputation.setRating(spCurrentContext, _listId, itemId, value);

    console.log("");
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
                PropertyPaneSlider('pageSize',{
                    min:1,
                    max:30,
                    label:"Page Size",
                    showValue:true,  
                    value:1                
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
