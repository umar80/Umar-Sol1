import {BaseDialog,IDialogShowOptions,IDialogConfiguration} from "@microsoft/sp-dialog";

export default class InjioDialog extends BaseDialog {
    public message: string;
    public colorCode: string;
   
    public render(): void {     
      this.domElement.innerHTML=`
        <div>
        HELLO WORLD!
            //<iframe width="500px" height="500px" src="https://webvine.sharepoint.comhttps://webvine.sharepoint.com/Lists/ServiceComments/NewForm.aspx?IsDlg=1Lists/ServiceComments/NewForm.aspx"/>
        </div>`; 
    }
   
    public getConfig(): IDialogConfiguration {
      return {
        isBlocking: true
      };
    }    
   }