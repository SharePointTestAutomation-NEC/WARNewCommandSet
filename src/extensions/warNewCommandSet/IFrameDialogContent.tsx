import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import { DialogContent } from 'office-ui-fabric-react';
import * as jquery from 'jquery';

interface IIFrameDialogContentProps {
    close: () => void;
    url: string;
    iframeOnLoad?: (iframe: any) => void;
}
class IFrameDialogContent extends React.Component<IIFrameDialogContentProps, {} > {
    private iframe: any;
    private checkflag:boolean;
    constructor(props) {
        super(props);
    }
    public render(): JSX.Element {
       return <DialogContent
            title='New Item'
            onDismiss={this.props.close}
            showCloseButton={true}
        >
        <div id="loader1">
            <table style={{position: 'relative',left: '30%'}}>
                <tr>
                    <td><img  src={require('./assets/loadingForm.gif')} width='150' height='150' alt='loading gif'/></td>
                </tr>
                <tr>
                    <td><span style={{position: 'relative',fontSize:'40px',color:'#0072c6'}}>Working on it...</span></td>
                </tr>
            </table>
         </div>
         <div dangerouslySetInnerHTML={{__html: ''}} />
        
        <iframe ref={(iframe) => { this.iframe = iframe; }} onLoad={this._iframeOnLoad.bind(this)}
           src={this.props.url} frameBorder={0} style={{width: '620px', height: '150px'}}>
        </iframe>
    </DialogContent>;
    }

    private _iframeOnLoad(): void {

        jquery('#loader1').hide();
        this.IncreaseSize();

        if(this.checkflag) {

            this.props.close();
            window.location.reload();
        }
        
        this.checkflag = true;
        try {   
            
            var y=this.iframe.contentWindow.$("input[id$='idIOGoBack']")[1];
            y.setAttribute('onclick',null);
            y.onclick=this.CloseDialog.bind(this);
        
            var scollH=this.iframe.contentWindow.document.getElementById('s4-workspace');
            scollH.setAttribute("style", "overflow-x: hidden;");
            
            var SaveButton=this.iframe.contentWindow.document.querySelectorAll('input[value="Save"]')[1];
            var AttachmentButton=this.iframe.contentWindow.document.getElementById('AddAttachments');
            var AttachFile=this.iframe.contentWindow.document.getElementById('attachOKbutton');
           
            addEvent(SaveButton, "click", this.IncreaseSize.bind(this) );
            addEvent(AttachmentButton, "click", this.IncreaseSize.bind(this));
            addEvent(AttachFile, "click", this.IncreaseSize.bind(this) );
            
        } catch (err) {
            if (err.name !== 'SecurityError') {
                throw err;
            }
        }
        if (this.props.iframeOnLoad) {
            this.props.iframeOnLoad(this.iframe);
        }
    }

    public IncreaseSize() {
        this.iframe.style.height = this.iframe.contentWindow.document.body.scrollHeight + 180 + 'px';
    }
    public onClick() {
        this.props.close();
    }
    public CloseDialog() {
        this.props.close();
    }
}
function addEvent(obj, evType, fn) {
    if (obj.addEventListener) {
        obj.addEventListener(evType, fn, false);
        return true;
    } else if (obj.attachEvent) {
        var r = obj.attachEvent("on" + evType, fn);
        return r;
    } else {
        alert("Handler could not be attached");
    }
}
export default class IFrameDialog extends BaseDialog {
    
    constructor(private url: string) {
        //super();
        super({isBlocking: true});
    }
    public render(): void {
        window.addEventListener('CloseDialog', () => { this.close(); });
        ReactDOM.render(
            <IFrameDialogContent
                close={this.close}
                url={this.url}
            />, this.domElement);
    }
    public getConfig(): IDialogConfiguration {
        return {
            isBlocking: false
        };
    }
}