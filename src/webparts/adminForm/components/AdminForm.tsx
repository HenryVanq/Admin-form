
import * as React from 'react';
import styles from './AdminForm.module.scss';
import { IAdminFormProps } from './IAdminFormProps';

import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { TaxonomyPicker, IPickerTerms } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";
import { DateTimePicker, DateConvention, TimeConvention } from '@pnp/spfx-controls-react/lib/dateTimePicker';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { IReactSpFxPnP } from "../Model/IReactSpFxPnP";
import { default as pnp, ItemUpdateResult, Web, Item, sp, ItemAddResult } from "sp-pnp-js";
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/components/Button';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { UrlQueryParameterCollection } from '@microsoft/sp-core-library';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPComponentLoader } from '@microsoft/sp-loader';
import * as $ from 'jquery';
import './custom.css';




var queryParms = new UrlQueryParameterCollection(window.location.href);
var idd = queryParms.getValue("idd");
var iddd = queryParms.getValue("iddd");
let siteUrl = "https://idikagr.sharepoint.com/sites/ExternalSharing";
let web = new Web(siteUrl);
let cssUrl = 'https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css';
SPComponentLoader.loadCss(cssUrl)


export default class AdminForm extends React.Component<IAdminFormProps, IReactSpFxPnP> {
  constructor(props) {
    super(props);


    this.handleRequesterEmail = this.handleRequesterEmail.bind(this);
    this.handleRequestDate = this.handleRequestDate.bind(this);
    this.handleRequest = this.handleRequest.bind(this);
    this.handleReferenceNumberIn = this.handleReferenceNumberIn.bind(this);
    this.handleReferenceNumberOut = this.handleReferenceNumberOut.bind(this);
    this.handleDate = this.handleDate.bind(this)
    this.handleVerificationCode = this.handleVerificationCode.bind(this);
    this.handleFullname = this.handleFullname.bind(this);
    this.handleOrganization = this.handleOrganization.bind(this);
    this.handlePhoneNumber = this.handlePhoneNumber.bind(this);
    this.handleDecryption = this.handleDecryption.bind(this);
    this.handleEmail = this.handleEmail.bind(this);
    this.handleReason = this.handleReason.bind(this);
    this.handleDesc = this.handleDesc.bind(this);
    this._onCheckboxChange = this._onCheckboxChange.bind(this);
    this._onRenderFooterContent = this._onRenderFooterContent.bind(this);
    this.createItem = this.createItem.bind(this);
    this.updateItem = this.updateItem.bind(this);

    //this.onTaxPickerChange = this.onTaxPickerChange.bind(this);
    this._getManager = this._getManager.bind(this);
    this.state = {
      date: "",
      request: "",
      requestDate: "",
      Fullname: "",
      Organization: "",
      PhoneNumber: "",
      Email: "",
      Reason: "",
      requesterEmail: "",
      referenceNumberIn: "",
      referenceNumberOut: "",
      referenceNumberOutDate: new Date(),
      verificationCode: "",
      decryption: "",
      description: "",
      selectedItems: [],
      hideDialog: true,
      showPanel: false,
      dpselectedItem: undefined,
      dpselectedItems: [],
      disableToggle: false,
      defaultChecked: false,
      termKey: undefined,
      userManagerIDs: [],
      pplPickerType: "",
      status: "",
      isChecked: false,
      required: "This is required",
      onSubmission: false,
      termnCond: false,


    }
  }

  componentDidMount() {

    pnp.sp.web.lists.getByTitle('SharedFiles').items.getById(1).effectiveBasePermissions.get().then(console.log)
    console.log(pnp.sp.web.getCurrentUserEffectivePermissions());

    pnp.sp.web.lists.getByTitle("Requests").items.getById(parseInt(idd)).get().then((item: any) => {
      let dateobj = new Date(item.RequestDate);
      console.log(item.ReceiverId)
      this.setState({
        request: ((item.Request == null) ? "" : item.Request),
        Fullname: ((item.Fullname == null) ? "" : item.Fullname),
        Organization: ((item.Organization == null) ? "" : item.Organization),
        PhoneNumber: ((item.PhoneNumber == null) ? "" : item.PhoneNumber),
        Email: ((item.Email == null) ? "" : item.Email),
        Reason: ((item.Reason == null) ? "" : item.Reason),
        requestDate: ((item.RequestDate == null) ? "" : dateobj.toLocaleDateString('en-GB')),
        referenceNumberIn: ((item.ReferenceNumberIn == null) ? "" : item.ReferenceNumberIn),



        requesterEmail: ((item.RequesterEmail == null) ? "" : item.RequesterEmail),
        referenceNumberOut: ((item.ReferenceNumberOut == null) ? "" : item.ReferenceNumberOut),
        //referenceNumberOutDate: ((item.ReferenceNumberOutDate == null) ? "" : item.ReferenceNumberOutDate),
        verificationCode: ((item.VerificationCode == null) ? "" : item.VerificationCode),
        decryption: ((item.Decryption == null) ? "" : item.Decryption),





      });
    });

    // let todayDateobj = new Date().toISOString().substr(0, 10);
    // $('#date').val(todayDateobj)


    $('#fileUpload').hide()
    //$('#fileUploadInput').hide()
  }

  public render(): React.ReactElement<IAdminFormProps> {
    const { dpselectedItem, dpselectedItems } = this.state;
    const { request, requestDate, requesterEmail, Fullname, Organization, PhoneNumber, Email, Reason } = this.state;
    pnp.setup({
      spfxContext: this.props.context
    });

    return (
      <form onSubmit={(e) => { return false }}>


        <div className={"card text-center bg-info mb-3"}>
          <div className={"card-header"}> <h3 className={"text-white"} id={"title"}> Φόρμα Εισαγωγής Αρχείου</h3> </div>
        </div>


        <div className="card" style={{ border: 'none' }}>
          <div className="card-header text-white bg-dark mb-3" >
            <h5> Στοιχεία Αίτησης </h5>
          </div>
          <br></br>
          <div className="form-row" >
            <div className="form-group col-md-6">
              <label><h6> Όνομα Αιτήματος </h6></label>
              <TextField className="form-control" readOnly value={this.state.request} required={true} onChanged={this.handleRequest}
                errorMessage={(this.state.request.length === 0 && this.state.onSubmission === true) ? this.state.required : ""} placeholder=" Όνομα Αιτήματος" />
            </div>
            <div className="form-group col-md-6 text-center">
              <label> <h6> Το αίτημα ολοκληρώθηκε </h6></label>
              <Toggle
                disabled={this.state.disableToggle}
                checked={this.state.defaultChecked}
                label=""
                onAriaLabel="This toggle is checked. Press to uncheck."
                offAriaLabel="This toggle is unchecked. Press to check."
                onText="Ναι"
                offText="Όχι"
                onChanged={(checked) => this._changeSharing(checked)}
                onFocus={() => console.log('onFocus called')}
                onBlur={() => console.log('onBlur called')}
              />
            </div>
          </div>
          <div className="form-row" >
            <div className="form-group col-md-6">
              <label> <h6 >Όνομα</h6></label>
              <TextField className="form-control" readOnly value={this.state.Fullname} required={true} onChanged={this.handleFullname}
                errorMessage={(this.state.Fullname.length === 0 && this.state.onSubmission === true) ? this.state.required : ""} placeholder="Όνομα "
              />
            </div>
            <div className="form-group col-md-6">
              <label> <h6>Οργανισμός</h6></label>
              <TextField className="form-control" readOnly value={this.state.Organization} required={true} onChanged={this.handleOrganization}
                errorMessage={(this.state.Organization.length === 0 && this.state.onSubmission === true) ? this.state.required : ""} placeholder="Οργανισμός"
              />
            </div>
          </div>
          <div className="form-row" >
            <div className="form-group col-md-6">
              <label> <h6 >Τηλέφωνο</h6></label>
              <TextField className="form-control" value={this.state.PhoneNumber} required={true} onChanged={this.handlePhoneNumber}
                errorMessage={(this.state.PhoneNumber.length === 0 && this.state.onSubmission === true) ? this.state.required : ""} placeholder="Τηλέφωνο" />
            </div>
            <div className="form-group col-md-6">
              <label> <h6> Email</h6></label>
              <TextField className="form-control" readOnly value={this.state.Email} required={true} onChanged={this.handleEmail}
                errorMessage={(this.state.Email.length === 0 && this.state.onSubmission === true) ? this.state.required : ""} placeholder="Email" />
            </div>
          </div>
          <br></br>
          <div className="form-group"  >
            <label><h6> Αιτιολογία </h6></label>
          </div>
          <TextField className="form-control" readOnly multiline={true} value={this.state.Reason} required={true} onChanged={this.handleReason}
            errorMessage={(this.state.Reason.length === 0 && this.state.onSubmission === true) ? this.state.required : ""} placeholder=" Αιτιολογία" />
          <br></br>
          <br></br>
          <div className="card-header text-white bg-dark mb-3" >
            <h5> Φορμα Διαχειριστή</h5>
          </div>
          <br></br>
          <div className="form-row" >
            <div className="form-group col-md-6">
              <label> <h6> Παραλήπτης </h6></label>
              <div className="form-control" id="PeoplePickerBorder">
                <PeoplePicker
                  context={this.props.context}
                  personSelectionLimit={1}
                  groupName={""} // Leave this blank in case you want to filter from all users    
                  showtooltip={true}
                  isRequired={true}
                  disabled={false}
                  ensureUser={true}
                  //selectedItems={this._getManager}
                  selectedItems={this._getManager}
                  showHiddenInUI={false}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000} />
              </div>
            </div>
            <div className="form-group col-md-6">
              <label> <h6> Ημερομηνία Αίτησης</h6></label>
              <TextField className="form-control" readOnly value={this.state.requestDate} required={true} onChanged={this.handleRequestDate}
                errorMessage={(this.state.requestDate.length === 0 && this.state.onSubmission === true) ? this.state.required : ""} placeholder="Ημερομηνία Αίτησης" />
            </div>
          </div>
          <br></br>

          <div className="form-row" >
            <div className="form-group col-md-6">
              <label> <h6 >Αρ. Πρωτ. Εισερχομένου </h6></label>
              <TextField className="form-control" value={this.state.referenceNumberIn} required={true} onChanged={this.handleReferenceNumberIn}
                errorMessage={(this.state.referenceNumberIn.length === 0 && this.state.onSubmission === true) ? this.state.required : ""} placeholder="Αρ. Πρωτ. Εισερχομένου " />
            </div>
            <div className="form-group col-md-6">
              <label> <h6> Κωδικός Επιβεβαίωσης </h6></label>
              <TextField className="form-control" value={this.state.verificationCode} required={true} onChanged={this.handleVerificationCode}
                errorMessage={(this.state.verificationCode.length === 0 && this.state.onSubmission === true) ? this.state.required : ""} placeholder="Κωδικός Επιβεβαίωσης" />
            </div>
            {/* <div className="form-group col-md-4">
              <label> <h6> Email Αιτούντα </h6></label>
              <TextField className="form-control" readOnly value={this.state.requesterEmail} required={true} onChanged={this.handleRequesterEmail}
                errorMessage={(this.state.requesterEmail.length === 0 && this.state.onSubmission === true) ? this.state.required : ""} placeholder="Email Αιτούντα " />
            </div> */}
          </div>
          <div className="form-row" >
            <div className="form-group col-md-6">
              <label> <h6> Αρ. Πρωτ. Εξερχομένου </h6></label>
              <TextField className="form-control" value={this.state.referenceNumberOut} required={true} onChanged={this.handleReferenceNumberOut}
                errorMessage={(this.state.referenceNumberOut.length === 0 && this.state.onSubmission === true) ? this.state.required : ""} placeholder="Αρ. Πρωτ. Εξερχομένου " />
            </div>
            <div className="form-group col-md-6">
              <label> <h6> Ημερομηνία Εξερχομένου </h6></label>
              <input
                id="date"
                type="date"
                className="form-control"
                name="date"
                style={{ height: '2.9em' }}
                onChange={this.handleDate}
              />
            </div>
            {/* <div className="form-group col-md-4">
              <label> <h6> Ημερομηνία Εξερχομένου </h6></label>
              <TextField className="form-control" value={this.state.referenceNumberOutDate} required={true} onChanged={this.handleReferenceNumberOutDate}
                errorMessage={(this.state.referenceNumberOutDate.length === 0 && this.state.onSubmission === true) ? this.state.required : ""} placeholder=" Ημερομηνία Εξερχομένου" />
            </div> */}
          </div>

          <br></br>
          <div className="form-row" >

            <div className="form-group col-md-6 ">
              <label> <h6> Κωδικός Αποκρυπτογράφησης </h6></label>
              <div className="input-group ">
                <PrimaryButton className="btn btn-dark text-white btn-sm" id="buttonGenPass" text="Δημιουργία" onClick={() => {
                  this.generatePassword();
                }} />

                <TextField id="inputGenPass" className="form-control border-0" value={this.state.decryption} required={true} onChanged={this.handleDecryption}
                  errorMessage={(this.state.decryption.length === 0 && this.state.onSubmission === true) ? this.state.required : ""} placeholder="Κωδικός Αποκρυπτογράφησης" />

              </div>
            </div>

            <div className="form-group col-md-6">
              <label> <h6> Ανέβασμα Αρχείου </h6></label>
              <input required={true} className="form-control" type='file' id='fileUploadInput' name='myfile' />
              <button id="fileUpload" name="uFile">upload</button>
            </div>



          </div>
        </div>
        <br></br>
        {/* <PrimaryButton text="Create" onClick={() => { this.validateForm(); }} /> */}
        <PrimaryButton id="btnForm" className="btn btn-primary btn-lg btn-block" onClick={() => {
          if ($('#fileUploadInput').val() != ""

            && this.state.userManagerIDs.length >= 1
            && this.state.referenceNumberOut != ""
            && this.state.referenceNumberIn
            && this.state.verificationCode != ""
            && this.state.referenceNumberOutDate != null
            && this.state.decryption != ""
          ) {

            if ($('#date').val() === "") {
              alert("Παρακαλω επιλέξτε την ημερομηνία")
            } else {

              this.updateItem();
              this.uploadingFileEventHandlers();
            }


          } else {
            alert("Παρακαλώ συμπληρώστε όλα τα πεδία")
          }
        }} style={{ marginRight: '8px' }}><h5> Υποβολή Αιτήματος </h5> </PrimaryButton>
        <DefaultButton id="btnForm" className="btn btn-secondary btn-lg btn-block" onClick={() => { this.setState({}); }}  > <h5> Ακύρωση </h5> </DefaultButton>
      </form>
    );
  }



  private todayDate() {

    let today = new Date().toISOString().substr(0, 10);
    $('#today').val('dd-mm-yyy')


  }

  private uploadingFileEventHandlers() {

    let fileUpload = document.getElementById("fileUpload")
    let test1 = document.getElementById("fileUploadInput")

    if (fileUpload) {
      this.uploadFiles(test1);
      console.log("file uploaded: ")
    }
  }

  protected uploadFiles(fileUpload) {

    this.createFile();

    // var dat = new Date(this.state.requestDate);
    // var day = dat.getDate();
    // var mon = dat.getMonth();
    // var yar = dat.getFullYear();
    var day = this.state.requestDate.charAt(0) + this.state.requestDate.charAt(1)
    var mon = this.state.requestDate.charAt(3) + this.state.requestDate.charAt(4)
    var yar = this.state.requestDate.charAt(6) + this.state.requestDate.charAt(7) + this.state.requestDate.charAt(8) + this.state.requestDate.charAt(9)


    let file = fileUpload.files[0];
    //let attachmentsArray = this.state.attachmentsToUpload;        

    if (file.size <= 10485760) {
      // small upload
      web.getFolderByServerRelativeUrl("/sites/ExternalSharing/SharedFiles/" + this.state.request + "-" + day + "-" + mon + "-" + yar)
        .files.add(file.name, file, true)
        .then(_ => console.log("done"));

    } else { // large upload
      web.getFolderByServerRelativeUrl("/sites/ExternalSharing/SharedFiles/" + this.state.request + "-" + day + "-" + mon + "-" + yar)
        .files
        .addChunked(file.name, file, data => {

        }, true)
        .then(_ => console.log("done!"));

    }

    pnp.sp.web.lists.getByTitle("Files").items.add({
      FileName: file.name,
      Path: "https://idikagr.sharepoint.com/sites/ExternalSharing/_layouts/download.aspx?sourceurl=/sites/ExternalSharing/SharedFiles/" + this.state.request + "-" + day + "-" + mon + "-" + yar + "/" + file.name,
      RequestId: parseInt(idd),
      ReferenceNumberIn: this.state.referenceNumberIn,
      ReferenceNumberOut: this.state.referenceNumberOut,
      ReferenceNumberOutDate: this.state.date.toString(),
      VerificationCode: this.state.verificationCode,
      Decryption: this.state.decryption,
      // refDateOut: this.state.date.toString()
    }).then((iar: ItemAddResult) => {
      console.log(iar);
    }).then(() => {
      alert("H καταχώρηση ολοκληρώθηκε επιτυχώς!")
      //window.location.href = 'https://idikagr.sharepoint.com/sites/ExternalSharing/SitePages/HomePage.aspx'
    })
  }

  private generatePassword() {
    var length = 8,
      charset = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789",
      retVal = "";
    for (var i = 0, n = charset.length; i < length; ++i) {
      retVal += charset.charAt(Math.floor(Math.random() * n));
    }
    this.setState({ decryption: retVal });
  }

  private _getManager(items: any[]) {
    this.state.userManagerIDs.length = 0;
    let tempuserMngArr = [];
    for (let item in items) {
      tempuserMngArr.push(items[item].id);
    }
    this.setState({ userManagerIDs: tempuserMngArr });
  }

  private _onRenderFooterContent = (): JSX.Element => {
    return (
      <div>

        <DefaultButton onClick={this._onClosePanel}>Cancel</DefaultButton>
      </div>
    );
  }

  private _log(str: string): () => void {
    return (): void => {
      console.log(str);
    };
  }

  protected createFile() {
    //create folder
    var dat = new Date(this.state.requestDate);
    console.log('reDate: ' + this.state.requestDate)

    var day = this.state.requestDate.charAt(0) + this.state.requestDate.charAt(1)
    var mon = this.state.requestDate.charAt(3) + this.state.requestDate.charAt(4)
    var yar = this.state.requestDate.charAt(6) + this.state.requestDate.charAt(7) + this.state.requestDate.charAt(8) + this.state.requestDate.charAt(9)

    console.log('/sites/ExternalSharing/SharedFiles/' + this.state.request + "-" + day + "-" + mon + "-" + yar);
    web
      .folders
      //.add('/sites/IDIKA/Shared%20Documents/' + filename + "-" + day + "-" + mon + "-" + yar)
      .add('/sites/ExternalSharing/SharedFiles/' + this.state.request + "-" + day + "-" + mon + "-" + yar)
      .then(console.log);
  }

  private _onClosePanel = () => {
    this.setState({ showPanel: false });
  }

  private _onShowPanel = () => {
    this.setState({ showPanel: true });
  }

  private _changeSharing(checked: any): void {
    this.setState({ defaultChecked: checked });
  }

  private _changeState = (item: IDropdownOption): void => {
    console.log('here is the things updating...' + item.key + ' ' + item.text + ' ' + item.selected);
    this.setState({ dpselectedItem: item });
    if (item.text == "Employee") {
      this.setState({ defaultChecked: false });
      this.setState({ disableToggle: true });
    }
    else {
      this.setState({ disableToggle: false });
    }
  }

  private handleFullname(value: string): void {
    return this.setState({
      Fullname: value
    });
  }

  private handleOrganization(value: string): void {
    return this.setState({
      Organization: value
    });
  }

  private handlePhoneNumber(value: string): void {
    return this.setState({
      PhoneNumber: value
    });
  }

  private handleEmail(value: string): void {
    return this.setState({
      Email: value
    });
  }

  private handleReason(value: string): void {
    return this.setState({
      Reason: value
    });
  }

  private handleRequest(value: string): void {
    return this.setState({
      request: value
    });
  }

  private handleRequestDate(value: string): void {
    return this.setState({
      requestDate: value
    });
  }
  private handleRequesterEmail(value: string): void {
    return this.setState({
      requesterEmail: value
    });
  }
  private handleReferenceNumberIn(value: string): void {
    return this.setState({
      referenceNumberIn: value
    });
  }
  private handleReferenceNumberOut(value: string): void {
    return this.setState({
      referenceNumberOut: value
    });
  }
  // private handleReferenceNumberOutDate(value: Date): void {
  //   return this.setState({
  //     referenceNumberOutDate: value
  //   });
  // }

  private handleDate(e) {
    return this.setState({
      date: e.target.value,
    })

  }
  private handleVerificationCode(value: string): void {
    return this.setState({
      verificationCode: value
    });
  }
  private handleDecryption(value: string): void {
    return this.setState({
      decryption: value
    });
  }

  private handleDesc(value: string): void {
    return this.setState({
      description: value
    });
  }

  private _onCheckboxChange(ev: React.FormEvent<HTMLElement>, isChecked: boolean): void {
    console.log(`The option has been changed to ${isChecked}.`);
    this.setState({ termnCond: (isChecked) ? true : false });
  }

  private _closeDialog = (): void => {
    this.setState({ hideDialog: true });
  }

  private _showDialog = (status: string): void => {
    this.setState({ hideDialog: false });
    this.setState({ status: status });
  }

  /**
   * A sample to show on how form can be validated
   */
  private validateForm(): void {
    let allowCreate: boolean = true;
    this.setState({ onSubmission: true });

    if (this.state.request.length === 0) {
      allowCreate = false;
    }
    // if (this.state.termKey === undefined) {
    //   allowCreate = false;
    // }

    if (allowCreate) {
      this._onShowPanel();
    }
    else {
      //do nothing
    }
  }

  private createItem(): void {
    this._onClosePanel();
    this._showDialog("Submitting Request");
    pnp.sp.web.lists.getByTitle("Employee Registeration").items.add({
      Title: this.state.request,
      Description: this.state.description,
      Department: this.state.dpselectedItem.key,
      Projects: {
        __metadata: { "type": "SP.Taxonomy.TaxonomyFieldValue" },
        Label: "1",
        TermGuid: this.state.termKey,
        WssId: -1
      },
      Reporting_x0020_ManagerId: this.state.userManagerIDs[0]
    }).then((iar: ItemUpdateResult) => {
      console.log(iar);
      this.setState({ status: "Your request has been submitted sucessfully." });
    });
  }

  private updateItem(): void {
    var checkboxValue = this.state.defaultChecked ? "Yes" : "No";
    console.log(this.state.defaultChecked);
    pnp.sp.web.lists.getByTitle("Requests").items.getById(parseInt(idd)).update({
      //'Title': `Updated Item ${new Date()}`
      // ReferenceNumberIn: this.state.referenceNumberIn,
      // ReferenceNumberOut: this.state.referenceNumberOut,
      // ReferenceNumberOutDate: this.state.referenceNumberOutDate,
      // VerificationCode: this.state.verificationCode,
      // Decryption: this.state.decryption,
      ReceiverId: this.state.userManagerIDs[0],
      ReferenceNumberIn: this.state.referenceNumberIn,
      Completed: checkboxValue
    }).then((iar: ItemUpdateResult) => {
      console.log(iar);
      this.setState({ status: "Your request has been submitted sucessfully." });
      console.log("new record added to files lists")
    });
  }

  // private updateItem(): void {  
  //   this.updateStatus('Loading latest items...');  
  //   let latestItemId: number = undefined;  
  //   let etag: string = undefined;  

  //   this.getLatestItemId()  
  //     .then((itemId: number): Promise<Item> => {  
  //       if (itemId === -1) {  
  //         throw new Error('No items found in the list');  
  //       }  

  //       latestItemId = itemId;  
  //       this.updateStatus(`Loading information about item ID: ${itemId}...`);  
  //       return sp.web.lists.getByTitle(this.properties.listName)  
  //         .items.getById(itemId).get(undefined, {  
  //           headers: {  
  //             'Accept': 'application/json;odata=minimalmetadata'  
  //           }  
  //         });  
  //     })  
  //     .then((item: Item): Promise<IListItem> => {  
  //       etag = item["odata.etag"];  
  //       return Promise.resolve((item as any) as IListItem);  
  //     })  
  //     .then((item: IListItem): Promise<ItemUpdateResult> => {  
  //       return sp.web.lists.getByTitle(this.properties.listName)  
  //         .items.getById(item.Id).update({  
  //           'Title': `Updated Item ${new Date()}`  
  //         }, etag);  
  //     })  
  //     .then((result: ItemUpdateResult): void => {  
  //       this.updateStatus(`Item with ID: ${latestItemId} successfully updated`);  
  //     }, (error: any): void => {  
  //       this.updateStatus('Loading latest item failed with error: ' + error);  
  //     });  
  // } 

  private gotoHomePage(): void {
    window.location.replace(this.props.siteUrl);
  }
}


