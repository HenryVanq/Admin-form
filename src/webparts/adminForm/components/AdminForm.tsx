import * as React from 'react';
// import styles from './AdminForm.module.scss';
import { IAdminFormProps } from './IAdminFormProps';

import { TextField } from 'office-ui-fabric-react/lib/TextField';
// import { TaxonomyPicker, IPickerTerms } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";
// import { DateTimePicker, DateConvention, TimeConvention } from '@pnp/spfx-controls-react/lib/dateTimePicker';
// import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { IReactSpFxPnP } from "../Model/IReactSpFxPnP";
import { default as pnp, ItemUpdateResult, Web, Item, sp, ItemAddResult, TypedHash } from "sp-pnp-js";
// import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/components/Button';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
// import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
// import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
// import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { UrlQueryParameterCollection } from '@microsoft/sp-core-library';
// import { escape } from '@microsoft/sp-lodash-subset';
import { SPComponentLoader } from '@microsoft/sp-loader';
import * as $ from 'jquery';
import './custom.css';
// import { graph } from "@pnp/graph";
// import { dateAdd } from "@pnp/common";


import { SharingResult, SharingRole, SharingLinkKind, ShareLinkResponse, EmailProperties } from "@pnp/sp";

var queryParms = new UrlQueryParameterCollection(window.location.href);
var idd = queryParms.getValue("idd");
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
      receiver: '',


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

      departmentPhone: ''
    }
  }

  formValidation() {
    if ($('#fileUploadInput').val() != ""

      || this.state.receiver === ""
      || this.state.referenceNumberOut === ""
      || this.state.referenceNumberIn
      || this.state.verificationCode === ""
      || $('#date').val() === ""
      || this.state.decryption === ""
    ) {

      return alert('Παρακαλώ συμπληρωστε όλα τα πεδία')
    }

    this.uploadingFileEventHandlers();
    $("#btnForm").hide()
  }

  sharingFolder() {

    const cleanReferenceNumberIn = this.cleanFolderName(this.state.referenceNumberIn)
    const cleanRequest = this.cleanFolderName(this.state.request)

    var day = this.state.requestDate.charAt(0) + this.state.requestDate.charAt(1)
    var month = this.state.requestDate.charAt(3) + this.state.requestDate.charAt(4)
    var year = this.state.requestDate.charAt(6) + this.state.requestDate.charAt(7) + this.state.requestDate.charAt(8) + this.state.requestDate.charAt(9)

    const date = this.state.requestDate
    console.log(date)

    sp
      .web
      .getFolderByServerRelativeUrl("/sites/ExternalSharing/SharedFiles/" + year + month + day + '-' + cleanReferenceNumberIn + '-' + cleanRequest)
      .shareWith(this.state.Email, SharingRole.View, false, false,
        {
          subject: `Σύστημα Διαμοιρασμού Αρχείων ΗΔΙΚΑ ΑΕ : Αίτημα διάθεσης στοιχείων ${year + month + day + '-' + cleanReferenceNumberIn}`,
          body: `Αγαπητέ Παραλήπτη,
Σας ενημερώνουμε ότι μπορείτε να παραλάβετε τα στοιχεία που ζητήσατε από την ΗΔΙΚΑ ΑΕ κατεβάζοντας το σχετικό αρχείο από το κουμπί open.
Πρόκειται για συμπιεσμένο και κρυπτογραφημένο αρχείο. Παρακαλώ επικοινωνήστε στο τηλέφωνο 2132168ΧΧΧ για να λάβετε το απαραίτητο κλειδί αποκρυπτογράφησης.
Οδηγίες για την ανάπτυξη του περιεχομένου του αρχείου θα βρείτε στο www.idika.gr
Παραμένουμε στη διάθεσή σας για περαιτέρω διευκρινήσεις.
Με εκτίμηση,
ΗΔΙΚΑ ΑΕ
`
        })
      .then((result: SharingResult) => {
        console.log(result);
      }).catch(e => {
        console.log(e);
      });
  }

  cancelApplication() {
    $('#btnFormCancel').on('click', () => {
      var confirmation = confirm("Aκύρωση αιτήματος?");
      if (confirmation === true) {
        window.location.replace('https://idikagr.sharepoint.com/sites/ExternalSharing')
      } else { }
    })
  }

  gettindDataFromRequesterApplication() {
    pnp.sp.web.lists.getByTitle("Requests").items.getById(parseInt(idd)).get().then((item: any) => {
      let dateobj = new Date(item.RequestDate);

      this.setState({
        request: ((item.Request == null) ? "" : item.Request),
        Fullname: ((item.Fullname == null) ? "" : item.Fullname),
        Organization: ((item.Organization == null) ? "" : item.Organization),
        PhoneNumber: ((item.PhoneNumber == null) ? "" : item.PhoneNumber),
        Email: ((item.Email == null) ? "" : item.Email),
        Reason: ((item.Reason == null) ? "" : item.Reason),
        requestDate: ((item.RequestDate == null) ? "" : dateobj.toLocaleDateString('en-GB')),
        referenceNumberIn: ((item.ReferenceNumberIn == null) ? "" : item.ReferenceNumberIn),
        receiver: ((item.Fullname == null) ? "" : item.Fullname),
        requesterEmail: ((item.RequesterEmail == null) ? "" : item.RequesterEmail),
        referenceNumberOut: ((item.ReferenceNumberOut == null) ? "" : item.ReferenceNumberOut),
        verificationCode: ((item.VerificationCode == null) ? "" : item.VerificationCode),
        decryption: ((item.Decryption == null) ? "" : item.Decryption),
      });
    });
  }

  componentDidMount() {
    $('#fileUpload').hide()
    this.cancelApplication()
    this.gettindDataFromRequesterApplication()

  }

  public render(): React.ReactElement<IAdminFormProps> {

    pnp.setup({
      spfxContext: this.props.context
    });

    return (
      <form id="form" onSubmit={(e) => { return false }}>

        <div className={"card text-center bg-info mb-3"}>
          <div className={"card-header"}> <h3 className={"text-white"} id={"title"}> Φόρμα Εισαγωγής Αρχείου</h3> </div>
        </div>

        <div className="card card bg-light mb-3">

          <div className="card-header" >
            <h5> Στοιχεία Αιτήματος </h5>
          </div>
          <br></br>

          <div className="form-row" >
            <div className="form-group col-md-6">
              <label><h6> Ονομασία Αιτήματος </h6></label>
              <TextField
                className="form-control"
                readOnly value={this.state.request}
                onChanged={this.handleRequest}
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
                onBlur={() => console.log('onBlur called')} />
            </div>
          </div>

          <div className="form-row" >
            <div className="form-group col-md-6">
              <label> <h6 >Ονοματεπώνυμο</h6></label>
              <TextField
                className="form-control"
                readOnly value={this.state.Fullname}
                onChanged={this.handleFullname}
                errorMessage={(this.state.Fullname.length === 0 && this.state.onSubmission === true) ? this.state.required : ""} placeholder="Όνομα " />
            </div>
            <div className="form-group col-md-6">
              <label> <h6>Οργανισμός/Υπηρεσία</h6></label>
              <TextField
                className="form-control"
                readOnly value={this.state.Organization}
                onChanged={this.handleOrganization}
                errorMessage={(this.state.Organization.length === 0 && this.state.onSubmission === true) ? this.state.required : ""} placeholder="Οργανισμός" />
            </div>
          </div>

          <div className="form-row" >
            <div className="form-group col-md-6">
              <label> <h6 >Τηλέφωνο</h6></label>
              <TextField className="form-control"
                value={this.state.PhoneNumber}
                onChanged={this.handlePhoneNumber}
                errorMessage={(this.state.PhoneNumber.length === 0 && this.state.onSubmission === true) ? this.state.required : ""} placeholder="Τηλέφωνο" />
            </div>
            <div className="form-group col-md-6">
              <label> <h6> Email</h6></label>
              <TextField className="form-control"
                readOnly value={this.state.Email}
                onChanged={this.handleEmail}
                errorMessage={(this.state.Email.length === 0 && this.state.onSubmission === true) ? this.state.required : ""} placeholder="Email" />
            </div>
          </div>
          <br></br>

          <div className="form-group"  >
            <label><h6> Αιτιολογία </h6></label>
          </div>
          <TextField
            className="form-control"
            readOnly multiline={true}
            value={this.state.Reason}
            onChanged={this.handleReason}
            errorMessage={(this.state.Reason.length === 0 && this.state.onSubmission === true) ? this.state.required : ""} placeholder=" Αιτιολογία" />
          <br></br>
          <br></br>
        </div>

        <div className="card bg-light mb-3">
          <div className="card-header" >
            <h5> Φόρμα Διαχειριστή</h5>
          </div>
          <br></br>

          <div className="form-row" >
            <div className="form-group col-md-6">
              <label> <h6 >Παραλήπτης</h6></label>
              <TextField
                className="form-control"
                readOnly value={this.state.Fullname}
                onChanged={this.handleFullname}
                errorMessage={(this.state.Fullname.length === 0 && this.state.onSubmission === true) ? this.state.required : ""} placeholder="Παραλήπτης" />
            </div>

            {/* <div className="form-group col-md-6">
              <label> <h6> Παραλήπτης * </h6></label>
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
                  resolveDelay={1000}
                />
              </div>
              
            </div> */}
            <div className="form-group col-md-6">
              <label> <h6> Ημερομηνία Αίτησης</h6></label>
              <TextField className="form-control" readOnly value={this.state.requestDate} onChanged={this.handleRequestDate}
                errorMessage={(this.state.requestDate.length === 0 && this.state.onSubmission === true) ? this.state.required : ""} placeholder="Ημερομηνία Αίτησης" />
            </div>
          </div>
          <br></br>

          <div className="form-row" >
            <div className="form-group col-md-6">
              <label> <h6 >Αρ. Πρωτ. Εισερχομένου * </h6></label>
              <TextField id="inputref" className="form-control" value={this.state.referenceNumberIn} onChanged={this.handleReferenceNumberIn}
                errorMessage={(this.state.referenceNumberIn.length === 0 && this.state.onSubmission === true) ? this.state.required : ""} placeholder="Αρ. Πρωτ. Εισερχομένου " />
            </div>
            <div className="form-group col-md-6">
              <label> <h6> Κωδικός Επιβεβαίωσης Aκεραιότητας * </h6></label>
              <TextField className="form-control" value={this.state.verificationCode} onChanged={this.handleVerificationCode}
                errorMessage={(this.state.verificationCode.length === 0 && this.state.onSubmission === true) ? this.state.required : ""} placeholder="Κωδικός Επιβεβαίωσης" />
            </div>
          </div>
          <div className="form-row" >
            <div className="form-group col-md-6">
              <label> <h6> Αρ. Πρωτ. Εξερχομένου * </h6></label>
              <TextField
                className="form-control"
                value={this.state.referenceNumberOut}
                onChanged={this.handleReferenceNumberOut}
                errorMessage={(this.state.referenceNumberOut.length === 0 && this.state.onSubmission === true) ? this.state.required : ""}
                placeholder="Αρ. Πρωτ. Εξερχομένου " />
            </div>
            <div className="form-group col-md-6">
              <label> <h6> Ημερομηνία Εξερχομένου * </h6></label>
              <input
                id="date"
                type="date"
                className="form-control"
                name="date"
                style={{ height: '2.9em' }}
                onChange={this.handleDate}
              />
            </div>
          </div>

          <br></br>
          <div className="form-row" >
            <div className="form-group col-md-6 ">
              <label> <h6> Κωδικός Αποκρυπτογράφησης * </h6></label>
              <div className="input-group ">
                <PrimaryButton className="btn btn-dark text-white btn-sm" id="buttonGenPass" text="Generate" onClick={() => {
                  this.generatePassword();
                }} />

                <TextField id="inputGenPass" className="form-control border-0" value={this.state.decryption} onChanged={this.handleDecryption}
                  errorMessage={(this.state.decryption.length === 0 && this.state.onSubmission === true) ? this.state.required : ""} placeholder="Κωδικός Αποκρυπτογράφησης" />

              </div>
            </div>

            <div className="form-group col-md-6">
              <label> <h6> Ανάρτηση Αρχείου * </h6></label>
              <input required={true} className="form-control" type='file' id='fileUploadInput' name='myfile' />
              <button id="fileUpload" name="uFile">upload</button>
            </div>
          </div>
        </div>

        <br></br>
        <PrimaryButton id="shareButton" onClick={() => { this.sharingFolder() }}> Share File </PrimaryButton>

        <PrimaryButton id="btnForm" className="btn btn-dark btn-lg btn-block" onClick={() => { this.formValidation() }} style={{ marginRight: '8px' }}><h5> Υποβολή Αιτήματος </h5> </PrimaryButton>
        <DefaultButton id="btnFormCancel" className="btn btn-light btn-lg btn-block border" onClick={() => { this.setState({}); }}  > <h5> Ακύρωση Αιτήματος </h5> </DefaultButton>
      </form>
    );
  }

  private uploadingFileEventHandlers() {
    let fileUpload = document.getElementById("fileUpload")
    let test1 = document.getElementById("fileUploadInput")

    if (fileUpload) {
      this.uploadFiles(test1);
    }
  }

  protected async uploadFiles(fileUpload) {
    var day = this.state.requestDate.charAt(0) + this.state.requestDate.charAt(1)
    var month = this.state.requestDate.charAt(3) + this.state.requestDate.charAt(4)
    var year = this.state.requestDate.charAt(6) + this.state.requestDate.charAt(7) + this.state.requestDate.charAt(8) + this.state.requestDate.charAt(9)

    const cleanReferenceNumberIn = this.cleanFolderName(this.state.referenceNumberIn)
    const cleanRequest = this.cleanFolderName(this.state.request)

    let file = fileUpload.files[0];

    if (file.size <= 10485760) {
      // small upload
      await web
        .folders
        .add('/sites/ExternalSharing/SharedFiles/' + year + month + day + '-' + cleanReferenceNumberIn + '-' + cleanRequest)
        .then(console.log);

      await web.getFolderByServerRelativeUrl("/sites/ExternalSharing/SharedFiles/" + year + month + day + '-' + cleanReferenceNumberIn + '-' + cleanRequest)
        .files.add(file.name, file, true)
        .then(result => console.log(result)).catch(e => {
          alert("Problima")
          return false
        })

      await sp
        .web
        .getFolderByServerRelativeUrl("/sites/ExternalSharing/SharedFiles/" + year + month + day + '-' + cleanReferenceNumberIn + '-' + cleanRequest)
        .shareWith(this.state.Email, SharingRole.None, false, false, { subject: 'ΗΔΙΚΑ', body: 'Αγαπητέ παραλήπτη, <br> Σας ενημερώνουμε ότι μπορείτε να παραλάβετε τα στοιχεία που ζητήσατε από την ΗΔΙΚΑ ΑΕ κατεβάζοντας το σχετικό αρχείο από εδώ(link)Πρόκειται για συμπιεσμένο και κρυπτογραφημένο αρχείο. Παρακαλώ επικοινωνήστε με το τμήμα ΧΧΧ στο τηλέφωνο 2132168ΧΧΧ για να λάβετε το απαραίτητο κλειδί αποκρυπτογράφησης Οδηγίες για την ανάπτυξη του περιεχομένου του αρχείου θα βρείτε εδώ(link με οδηγίες και συνδέσμους λήψης εργαλείου αποσυμπίεσης και παραγωγής integrity hash key που θα ανέβουν στο www.idika.gr <br> Παραμένουμε στη διάθεσή σας για περαιτέρω διευκρινήσεις <br> Με εκτίμηση, <br> ΗΔΙΚΑ ΑΕ' })
        .then((result: SharingResult) => {
          console.log(result);
        }).catch(e => {
          console.error(e);
        });

    } else { // large upload
      await web
        .folders
        .add('/sites/ExternalSharing/SharedFiles/' + year + month + day + '-' + cleanReferenceNumberIn + '-' + cleanRequest)
        .then(console.log);
      await web.getFolderByServerRelativeUrl("/sites/ExternalSharing/SharedFiles/" + year + month + day + '-' + cleanReferenceNumberIn + '-' + cleanRequest)
        .files
        .addChunked(file.name, file, data => {
        }, true)
        .then(_ => console.log("done!"));

      await sp
        .web
        .getFolderByServerRelativeUrl("/sites/ExternalSharing/SharedFiles/" + year + month + day + '-' + cleanReferenceNumberIn + '-' + cleanRequest)
        .shareWith(this.state.Email, SharingRole.None, false, false, { subject: 'test', body: 'test' })
        .then((result: SharingResult) => {
          console.log(result);
        }).catch(e => {
          console.error(e);
        });

    }

    this.updateItem(year, month, day, cleanReferenceNumberIn, cleanRequest, file.name)

    await pnp.sp.web.lists.getByTitle("Files").items.add({
      FileName: file.name,
      Path: "https://idikagr.sharepoint.com/sites/ExternalSharing/_layouts/download.aspx?sourceurl=/sites/ExternalSharing/SharedFiles/" + year + month + day + '-' + cleanReferenceNumberIn + '-' + cleanRequest + "/" + file.name,
      RequestId: parseInt(idd),
      ReferenceNumberIn: this.state.referenceNumberIn,
      ReferenceNumberOut: this.state.referenceNumberOut,
      ReferenceNumberOutDate: this.state.date.toString(),
      VerificationCode: this.state.verificationCode,
      Decryption: this.state.decryption,
      // refDateOut: this.state.date.toString()
    }).then((iar: ItemAddResult) => {
      // $("#form").hide()
      // $("#btnForm").hide()
      // $("#btnFormCancel").hide()
      console.log(iar);
    }).then(() => {
      alert("Το αίτημα καταχωρήθηκε επιτυχώς")
      // window.location.replace('https://idikagr.sharepoint.com/sites/ExternalSharing/SitePages/Home.aspx')
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

  protected cleanFolderName(referenceNumberIn: String) {
    const validFolderName = referenceNumberIn.replace(/\s+/gi, '-'); // Replace white space with dash
    return validFolderName.replace(/[^a-zA-Z0-9\^α-ωΑ-Ω-άέ-ήί-ό\-]/gi, ''); // Strip any special characterer
  }

  protected createFile() {
    //create folder
    var dat = new Date(this.state.requestDate);

    var day = this.state.requestDate.charAt(0) + this.state.requestDate.charAt(1)
    var month = this.state.requestDate.charAt(3) + this.state.requestDate.charAt(4)
    var year = this.state.requestDate.charAt(6) + this.state.requestDate.charAt(7) + this.state.requestDate.charAt(8) + this.state.requestDate.charAt(9)

    const cleanReferenceNumberIn = this.cleanFolderName(this.state.referenceNumberIn)
    const cleanRequest = this.cleanFolderName(this.state.request)

    web
      .folders
      .add('/sites/ExternalSharing/SharedFiles/' + year + month + day + '-' + cleanReferenceNumberIn + '-' + cleanRequest)
      .then(console.log);
  }

  private _onClosePanel = () => {
    this.setState({ showPanel: false });
  }

  private _changeSharing(checked: any): void {
    this.setState({ defaultChecked: checked });
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


  /**
 * This is a function.
 *
 * @param {string} year - A string param
 * @param {string} month - A string param
 * @param {string} day - A string param
 * @return {void} 
 * @example
 *
 *     updateItem()
 */
  private updateItem(_year, _month, _day, _cleanReferenceNumberIn, _cleanReaquest, _file): void {
    var checkboxValue = this.state.defaultChecked ? "Yes" : "No";
    console.log(this.state.defaultChecked);

    pnp.sp.web.lists.getByTitle("Requests").items.getById(parseInt(idd)).update({
      Path: "https://idikagr.sharepoint.com/sites/ExternalSharing/_layouts/download.aspx?sourceurl=/sites/ExternalSharing/SharedFiles/" + _year + _month + _day + '-' + _cleanReferenceNumberIn + '-' + _cleanReaquest + "/" + _file,
      ReceiverId: this.state.userManagerIDs[0],
      ReferenceNumberIn: this.state.referenceNumberIn,
      Completed: checkboxValue
    }).then((iar: ItemUpdateResult) => {
      console.log(iar);
      this.setState({ status: "Your request has been submitted sucessfully." });

    });
  }
}


