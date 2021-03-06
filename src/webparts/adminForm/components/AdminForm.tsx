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
import * as moment from 'moment';

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
    this.handleDepartment = this.handleDepartment.bind(this);

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
      department: '',
      departmentPhone: '',
      departmentEmail: '',
      lastShare: '',


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
      termnCond: false
    }
  }

  async departmentOption() {

    await sp.web.lists.getByTitle("Requests").items.getById(parseInt(idd)).get().then((data) => {
      const department = data.Department
      $("#selectteam").empty();
      $('#selectteam').append('<option select>' + department + '</option>')

      sp.web.lists.getByTitle("Department").items.get().then((data) => {

        data.map((item) => {
          if (department === item.NameDepartment) {
            return false
          }
          $('#selectteam').append('<option>' + item.NameDepartment + '</option>')
        })
      })
    })
  }

  protected hidingForm() {

    $(".form-group").hide()
    $("#btnForm").hide()
    $("#section").hide()
    $("#section2").hide()
    $("#TextField45").hide()
    // $("#btnFormCancel").hide()
    $("#btnShare").hide()
    $("#btnSave").hide()
    $("#section3").hide()

    $('#newSection').append(`
          <div class="card"> 
            <div class="card-body">
              <h5 class="card-title text-center">Επιτυχής Υποβολή</h5>
            </div>
          </div> 
              <br>`);
  }

  protected async formValidation() {



    if (($('#fileUploadInput').val() == "")
      || this.state.receiver == ""
      || this.state.referenceNumberOut == ""
      || this.state.verificationCode == ""
      || this.state.Fullname == ""
      || this.state.referenceNumberIn == ""
      || this.state.referenceNumberOutDate == null
      || ($('#date').val() === "")
      || this.state.decryption == "") {
      return alert('Παρακαλώ συμπληρωστε όλα τα πεδία')
    }

    await this.uploadingFileEventHandlers().then((result) => {
      this.hidingForm()
    }).catch(err => console.log(err))

  }

  protected async sharingFolder() {

    let sucessfully = false

    return new Promise(async (resolve, reject) => {
      const cleanReferenceNumberIn = this.cleanFolderName(this.state.referenceNumberIn)
      const cleanRequest = this.cleanFolderName(this.state.request)

      var day = this.state.requestDate.charAt(0) + this.state.requestDate.charAt(1)
      var month = this.state.requestDate.charAt(3) + this.state.requestDate.charAt(4)
      var year = this.state.requestDate.charAt(6) + this.state.requestDate.charAt(7) + this.state.requestDate.charAt(8) + this.state.requestDate.charAt(9)

      const date = this.state.requestDate

      await pnp.sp.web.lists.getByTitle("Requests").items.getById(parseInt(idd)).get().then((item: any) => {

        this.setState({
          departmentPhone: item.DepartmentPhone,
          departmentEmail: item.eMailDepartment
        });
      });


      const emailBody = `Αγαπητέ Παραλήπτη,
Σας ενημερώνουμε ότι μπορείτε να παραλάβετε τα στοιχεία που ζητήσατε από την ΗΔΙΚΑ ΑΕ κατεβάζοντας το σχετικό αρχείο από το κουμπί open.
Πρόκειται για συμπιεσμένο και κρυπτογραφημένο αρχείο. Παρακαλώ επικοινωνήστε με το τμήμα ${this.state.department} στο τηλέφωνο ${this.state.departmentPhone} για να λάβετε το απαραίτητο κλειδί αποκρυπτογράφησης.
Οδηγίες για την ανάπτυξη του περιεχομένου του αρχείου θα βρείτε στο www.idika.gr
Παραμένουμε στη διάθεσή.
Με εκτίμηση,
ΗΔΙΚΑ ΑΕ`

      const emailSubject = `Σύστημα Διαμοιρασμού Αρχείων ΗΔΙΚΑ ΑΕ : Αίτημα διάθεσης στοιχείων ${year + month + day + '-' + cleanReferenceNumberIn}`

      await sp
        .web
        .getFolderByServerRelativeUrl("/sites/ExternalSharing/SharedFiles/" + year + month + day + '-' + cleanReferenceNumberIn + '-' + cleanRequest)
        .shareWith(this.state.Email, SharingRole.View, false, false,
          {
            subject: emailSubject,
            body: emailBody
          })
        .then((result: SharingResult) => {
          console.log(result);
          // const emailProps: EmailProperties = {
          //   To: [this.state.departmentEmail],
          //   Subject: emailSubject,
          //   Body: emailBody
          // };
          sucessfully = true
          pnp.sp.web.lists.getByTitle("Requests").items.getById(parseInt(idd)).update({
            LastShare: moment().format('dddd, DD/MM/YYYY, h:mm:ss a')
          }).then((iar: ItemUpdateResult) => {
            console.log(iar);
            this.setState({ status: "Your request has been submitted sucessfully." });

          });


          // sp.utility.sendEmail(emailProps).then(result => {
          //   console.log(result);
          // }).catch(e => { console.log(e) })
        }).catch(e => {
          alert('Ο διαμοιρασμός αρχείων δεν μπορεί να πραγματοποιηθεί καθώς ο φάκελος δεν υπάρχει ή έχει διαγραφεί. Παρακαλώ συμπληρώστε τα πεδία και πατήστε υποβολή και προσπαθήστε ξανά.')
          // console.log(e);
        });

      if (sucessfully) {
        resolve(true)
      } else {
        reject(false)
      }


    })
  }

  cancelApplication() {
    $('#btnFormCancel').on('click', () => {
      window.location.replace('https://idikagr.sharepoint.com/sites/ExternalSharing')
    })
  }

  gettindDataFromRequesterApplication() {
    pnp.sp.web.lists.getByTitle("Requests").items.getById(parseInt(idd)).get().then((item: any) => {
      let dateobj = new Date(item.RequestDate);

      this.setState({
        department: item.Department,
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
        departmentPhone: ((item.DepartmentPhone == null) ? "" : item.DepartmentPhone),
        departmentEmail: ((item.eMailDepartment == null) ? "" : item.eMailDepartment)

      });
    });
  }

  saveChanges() {

    if (this.state.Fullname == "" || this.state.referenceNumberIn == "") {
      return alert('Τα πεδία Ονοματεπώνυμο Αιτούντος και Αρ. πρωτ. Εισερχομένου ΗΔΙΚΑ θα πρέπει να είναι συμπληρωμένα')
    }

    sp.web.lists.getByTitle("Department").items.get().then((item: any) => {
      item.map((data) => {

        if (this.state.department != data.NameDepartment) {
          return false
        }

        pnp.sp.web.lists.getByTitle("Requests").items.getById(parseInt(idd)).update({
          Department: this.state.department,
          Fullname: this.state.Fullname,
          DepartmentPhone: data.PhoneDepartment,
          eMailDepartment: data.email
        }).then((result) => {
          console.log(result);
        });
      })
    })

    alert('Οι αλλαγές αποθηκεύτηκαν επιτυχώς')

  }

  async componentDidMount() {
    await this.gettindDataFromRequesterApplication()
    $('#fileUpload').hide()
    await this.cancelApplication()
    await this.departmentOption()

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
        <div id="newSection"></div>
        <div id="section" className="card card bg-light mb-3">
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
              <label> <h6> Αίτημα </h6></label>
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
              <label> <h6 >Ονοματεπώνυμο Αιτούντος</h6></label>
              <TextField
                className="form-control"
                readOnly value={this.state.Fullname}
                onChanged={this.handleFullname}
                errorMessage={(this.state.Fullname.length === 0 && this.state.onSubmission === true) ? this.state.required : ""} placeholder="Όνομα " />
            </div>
            <div className="form-group col-md-6">
              <label> <h6>Οργανισμός/Υπηρεσία Αιτούντος</h6></label>
              <TextField
                className="form-control"
                readOnly value={this.state.Organization}
                onChanged={this.handleOrganization}
                errorMessage={(this.state.Organization.length === 0 && this.state.onSubmission === true) ? this.state.required : ""} placeholder="Οργανισμός" />
            </div>
          </div>

          <div className="form-row" >
            <div className="form-group col-md-6">
              <label> <h6 >Τηλέφωνο Επικοινωνίας Αιτούντος</h6></label>
              <TextField className="form-control"
                value={this.state.PhoneNumber}
                onChanged={this.handlePhoneNumber}
                errorMessage={(this.state.PhoneNumber.length === 0 && this.state.onSubmission === true) ? this.state.required : ""} placeholder="Τηλέφωνο" />
            </div>
            <div className="form-group col-md-6">
              <label> <h6> eMail Αιτούντος</h6></label>
              <TextField className="form-control"
                readOnly value={this.state.Email}
                onChanged={this.handleEmail}
                errorMessage={(this.state.Email.length === 0 && this.state.onSubmission === true) ? this.state.required : ""} placeholder="Email" />
            </div>
          </div>
          <br></br>

          <div className="form-group"  >
            <label><h6> Αιτιολόγηση Αιτήματος </h6></label>
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

        <div id="section2" className="card bg-light mb-3">
          <div className="card-header" >
            <h5> Φόρμα Διαχειριστή</h5>
          </div>
          <br></br>

          <div id="newSection"></div>

          <div className="form-row" >
            <div className="form-group col-md-4">
              <label> <h6 >Ονοματεπώνυμο Αιτούντος</h6></label>
              <TextField
                className="form-control"
                // readOnly 
                value={this.state.Fullname}
                onChanged={this.handleFullname}
                errorMessage={(this.state.Fullname.length === 0 && this.state.onSubmission === true) ? this.state.required : ""} />
            </div>

            <div className="form-group col-md-4">
              <label> <h6>  Τμήμα ΗΔΙΚΑ </h6> </label>
              <select style={{ height: '2.88em' }} onChange={this.handleDepartment} value={this.state.department} id="selectteam" className="form-control">
              </select>
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
            <div className="form-group col-md-4">
              <label> <h6> Ημερομηνία Αίτησης</h6></label>
              <TextField className="form-control" readOnly value={this.state.requestDate} onChanged={this.handleRequestDate}
                errorMessage={(this.state.requestDate.length === 0 && this.state.onSubmission === true) ? this.state.required : ""} placeholder="Ημερομηνία Αίτησης" />
            </div>
          </div>
          <br></br>

          <div className="form-row" >
            <div className="form-group col-md-6">
              <label> <h6 >Αρ. πρωτ. Εισερχομένου ΗΔΙΚΑ</h6></label>
              <TextField id="inputref" className="form-control" readOnly value={this.state.referenceNumberIn} onChanged={this.handleReferenceNumberIn}
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
              <label> <h6> Αρ. Πρωτ. Εξερχομένου ΗΔΙΚΑ * </h6></label>
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
        <br></br>

        <PrimaryButton id="btnForm" className="btn btn-dark btn-lg btn-block" onClick={() => { this.formValidation() }} style={{ marginRight: '8px' }}><h5> Υποβολή Αιτήματος </h5> </PrimaryButton>
        <DefaultButton id="btnFormCancel" className="btn btn-light btn-lg btn-block border" onClick={() => { window.location.replace('https://idikagr.sharepoint.com/sites/ExternalSharing') }}  > <h5> Μετάβαση στη σελίδα διαχείρισης </h5> </DefaultButton>

        <br></br>
        <br></br>
        <div id="section3" className="card bg-light mb-3">
          <div className="card-header" >
            <h5> Περισσότερες ενέργειες  </h5>
          </div>

          <div className="form-row" >
            <div className="form-group col-md-6">
              <div className="input-group ">
                <label > <p style={{ margin: '1em', marginBottom: '2.5em' }}> Εφόσον έχετε υποβάλει το αίτημα και θέλετε να επαναλάβετε την διαδικασία διαμοιρασμού φακέλου με τον εξωτερικό χρήστη. </p></label>
                <PrimaryButton style={{ width: '13em' }} text="Επανάληψη Διαμοιρασμού Φακέλου" className="btn btn-secondary text-white btn-sm col text-center" id="btnShare" onClick={() => {
                  this.sharingFolder().then((result) => {
                    alert("H Επανάληψη Διαμοιρασμού Φακέλου ολοκληρώθηκε επιτυχώς")
                  })

                }}> </PrimaryButton>
              </div>
            </div>

            <div className="form-group col-md-6">
              <div className="input-group ">
                <label> <p style={{ margin: '1em' }}> Σε περίπτωση που έχετε κάνει αλλαγές στη φόρμα διαχειριστή (τμήμα ΗΔΙΚΑ - ονοματεπώνυμο αιτούντος)) και θέλετε να τις αποθηκεύσετε. Χωρίς να γίνεται υποβολή αιτήματος. </p></label>
                <PrimaryButton style={{ width: '13em' }} text="Αποθήκευση Αλλαγών" className="btn btn-secondary text-white btn-sm col text-center" id="btnSave" onClick={() => { this.saveChanges() }}> </PrimaryButton>
              </div>
            </div>
          </div>
        </div>
      </form>
    );
  }

  private uploadingFileEventHandlers() {

    let flag = false

    return new Promise((resolve, reject) => {
      let fileUpload = document.getElementById("fileUpload")
      let test1 = document.getElementById("fileUploadInput")

      if (fileUpload) {
        this.uploadFiles(test1);
        flag = true
      }

      if (flag) {
        return resolve(true)
      }
      reject('Error:file not uploaded')
    })
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
      await this.updateItem(year, month, day, cleanReferenceNumberIn, cleanRequest, file.name)

      await this.sharingFolder()

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

      await this.updateItem(year, month, day, cleanReferenceNumberIn, cleanRequest, file.name)

      await this.sharingFolder()

    }

    await pnp.sp.web.lists.getByTitle("Files").items.add({
      FileName: file.name,
      Path: "https://idikagr.sharepoint.com/sites/ExternalSharing/_layouts/download.aspx?sourceurl=/sites/ExternalSharing/SharedFiles/" + year + month + day + '-' + cleanReferenceNumberIn + '-' + cleanRequest + "/" + file.name,
      RequestId: parseInt(idd),
      ReferenceNumberIn: this.state.referenceNumberIn,
      ReferenceNumberOut: this.state.referenceNumberOut,
      ReferenceNumberOutDate: this.state.date.toString(),
      VerificationCode: this.state.verificationCode,
      Decryption: this.state.decryption,
      FolderReceiver: this.state.receiver,
      RequestDate: this.state.requestDate,
      Department: this.state.department,
      ShareDate: moment().format('dddd, DD/MM/YYYY, h:mm:ss a')
    }).then((result) => {
      console.log(result);
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

  private handleDepartment(e) {
    return this.setState({
      department: e.target.value,
    })
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

  private updateItem(_year, _month, _day, _cleanReferenceNumberIn, _cleanReaquest, _file): void {
    var checkboxValue = this.state.defaultChecked ? "Yes" : "No";
    console.log(this.state.defaultChecked);

    sp.web.lists.getByTitle("Department").items.get().then((item: any) => {
      item.map((data) => {
        if (this.state.department != data.NameDepartment) {
          return false
        }

        pnp.sp.web.lists.getByTitle("Requests").items.getById(parseInt(idd)).update({
          Path: "https://idikagr.sharepoint.com/sites/ExternalSharing/_layouts/download.aspx?sourceurl=/sites/ExternalSharing/SharedFiles/" + _year + _month + _day + '-' + _cleanReferenceNumberIn + '-' + _cleanReaquest + "/" + _file,
          ReceiverId: this.state.userManagerIDs[0],
          Fullname: this.state.Fullname,
          ReferenceNumberIn: this.state.referenceNumberIn,
          Completed: checkboxValue,
          DepartmentPhone: data.PhoneDepartment,
          eMailDepartment: data.email,
          Department: this.state.department,
        }).then((iar: ItemUpdateResult) => {
          console.log(iar);
          this.setState({ status: "Your request has been submitted sucessfully." });
        });
      })
    })
  }
}


