import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';

export interface IReactSpFxPnP {
  selectedItems: any[];
  request: string;
  requestDate: string;
  requesterEmail: string;
  referenceNumberIn: string;
  referenceNumberOut: string;
  referenceNumberOutDate: Date;
  verificationCode: string;
  decryption: string;
  description: string;
  dpselectedItem?: { key: string | number | undefined };
  termKey?: string | number;
  dpselectedItems: IDropdownOption[];
  disableToggle: boolean;
  defaultChecked: boolean;
  pplPickerType: string;
  userManagerIDs: number[];
  hideDialog: boolean;
  status: string;
  isChecked: boolean;
  showPanel: boolean;
  required: string;
  onSubmission: boolean;
  termnCond: boolean;
  Fullname: string;
  Organization: string;
  PhoneNumber: string;
  Email: string;
  Reason: string;
  date: '';
  receiver: '';
  departmentPhone: '';
  department: '';
  departmentEmail: '';
  lastShare: '';
}