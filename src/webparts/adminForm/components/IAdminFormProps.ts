import { WebPartContext } from '@microsoft/sp-webpart-base';


export interface IAdminFormProps {
  description: string;
  context: WebPartContext;
  siteUrl: string;
  fn: string;
  rd: string;
  re: string;
  rni: string;
  rno: string;
  rnod: string;
  vc: string;
  de: string;
  digest: string;
  listName: string;
  request: string;
  queryString: string;
  uploadFilesTo: string;
}
