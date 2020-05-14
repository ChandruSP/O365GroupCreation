export interface ICreateO365GroupsState {
  hasMarketingMember: boolean;
  commonFolder: boolean;
  formData: {
    members: string;
    countryCode: string;
    companyCode: string;
    groupCode: string;
    projectNumber: string;
    taskNumber: string;
    shortDescription: string;
    description: string;
    visibility: string;
    sendmail: string;
  };
}
