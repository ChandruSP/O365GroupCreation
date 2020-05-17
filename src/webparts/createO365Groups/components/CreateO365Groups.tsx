import * as React from 'react';
import styles from './CreateO365Groups.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { SPComponentLoader } from "@microsoft/sp-loader";
import { HttpClient, HttpClientResponse, IHttpClientOptions } from '@microsoft/sp-http';
import { ICreateO365GroupsProps } from './ICreateO365GroupsProps';
import { ICreateO365GroupsState } from './ICreateO365GroupsState';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";



import { sp } from "@pnp/sp";
import { Web } from "@pnp/sp/webs";
import { Lists, ILists } from "@pnp/sp/lists";

import "@pnp/sp/webs";
import "@pnp/sp/lists/web";
import "@pnp/sp/items/list";
import "@pnp/sp/folders";
import "@pnp/sp/folders/list";



import { Dropdown, DropdownMenuItemType, IDropdownOption, IDropdownStyles } from 'office-ui-fabric-react/lib/Dropdown';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { TextField, MaskedTextField } from 'office-ui-fabric-react/lib/TextField';
import { Stack, IStackProps, IStackStyles } from 'office-ui-fabric-react/lib/Stack';
import { DefaultButton, PrimaryButton, IStackTokens } from 'office-ui-fabric-react';
import { ChoiceGroup, IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';
import 'alertifyjs';
import '../../../ExternalRef/CSS/style.css'
import '../../../ExternalRef/CSS/alertify.min.css';  
var alertify: any = require('../../../ExternalRef/JS/alertify.min.js');
export default class CreateO365Groups extends React.Component<ICreateO365GroupsProps, ICreateO365GroupsState> {
  globalDisplayName: string = '';
  createPlannerNow: boolean = false;
  members: string[] = [];
  admins = {
    practiceLeader: null,
    groupLeader: null,
    projAdmin: null,
    marketingMember: null,
  };
  mailNickName = '';

  flowpostURL: string = 'https://prod-09.centralindia.logic.azure.com:443/workflows/54949b246a8645069ba77a35619daa08/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=3UodVAI2oglFIB9tZq3-HubBvnrBmd_yxfLkPVWKFxU';


  commonFolders: string[] = ['CORRESPONDENCE', 'MARKETING', 'POSTDESIGN', 'PROJADMIN', 'CORRESPONDENCE/FROM CLIENT', 'CORRESPONDENCE/FROM SUBS', 'CORRESPONDENCE/TRANSMITTALS', 'CORRESPONDENCE/TRANSMITTALS/30\% DESIGN', 'CORRESPONDENCE/TRANSMITTALS/50\% DESIGN', 'CORRESPONDENCE/TRANSMITTALS/90\% DESIGN', 'CORRESPONDENCE/TRANSMITTALS/100\% DESIGN', 'CORRESPONDENCE/TRANSMITTALS/AS-BUILT', 'MARKETING/PROPOSAL', 'MARKETING/RESUMES', 'PROJADMIN/CONTRACT', 'PROJADMIN/SUBCONTRACTS', 'PROJADMIN/INVOICES'];

  constructor(props: ICreateO365GroupsProps) {
    super(props);
    this.state = {
      hasMarketingMember: false,
      commonFolder: true,
      formData: {
        countryCode: '',
        companyCode: '',
        groupCode: '',
        projectNumber: '',
        taskNumber: '',
        shortDescription: '',
        members: '',
        description: '',
        visibility: 'Public',
        sendmail: '',
      }
    };
  }


  createPlanner(groupId, title) {
    var plannerPlan = {
      owner: groupId,
      title: title
    };
    this.props.graphClient
      .api('/planner/plans')
      .post(plannerPlan)
      .then((content: any) => {
        alertify.set("notifier", "position", "top-right");
        alertify.success('Planner created successfully');
      })
      
      .catch(err => {
        alertify.set("notifier", "position", "top-right");
        alertify.error('Error while creating planner');
      });
  }

  folderCreation(index, connectWeb) {
    console.log('Folder create git: ' + this.commonFolders[index]);
    var that = this;
    connectWeb.folders.add('Shared%20Documents/' + this.commonFolders[index])
      .then(function (data) {
        index = index + 1;
        if (index < that.commonFolders.length) {
          that.folderCreation(index, connectWeb);
        } else {
          alertify.set("notifier", "position", "top-right");
          alertify.success('Folder created successfully');
        }
      })
      .catch(function (err) {
        console.log(err);
        alertify.set("notifier", "position", "top-right");
        alertify.error('Error while creationg folders');
      });
  }

  async createFolder(siteName) {
    var absoluteUrl = this.props.context.pageContext.site.absoluteUrl.split('/');
    var domain = absoluteUrl[0] + '//' + absoluteUrl[2];
    siteName = domain + '/sites/' + siteName;
    var connectWeb = Web(siteName);
    this.folderCreation(0, connectWeb);
  }

  addmembertogroup(userid, email, groupId, role) {
    var that = this;
    var odata = "https://graph.microsoft.com/v1.0/users/" + userid;
    if (role == 'members') {
      odata = "https://graph.microsoft.com/v1.0/directoryObjects/" + userid;
    }

    var user = {
      "@odata.id": odata
    };
    this.props.graphClient
      .api('/groups/' + groupId + '/' + role + '/$ref')
      .post(user)
      .then((content: any) => {
        alertify.set("notifier", "position", "top-right");
        alertify.success('User ' + email + ' added');

        if (that.createPlannerNow) {
          that.createPlannerNow = false;
          setTimeout(function () {
            that.createPlanner(groupId, that.globalDisplayName);
            var postData = {
              "groupId": groupId
            };
              
            const requestHeaders: Headers = new Headers();
            requestHeaders.append('Content-type', 'application/json');
            requestHeaders.append('Cache-Control', 'no-cache');
            const httpClientOptions: IHttpClientOptions = {
              body: JSON.stringify(postData),
              headers: requestHeaders
            };

            that.props.httpClient.post(
              that.flowpostURL,
              HttpClient.configurations.v1,
              httpClientOptions)
              .then(function (res) {
                alertify.set("notifier", "position", "top-right");
                alertify.success('Teams created successfully');
                if (that.state.commonFolder == true) {
                  that.createFolder(that.mailNickName);
                }
              });

          }, 5000)
        }

      })
      .catch(err => {
        alertify.set("notifier", "position", "top-right");
        alertify.error('Error while creating member');
      });

  }

  getuser(email, groupId, role) {
    var that = this;
    var cleartext = email.replace(/\s+/g, '');
    this.props.graphClient
      .api('/users/' + cleartext)
      .get()
      .then((content: any) => {
        that.addmembertogroup(content.id, email, groupId, role);
      })
      .catch(err => {

      });
  }

  createADMembers(groupId) {
    if (this.members.length == 0) {
      this.createPlannerNow = true;
    }
    if (this.admins.groupLeader) {
      this.getuser(this.admins.groupLeader, groupId, 'owners');
    }
    if (this.admins.practiceLeader) {
      this.getuser(this.admins.practiceLeader, groupId, 'owners');
    }
    if (this.admins.projAdmin) {
      this.getuser(this.admins.projAdmin, groupId, 'owners');
    }
    if (this.state.hasMarketingMember) {
      if (this.admins.marketingMember) {
        this.getuser(this.admins.marketingMember, groupId, 'owners');
      }
    }

    this.getuser(this.props.userEmail, groupId, 'members');

    for (let index = 0; index < this.members.length; index++) {
      const member = this.members[index];
      if (index == this.members.length - 1) {
        this.createPlannerNow = true;
      }
      if (member) {
        this.getuser(member, groupId, 'members');
      }
    }
  }


  createADGroup() {
    let formData = this.state.formData;
    var mailNickname = formData.countryCode + '-' + formData.companyCode + '-' + formData.groupCode + '-' + formData.projectNumber + '-' + formData.taskNumber;
    var displayName = mailNickname + '-' + formData.shortDescription;
    var clearMailNickname = mailNickname.replace(/[^a-zA-Z0-9]/g, "");
    this.globalDisplayName = displayName;
    this.mailNickName = clearMailNickname;

    var details = {
      "displayName": displayName,
      "groupTypes": [
        "Unified"
      ],
      "mailEnabled": true,
      "mailNickname": clearMailNickname,
      "securityEnabled": false,
      "visibility": formData.visibility
    };

    if (formData.description) {
      details["description"] = formData.description;
    }

    var that = this;
    this.props.graphClient
      .api('/groups')
      .post(details)
      .then((content: any) => {
        that.createADMembers(content.id);
        alertify.set("notifier", "position", "top-right");
        alertify.success('Group created successfully');
      })
      .catch(err => {
        alertify.set("notifier", "position", "top-right");
        alertify.error('Error while creating a group');
      });

  }

  formHandler() {

    let formData = this.state.formData;
    if (!formData.countryCode) {
      alertify.set("notifier", "position", "top-right");
      alertify.error('Country code is required');
      return;
    }
    if (!formData.companyCode) {
      alertify.set("notifier", "position", "top-right");
      alertify.error('Company code is required');
      return;
    }
    if (!formData.groupCode) {
      alertify.set("notifier", "position", "top-right");
      alertify.error('Group code is required');
      return;
    }
    if (!formData.projectNumber) {
      alertify.set("notifier", "position", "top-right");
      alertify.error('Project number is required');
      return;
    }
    if (!formData.taskNumber) {
      alertify.set("notifier", "position", "top-right");
      alertify.error('Task number is required');
      return;
    }
    if (!formData.shortDescription) {
      alertify.set("notifier", "position", "top-right");
      alertify.error('Short description is required');
      return;
    }
    this.createADGroup();
  }

  inputChangeHandler(e) {
    let formData = this.state.formData;
    formData[e.target.name] = e.target.value;
    this.setState({
      formData
    });
  }

  private getUserData(user) {
    var sdata = user.id.split('|');
    return sdata[sdata.length - 1];
  }

  private _getPeoplePickerItems(items: any[]) {
    this.members = [];
    for (let index = 0; index < items.length; index++) {
      var user = items[index];
      this.members.push(this.getUserData(user));
    }
  }

  private getPracticeLeader(items: any[]) {
    if (items.length == 0) {
      this.admins.practiceLeader = null;
      return;
    }
    this.admins.practiceLeader = this.getUserData(items[0]);
  }

  private getGroupLeader(items: any[]) {
    if (items.length == 0) {
      this.admins.groupLeader = null;
      return;
    }
    this.admins.groupLeader = this.getUserData(items[0]);
  }

  private getPracticeAdministrator(items: any[]) {
    if (items.length == 0) {
      this.admins.practiceLeader = null;
      return;
    }
    this.admins.practiceLeader = this.getUserData(items[0]);
  }


  private getMarketingMember(items: any[]) {
    if (items.length == 0) {
      this.admins.marketingMember = null;
      return;
    }
    this.admins.marketingMember = this.getUserData(items[0]);
  }

  private showMarketingDept(ev: React.FormEvent<HTMLElement>, isChecked: boolean) {
    this.setState({
      hasMarketingMember: isChecked
    });
  }

  public render(): React.ReactElement<ICreateO365GroupsProps> {


    const onChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
      if (item) {
        let formData = this.state.formData;
        formData[event.target["id"]] = item.key;
        this.setState({
          formData
        });
      }
    };

    function choiceChange(ev: React.FormEvent<HTMLInputElement>, option: IChoiceGroupOption): void {
      this.setState({
        createFolder: option.key == '1'
      });
    }

    const countryCode = [
      { key: 'USA', text: 'USA' }
    ];

    const companyCode = [
      { key: 'CTS', text: 'CTS' }
    ];

    const groupCode = [
      { key: 'G001', text: 'G001' }
    ];

    const projNumber = [
      { key: 'PN001', text: 'PN001' }
    ];

    const visibility = [
      { key: 'Public', text: 'Public' },
      { key: 'Private', text: 'Private' }
    ];

    const options: IChoiceGroupOption[] = [
      { key: '1', text: 'Create the common folder struture' },
      { key: '0', text: 'I will create my own folder structure in SharePoint' },
    ];


    function htmlMarketingDept() {
      return <Stack {...columnProps}>
        <PeoplePicker
          context={this.props.context}
          titleText="Practice Administrator"
          personSelectionLimit={1}
          groupName={""}
          showtooltip={true}
          isRequired={false}
          disabled={false}
          selectedItems={this.getPracticeAdministrator.bind(this)}
          showHiddenInUI={false}
          principalTypes={[PrincipalType.User]}
          resolveDelay={1000} />
      </Stack>;
    }


    const stackTokens = { childrenGap:15 };
    const stackStyles: Partial<IStackStyles> = { root: { width: 720 } };
    const columnProps: Partial<IStackProps> = {
      tokens: { childrenGap: 0 },
      styles: { root: { width: 720 } },
    };
    const dropDowncolumnProps: Partial<IStackProps> = {
      tokens: { childrenGap: 5 },
      styles: { root: { width: 100 } },
    };
    const shortdropDowncolumnProps: Partial<IStackProps> = {
      tokens: { childrenGap: 5 },
      styles: { root: { width: 155 } },
    };
    const visibilitycolumnProps: Partial<IStackProps> = {
      tokens: { childrenGap: 0 },
      styles: { root: { width: 360} },
    };

    return (
      <div>

        <Stack  horizontal tokens={stackTokens} styles={stackStyles} >
          <Stack {...dropDowncolumnProps} >
            <Dropdown
              id="countryCode"
              placeholder="Select"
              label="Country Code"
              onChange={onChange}
              options={countryCode}
            />
            </Stack>
            <Stack {...dropDowncolumnProps} >
            <Dropdown
              id="companyCode"
              placeholder="Select"
              label="Company Code"
              onChange={onChange}
              options={companyCode}
            />
      </Stack>
      <Stack {...dropDowncolumnProps} >
            <Dropdown
              id="groupCode"
              placeholder="Select"
              label="Group Code"
              onChange={onChange}
              options={groupCode}
            />
            </Stack>
            <Stack {...dropDowncolumnProps} >
            <Dropdown
              id="projectNumber"
              placeholder="Select"
              label="Project Number"
              onChange={onChange}
              options={projNumber}
            />
          </Stack>
          <Stack {...dropDowncolumnProps} >
            <TextField label="Task Number" name="taskNumber" onChange={(e) => this.inputChangeHandler.call(this, e)} value={this.state.formData.taskNumber} />
          </Stack>
          <Stack {...shortdropDowncolumnProps} >
            <TextField label="Short Description" name="shortDescription" onChange={(e) => this.inputChangeHandler.call(this, e)} value={this.state.formData.shortDescription} />
          </Stack>
   
          </Stack>
          <Stack {...columnProps} >
          <Checkbox label="Promotional Project" onChange={this.showMarketingDept.bind(this)} /> 
          </Stack>

          <Stack {...columnProps} >
          <TextField label="Description" multiline rows={3} name="description" onChange={(e) => this.inputChangeHandler.call(this, e)} value={this.state.formData.description} />
          </Stack>

        <Stack  horizontal tokens={stackTokens} styles={stackStyles} >
        <Stack {...visibilitycolumnProps} >
        <Dropdown
              id="visibility"
              placeholder="Select"
              label="Visibility"
              onChange={onChange}
              options={visibility} 
            /> 
            </Stack>
            <Stack {...visibilitycolumnProps} > <Checkbox label="Send mail to members" className={"common-form-group"}  /></Stack></Stack>
           

            <Stack  horizontal tokens={stackTokens} styles={stackStyles} >
          <Stack {...visibilitycolumnProps} >
            <PeoplePicker  
              context={this.props.context}
              titleText="Project Leader"
              personSelectionLimit={1}
              groupName={""}
              showtooltip={true}
              isRequired={false}
              disabled={false}
              selectedItems={this.getPracticeLeader.bind(this)}
              showHiddenInUI={false}
              principalTypes={[PrincipalType.User]}
              resolveDelay={1000} 
             
              />
              </Stack>
              <Stack {...visibilitycolumnProps} >
<PeoplePicker
              context={this.props.context}
              titleText="Group Leader"
              personSelectionLimit={1}
              groupName={""}
              showtooltip={true}
              isRequired={false}
              disabled={false}
              selectedItems={this.getGroupLeader.bind(this)}
              showHiddenInUI={false}
              principalTypes={[PrincipalType.User]}
              resolveDelay={1000} />
</Stack></Stack>
<Stack  horizontal tokens={stackTokens} styles={stackStyles} >
<Stack {...visibilitycolumnProps} >
<PeoplePicker
              context={this.props.context}
              titleText="Practice Administrator"
              personSelectionLimit={1}
              groupName={""}
              showtooltip={true}
              isRequired={false}
              disabled={false}
              selectedItems={this.getPracticeAdministrator.bind(this)}
              showHiddenInUI={false}
              principalTypes={[PrincipalType.User]}
              resolveDelay={1000} />
</Stack>
{
            this.state.hasMarketingMember ? 
            <Stack {...visibilitycolumnProps} >
              <PeoplePicker
                context={this.props.context}
                titleText="Marketing Dept member"
                personSelectionLimit={1}
                groupName={""}
                showtooltip={true}
                isRequired={false}
                disabled={false}
                selectedItems={this.getMarketingMember.bind(this)}
                showHiddenInUI={false}
                principalTypes={[PrincipalType.User]}
                resolveDelay={1000} /></Stack>
             : null
          }
</Stack>



           <Stack {...columnProps} >
           <PeoplePicker
              context={this.props.context}
              titleText="Members"
              personSelectionLimit={100}
              groupName={""}
              showtooltip={true}
              isRequired={false}
              disabled={false}
              selectedItems={this._getPeoplePickerItems.bind(this)}
              showHiddenInUI={false}
              principalTypes={[PrincipalType.User]}
              resolveDelay={1000}  />
           </Stack>
         

<Stack {...columnProps} >
<ChoiceGroup defaultSelectedKey="1" options={options} onChange={choiceChange.bind(this)} label="Pick one" />
</Stack>   

<PrimaryButton text="Submit" onClick={this.formHandler.bind(this)} />


          
       {/* </Stack> */}

        {/* <Stack horizontal tokens={stackTokens} styles={stackStyles}>
          <Stack {...columnProps}>
            <PeoplePicker
              context={this.props.context}
              titleText="Group Leader"
              personSelectionLimit={1}
              groupName={""}
              showtooltip={true}
              isRequired={false}
              disabled={false}
              selectedItems={this.getGroupLeader.bind(this)}
              showHiddenInUI={false}
              principalTypes={[PrincipalType.User]}
              resolveDelay={1000} />
          </Stack>

          <Stack {...columnProps}>
            <PeoplePicker
              context={this.props.context}
              titleText="Practice Administrator"
              personSelectionLimit={1}
              groupName={""}
              showtooltip={true}
              isRequired={false}
              disabled={false}
              selectedItems={this.getPracticeAdministrator.bind(this)}
              showHiddenInUI={false}
              principalTypes={[PrincipalType.User]}
              resolveDelay={1000} />
          </Stack>
        </Stack>



        <Stack horizontal tokens={stackTokens} styles={stackStyles}>
          <Stack {...columnProps}>
            <Checkbox label="Promotional Project" onChange={this.showMarketingDept.bind(this)} />
          </Stack>

          {
            this.state.hasMarketingMember ? <Stack {...columnProps}>
              <PeoplePicker
                context={this.props.context}
                titleText="Marketing Dept member"
                personSelectionLimit={1}
                groupName={""}
                showtooltip={true}
                isRequired={false}
                disabled={false}
                selectedItems={this.getMarketingMember.bind(this)}
                showHiddenInUI={false}
                principalTypes={[PrincipalType.User]}
                resolveDelay={1000} />
            </Stack> : null
          }

        </Stack>


        <Stack horizontal tokens={stackTokens} styles={stackStyles}>
          <Stack {...columnProps}>
            <Dropdown
              id="countryCode"
              placeholder="Select"
              label="Country Code"
              onChange={onChange}
              options={countryCode}
            />
          </Stack>

          <Stack {...columnProps}>
            <Dropdown
              id="companyCode"
              placeholder="Select"
              label="Company Code"
              onChange={onChange}
              options={companyCode}
            />
          </Stack>

          <Stack {...columnProps}>
            <Dropdown
              id="groupCode"
              placeholder="Select"
              label="Group Code"
              onChange={onChange}
              options={groupCode}
            />
          </Stack>

          <Stack {...columnProps}>
            <Dropdown
              id="projectNumber"
              placeholder="Select"
              label="Project Number"
              onChange={onChange}
              options={projNumber}
            />
          </Stack>

          <Stack {...columnProps}>
            <TextField label="Task Number" name="taskNumber" onChange={(e) => this.inputChangeHandler.call(this, e)} value={this.state.formData.taskNumber} />
          </Stack>

          <Stack {...columnProps}>
            <TextField label="Short Description" name="shortDescription" onChange={(e) => this.inputChangeHandler.call(this, e)} value={this.state.formData.shortDescription} />
          </Stack>

        </Stack>

        <Stack horizontal tokens={stackTokens} styles={stackStyles}>
          <Stack {...columnProps}>
            <Dropdown
              id="visibility"
              placeholder="Select"
              label="Visibility"
              onChange={onChange}
              options={visibility}
            />
          </Stack>
        </Stack>

        <Stack horizontal tokens={stackTokens} styles={stackStyles}>
          <Stack {...columnProps}>
            <TextField label="Description" multiline rows={3} name="description" onChange={(e) => this.inputChangeHandler.call(this, e)} value={this.state.formData.description} />
          </Stack>
        </Stack>

        <Stack horizontal tokens={stackTokens} styles={stackStyles}>
          <Stack {...columnProps}>
            <Checkbox label="Send mail to members" />
          </Stack>
        </Stack>

        <Stack horizontal tokens={stackTokens} styles={stackStyles}>
          <ChoiceGroup defaultSelectedKey="1" options={options} onChange={choiceChange.bind(this)} label="Pick one" />
        </Stack>


        <Stack horizontal tokens={stackTokens} styles={stackStyles}>
          <PrimaryButton text="Submit" onClick={this.formHandler.bind(this)} />
        </Stack> */}




      </div>
    );
  }
}
