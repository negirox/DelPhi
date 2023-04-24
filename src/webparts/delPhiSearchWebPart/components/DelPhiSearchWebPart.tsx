import * as React from 'react';
import styles from './DelPhiSearchWebPart.module.scss';
import { IDelPhiSearchWebPartProps } from './IDelPhiSearchWebPartProps';
/* import { getSP } from '../../../pnpjsConfig';
import { SPFI } from '@pnp/sp'; */
import { HelperUtils } from '../../../utils/HelperUtils';
import { UserService } from '../../../services/UserService';
import * as autocompleteutils from './autohelpers';
//import { AadHttpClient } from "@microsoft/sp-http";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import { SearchUserModel } from '../../../models/SearchUserModel';
import { IDelphiSearchState } from './IDelphiSearchState';
import { Stack } from '@fluentui/react/lib/Stack';
import { IPersonaProps, IPersonaSharedProps, IPersonaStyles, Persona, PersonaPresence, PersonaSize } from '@fluentui/react/lib/Persona';
import { Icon, IIconStyles } from '@fluentui/react/lib/Icon';
import { TestImages } from '@fluentui/example-data';

require('./autocomplete.css');
const personaStyles: Partial<IPersonaStyles> = { root: { margin: '0 0 10px 0' } };
const iconStyles: Partial<IIconStyles> = { root: { marginRight: 5 } };
const ColoredLine = ({ color }) => (
  <hr
    style={{
      color: color,
      backgroundColor: color,
      height: 5
    }}
  />
);
export default class DelPhiSearchWebPart extends React.Component<IDelPhiSearchWebPartProps, IDelphiSearchState> {
  //private _sp: SPFI;
  private _userCollection: Array<SearchUserModel>;
  constructor(props: IDelPhiSearchWebPartProps, state: IDelphiSearchState) {
    super(props);
    this.state = {
      items: new Array<SearchUserModel>(),
      searchText: '',
      searchResults: new Array<SearchUserModel>()
    }
    //this._sp = getSP();
    this.getUsers('h');
   // this._searchWithAad();
    this._userCollection = new Array<SearchUserModel>();
    this.mockJson();
    this.searchUsers = this.searchUsers.bind(this);
  }
  private mockJson(): void {
    let user: SearchUserModel = new SearchUserModel();
    user.DisplayName = 'Mukesh Singh Negi';
    user.Manager = 'Jagati Rishi';
    user.EmailAddress = 'mnegi@delphime.com';
    user.EmployeeId = '101';
    user.ProfilePic = TestImages.personaMale;
    user.Designation = 'Senior Software Engineer';
    user.MobilePhone = '9044988629';
    user.State = 'U.P.';
    user.JobTitle = 'Senior Software Engineer';
    this._userCollection.push(user);
    for (let index = 0; index < 1000; index++) {
      user = new SearchUserModel();
      user.DisplayName = this.genearateString(7, 0);
      user.Manager = this.genearateString(7, 0);
      user.EmailAddress = this.makeEmail();
      user.EmployeeId = this.genearateString(7, 1);
      user.ProfilePic = index % 2 === 0 ? TestImages.personaFemale : TestImages.personaMale;
      user.Designation = this.genearateString(7, 0);
      user.MobilePhone = this.genearateString(10, 1);
      user.State = `${this.genearateString(1, 0)}.${this.genearateString(1, 0)}`;
      user.JobTitle = this.genearateString(10, 0);
      this._userCollection.push(user);
    }
  }
  private genearateString(length: number, type: number): string {
    let result = '';
    let characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz';
    if (type === 1) {
      characters = '0123456789';
      length = characters.length - 1;
    }
    const charactersLength = characters.length;
    let counter = 0;
    while (counter < length) {
      result += characters.charAt(Math.floor(Math.random() * charactersLength));
      counter += 1;
    }
    return result;
  }
  private makeEmail(): string {
    let strValues = "aabcdefghijklmnopqrstuvwxyz0123456789";
    let strEmail = '';
    let strTmp;
    for (let i = 0; i < 10; i++) {
      strTmp = strValues.charAt(Math.round(strValues.length * Math.random()));
      strEmail = strEmail + strTmp;
    }
    strEmail = strEmail + "@delphime.com";
    return strEmail;
  }
  componentDidMount(): void {
    let countries = this._userCollection.map(x => { return x.DisplayName });
    autocompleteutils.autocomplete(document.getElementById("searchUser"), countries);
  }
  private async getUsers(searchText: string): Promise<void> {
    let results;
    if (!HelperUtils.isEmpty(searchText)) {
      await UserService.GetAllUsers();
      await this.props.context.msGraphClientFactory.getClient('3')
        .then((client): void => {
          // get information about the current user from the Microsoft Graph
          client
            .api('/me')
            .get((error: any, response: any, rawResponse?: any) => {
              console.log(response);
              console.log(rawResponse);
              console.log(error);
            });
        });
    }
    console.log(results);
  }
  private searchUsers(): void {
    const text = window['searchText'] || '';
    if (text === '') {
      this.setState({
        searchResults: new Array<SearchUserModel>(),
        searchText: text
      });
      return;
    }
    const searchResults = this._userCollection.filter(
      x => {
        return x.DisplayName.toUpperCase().indexOf(text.toUpperCase()) > -1;
      }
    );

    this.setState({
      searchResults: searchResults,
      searchText: text
    });

  }
/*   private _searchWithAad = (): void => {
    // Log the current operation
    console.log("Using _searchWithAad() method");
    let searchFor = 'Muk';
    // Using Graph here, but any 1st or 3rd party REST API that requires Azure AD auth can be used here.
    this.props.context.aadHttpClientFactory
      .getClient("https://graph.microsoft.com")
      .then((client: AadHttpClient) => {
        // Search for the users with givenName, surname, or displayName equal to the searchFor value
        return client
          .get(
            `https://graph.microsoft.com/v1.0/users?$select=displayName,mail,userPrincipalName&$filter=(givenName%20eq%20'${escape(searchFor)}')%20or%20(surname%20eq%20'${escape(searchFor)}')%20or%20(displayName%20eq%20'${escape(searchFor)}')`,
            AadHttpClient.configurations.v1
          );
      })
      .then(response => {
        return response.json();
      })
      .then(json => {

        // Prepare the output array
        let users: Array<any> = new Array<any>();

        // Log the result in the console for testing purposes
        console.log(json);

        // Map the JSON response to the output array
        json.value.map((item: any) => {
          users.push({
            displayName: item.displayName,
            mail: item.mail,
            userPrincipalName: item.userPrincipalName,
          });
        });

        // Update the component state accordingly to the result

      })
      .catch(error => {
        console.error(error);
      });
  } */
  private updateInputValue(evt: React.ChangeEvent<HTMLInputElement>): void {
    const val = evt.target.value;

    this.setState({
      searchText: val,
      searchResults: val === '' ? new Array<SearchUserModel>() : this.state.searchResults
    });
  }
  private _onRenderSecondaryText(props: IPersonaProps): JSX.Element {

    return (
      <div>
        <Icon iconName="Suitcase" styles={iconStyles} />
        {props.secondaryText}
      </div>
    );
  }
  private handleEnterKey(){
    const closeList = window['closeItems'] || null;
    if(closeList != null)
      closeList();
    this.searchUsers();
  }

  public render(): React.ReactElement<IDelPhiSearchWebPartProps> {

    return (
      <section className={`${styles.delPhiSearchWebPart}`}>
        <div className={styles.root}>
          <div className={styles.container}>
            <div className={styles.searchBoxRoot}>
              <input type="search" className={styles.inptSearch} id='searchUser' value={this.state.searchText}
                onChange={evt => this.updateInputValue(evt)}
                onKeyDown={(ev) => {
                  if (ev.key === 'Enter') {
                    // Do code here
                    this.handleEnterKey();
                    ev.preventDefault();
                  }
                }}
                placeholder="Search For Users..." />
              <button className={styles.btnSearch} >Search<Icon iconName='ProfileSearch' /></button>
              <Icon iconName='ProfileSearch' onClick={this.searchUsers} className={styles.searchIcon} />
            </div>
            <div style={{ minHeight: '195px', maxHeight: '195px', overflowY: 'auto' }}>
              {
                this.state.searchResults.map(x => {
                  const person: IPersonaSharedProps = {
                    imageUrl: x.ProfilePic,
                    imageInitials: this.getInitials(x.DisplayName),
                    text: x.DisplayName,
                    secondaryText: `${x.Designation}, Reports To : ${x.Manager}`,
                    tertiaryText: x.MobilePhone + ', ' + x.State,//email,manager,mobile,location
                    optionalText: x.EmailAddress,
                    showOverflowTooltip: true,
                    showSecondaryText: true,
                    presenceTitle: ''
                  };
                  return (
                    <Stack tokens={{ childrenGap: 10 }}>

                      <Persona
                        {...person}
                        size={PersonaSize.size100}
                        presence={PersonaPresence.offline}
                        onRenderSecondaryText={this._onRenderSecondaryText}
                        styles={personaStyles}
                        imageAlt={'Profile Pic of' + x.DisplayName}
                      />
                      <ColoredLine color="black" />
                    </Stack>
                  )
                })
              }
            </div>
          </div>
        </div>

      </section>
    );
  }
  private getInitials(DisplayName: string): string {
    let splitText = DisplayName?.split(' ');
    return splitText.length > 2 ? (splitText[1].charAt(0) + splitText[0].charAt(0)).toUpperCase() : splitText[0].charAt(0);
  }
}
