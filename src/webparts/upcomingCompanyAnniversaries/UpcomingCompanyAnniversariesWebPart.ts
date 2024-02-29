import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
  PropertyPaneTextField,
  PropertyPaneDropdownOptionType,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './UpcomingCompanyAnniversariesWebPart.module.scss';
import * as strings from 'UpcomingCompanyAnniversariesWebPartStrings';
import iconPerson from './assets/iconPerson.png';

interface ListItem {
  UserName: string;
  JoiningDate: string;
  Department: string;
  Company: string;
  ProfilePicture: {
    Url: string;
  }
  Email: string;
  RefreshedOn: string;
}
export interface IUpcomingCompanyAnniversariesWebPartProps {
  description: string;
  message: string;
  UserName: string;
  JoiningDate: string;
  Department: string;
  Company: string;
  ProfilePicture: {
    Url: string;
  }
  Email: string;
  RefreshedOn: string;
}

export default class UpcomingCompanyAnniversariesWebPart extends BaseClientSideWebPart<IUpcomingCompanyAnniversariesWebPartProps> {

  private userEmail: string = "";

  private async userDetails(): Promise<void> {
    // Ensure that you have access to the SPHttpClient
    const spHttpClient: SPHttpClient = this.context.spHttpClient;
  
    // Use try-catch to handle errors
    try {
      // Get the current user's information
      const response: SPHttpClientResponse = await spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties`, SPHttpClient.configurations.v1);
      const userProperties: any = await response.json();
  
      console.log("User Details:", userProperties);
  
      // Access the userPrincipalName from userProperties
      const userPrincipalNameProperty = userProperties.UserProfileProperties.find((property: any) => property.Key === 'SPS-UserPrincipalName');
  
      if (userPrincipalNameProperty) {
        this.userEmail = userPrincipalNameProperty.Value.toLowerCase(); 
        console.log('User Email using User Principal Name:', this.userEmail);
        // Now you can use this.userEmail as needed
      } else {
        console.error('User Principal Name not found in user properties');
      }
    } catch (error) {
      console.error('Error fetching user properties:', error);
    }
  } 

  public render(): void {
    this.userDetails().then(() => {
    const decodedDescription = decodeURIComponent(this.properties.description);
    console.log("Title: ",decodedDescription);
    this.domElement.innerHTML = `
      <div class="${styles.parentDiv}">
        <div id="upcomingBirthdays" class="${styles.upcomingBirthdays}">
          <h3>${decodedDescription}</h3>
        </div>
      </div>`;
      this._renderButtons();
    });
  }

  public getItemsFromSPList(listName: string): Promise<any[]> {
    return new Promise((resolve, reject) => {
      let open = indexedDB.open("MyDatabase", 1);
   
      open.onsuccess = function() {
        let db = open.result;
        let tx = db.transaction(`${listName}`, "readonly");
        let store = tx.objectStore(`${listName}`);
   
        let getAllRequest = store.getAll();
   
        getAllRequest.onsuccess = function() {
          resolve(getAllRequest.result);
        };
   
        getAllRequest.onerror = function() {
          reject(getAllRequest.error);
        };
      };
   
      open.onerror = function() {
        reject(open.error);
      };
    });
  }

  private _renderButtons(): void {
    const buttonsContainer: HTMLElement | null = this.domElement.querySelector('#upcomingBirthdays');

    const adminEmailSplit: string[] = this.userEmail.split('.admin@');
    if (this.userEmail.includes(".admin@")){
        console.log("Admin Email after split: ", adminEmailSplit);
    }
    const parts = this.userEmail.split('_');
    const secondPart = parts.length > 1 ? parts[1] : '';
    const otherUsersSplit =  secondPart.split('.com')[0];
    if (this.userEmail.includes("_")){
        console.log("User's company after split: ", otherUsersSplit);
    }

    this.getItemsFromSPList('SPList')
    .then((items: ListItem[]) => {
      console.log("All items retrieved:", items);
        let buttonsCreated = 0; // Variable to keep track of the number of buttons created
        if (items && items.length > 0) {
          const today = new Date();
          const sixDaysLater = new Date(today);
          sixDaysLater.setDate(today.getDate() + 6);
          // console.log("today: ", today);
          // console.log("sixDaysLater: ", sixDaysLater);

          const formattedToday = `${(today.getMonth() + 1).toString().padStart(2, '0')}-${today.getDate().toString().padStart(2, '0')}`;
          const formattedSixDaysLater = `${(sixDaysLater.getMonth() + 1).toString().padStart(2, '0')}-${sixDaysLater.getDate().toString().padStart(2, '0')}`;

          // console.log("formattedToday: ", formattedToday);
          // console.log("formattedSixDaysLater: ", formattedSixDaysLater);


          // Filter items based on month and date on the client side
          const filteredItemsWithDate = items.filter(item => {
            if (!item.JoiningDate || !item.UserName) {
              return false;
            }
          
            let itemDate = this.adjustDateForTimeZone(item.JoiningDate);
            const itemYear = itemDate.getFullYear();
            
          
            // Convert the Date object to a string in the format MM-DD
            const monthAndDay = (itemDate.getMonth() + 1).toString().padStart(2, '0') + '-' +
                                itemDate.getDate().toString().padStart(2, '0');
            
            // console.log("itemDate for user", item.UserName, "is: ", monthAndDay);
            // console.log("itemDate >= formattedToday:", monthAndDay >= formattedToday);
            // console.log("itemDate <= formattedSixDaysLater:", monthAndDay <= formattedSixDaysLater);
            // // Check if the Joining Date year is not greater than or equal to the current year
            // console.log("itemYear:", itemYear);
            // console.log("Current year:", new Date().getFullYear())
            if (itemYear >= new Date().getFullYear()) {
              return false;
            }
            
            return monthAndDay >= formattedToday && monthAndDay <= formattedSixDaysLater;
          });
          

          console.log('filteredItems: ',filteredItemsWithDate);
          
          const sortedItems = this.sortItems(filteredItemsWithDate);
          console.log('Sorted Items:', sortedItems);
          sortedItems.forEach((item: IUpcomingCompanyAnniversariesWebPartProps) => {
            if(!item.Company){
              item.Company = " ";
            }
                if((this.userEmail.includes("@"+item.Company.toLowerCase()+".") && !this.userEmail.includes(".admin@") && !otherUsersSplit) || (this.userEmail.includes(".admin@") && adminEmailSplit.includes("@"+item.Company.toLowerCase()+".")) || (otherUsersSplit.length >= 0 && otherUsersSplit.includes(item.Company.toLowerCase()))){
                    const buttonDiv: HTMLDivElement = document.createElement('div');
                    buttonDiv.classList.add(styles.innerContents);

                    const profileSection: HTMLDivElement = document.createElement('div');
                    profileSection.classList.add(styles.profileSection); 

                    const imgBox: HTMLDivElement = document.createElement('div');
                    imgBox.classList.add(styles.imgBox); 
                    const img: HTMLImageElement = document.createElement('img');
                    img.src = item.ProfilePicture && item.ProfilePicture.Url ? item.ProfilePicture.Url : iconPerson;

                    // Add an error event listener to handle image load errors
                    img.addEventListener('error', () => {
                      img.src = iconPerson; // Set to default image if an error occurs
                    });

                    imgBox.appendChild(img);
                    
                    const nameDiv: HTMLDivElement = document.createElement('div');
                    nameDiv.classList.add(styles.name); 
                    const h5: HTMLHeadingElement = document.createElement('h5');
                    h5.textContent = item.UserName;
                    const divCompany: HTMLDivElement = document.createElement('div');
                    divCompany.textContent = item.Company;
                    divCompany.classList.add(styles.text); 
                    const divDept: HTMLDivElement = document.createElement('div');
                    divDept.classList.add(styles.text); 
                    divDept.textContent = item.Department;

                    
                    const fullDateString = this.adjustDateForTimeZone(item.JoiningDate).toISOString().substring(0, 10);
                    // const fullDate = this.fullDate(fullDateString);
                    const monthAndDateString = this.adjustDateForTimeZone(item.JoiningDate).toISOString().substring(5, 10);

                    const monthAndDateText = this.monthAndDate(monthAndDateString);
                    // const birthdayText = this.properties.displayDate
                    // ? `Date of Joining: ${fullDate}` 
                    // : 'Date of Joining: DD-MM-YYYY';
                
                    // const birthdayElement: HTMLDivElement = document.createElement('div');
                    // birthdayElement.classList.add(styles.text); 
                    // birthdayElement.textContent = birthdayText;
                    // birthdayElement.style.display = this.properties.displayDate ? 'block' : 'none';
                    
                    const formattedYears = this.calculateNoOfYears(fullDateString);
                    const noOfYears = `Celebrating ${formattedYears} on ${monthAndDateText}!`;
                
                    const divYears: HTMLDivElement = document.createElement('div');
                    divYears.classList.add(styles.text); 
                    divYears.textContent = noOfYears;
                    // divYears.style.display = this.properties.displayYears ? 'block' : 'none';

                    nameDiv.appendChild(h5);  
                    nameDiv.appendChild(divDept);
                    // nameDiv.appendChild(birthdayElement);
                    nameDiv.appendChild(divYears);
                    profileSection.appendChild(imgBox); 
                    profileSection.appendChild(nameDiv);

                    const chatBtn: HTMLButtonElement = document.createElement('button');
                    chatBtn.classList.add(styles.chatBtn);
                    chatBtn.textContent = "Chat";
                    chatBtn.onclick = () =>{
                      window.open(`msteams://teams.microsoft.com/l/chat/0/0?users=${item.Email}&message=${this.properties.message}`, '_blank');
                    };

                    buttonDiv.appendChild(profileSection);
                    buttonDiv.appendChild(chatBtn);

                    buttonsContainer!.appendChild(buttonDiv); // Append the button to the buttons container
                    buttonsCreated++; // Increment the count of buttons created
                } 
            });
            if (buttonsCreated === 0) {
              const noDataMessage: HTMLDivElement = document.createElement('div');
              noDataMessage.classList.add(styles.innerContents);
              noDataMessage.textContent = 'There are no company anniversaries this week';
              console.log("No new joinees the last 30 days");
              buttonsContainer!.appendChild(noDataMessage);// Non-null assertion operator
            }
        } else {
            const noDataMessage: HTMLDivElement = document.createElement('div');
            noDataMessage.classList.add(styles.innerContents);
            noDataMessage.textContent = 'There are no company anniveraries this week.';
            buttonsContainer!.appendChild(noDataMessage);// Non-null assertion operator
        }
    })
    .catch(error => {
        console.error("Error fetching user data: ", error);
    });
}

private adjustDateForTimeZone(dateString) {
  // Add your timezone adjustment logic here
  const timeZoneDifferenceHours = 5; // Adjust this based on your timezone
  const timeZoneDifferenceMinutes = 30;

  const date = new Date(dateString);
  date.setHours(date.getHours() + timeZoneDifferenceHours);
  date.setMinutes(date.getMinutes() + timeZoneDifferenceMinutes);

  return date;
}

private sortItems(items: ListItem[]): ListItem[] {
  // Sort the items by DateofBirth in ascending order (earliest first)
  return items.sort((a, b) => {
    // Create new Date objects with a fixed year
    const dateA = new Date(`2000-${new Date(a.JoiningDate).getMonth() + 1}-${new Date(a.JoiningDate).getDate()}`);
    const dateB = new Date(`2000-${new Date(b.JoiningDate).getMonth() + 1}-${new Date(b.JoiningDate).getDate()}`);

    // Compare the dates
    return dateA.getTime() - dateB.getTime();
  });
}
private calculateNoOfYears(dateOfJoining: string): string {
  const joiningDate = new Date(dateOfJoining);
  const currentDate = new Date();

  // Check if the current date is greater than or equal to the joining date
  if (currentDate.getTime() >= joiningDate.getTime()) {
    const years = currentDate.getFullYear() - joiningDate.getFullYear();
    return `${years} ${years === 1 ? 'year' : 'years'}`;
  } else {
    return "Joining Date is greater than or equal to the current date";
  }
}

private fullDate(dateOfJoining: string): string {
  const [year, month, day] = dateOfJoining.split('-');
  return `${day}-${month}-${year}`;
}

private monthAndDate(dateOfJoining: string): string {
  const [month, day] = dateOfJoining.split('-');
  return `${day}-${month}`;
}

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('description', {
                  label: "Title For The Application",
                }),
                PropertyPaneTextField('message', {
                  placeholder: "Happy Company Anniversary!",
                  label: "Message for birthday wishes",
                }),
                // PropertyPaneToggle('displayDate', {
                //   label: 'Display Full Date',
                //   onText: 'Yes',
                //   offText: 'No',
                // }),
                // PropertyPaneToggle('displayYears', {
                //   label: 'Display Number of Years',
                //   onText: 'Yes',
                //   offText: 'No',
                // }),
              ]
            }
          ]
        }
      ]
    };
  }
}
