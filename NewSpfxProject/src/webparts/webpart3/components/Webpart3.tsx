import * as React from 'react';
// import styles from './Webpart3.module.scss';
import type { IWebpart3Props } from './IWebpart3Props';
import { IWebpart3State } from './IWebpart3State';
import { SPFx, spfi } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/items/get-all";
// import { escape } from '@microsoft/sp-lodash-subset';

export default class Webpart3 extends React.Component<IWebpart3Props, IWebpart3State ,{ items : any}> {

  constructor(props: IWebpart3Props){
  super(props);
  this.state = {
    Title: '',
    UserEMail: '',
    UserTitle: '',
    items: [],
    selectedUsers : [],
  };
}


  componentDidMount(): void {
    this.fetchItems();
  }

  public async fetchItems() {
    const sp = spfi().using(SPFx(this.props.context));
    const list = sp.web.lists.getByTitle("Test");
    const itemsResponse = await list.items.select("Title", "selectedUsers/Title", "selectedUsers/EMail").expand("selectedUsers").getAll();
  
    console.log("Query:", itemsResponse); // Log the query URL
  
    // Extract titles and emails from the people picker column and create new arrays
    const PeopleTitle = itemsResponse.map((item) => item.selectedUsers !== undefined ? item.selectedUsers.Title : null);
    const PeopleEmail = itemsResponse.map((item) => item.selectedUsers !== undefined ? item.selectedUsers.EMail : null);
  
    // Add PeopleTitle and PeopleEmail properties to each item in the itemsResponse array
    const itemsWithPeopleTitleAndEmail:any = itemsResponse.map((item, index) => {
      return {...item,
        UserTitle: PeopleTitle[index],
        UserEMail: PeopleEmail[index],
      };
    });
  
    console.log("Items Merge:", itemsWithPeopleTitleAndEmail); // Check the console for retrieved data
    console.log("People Title:", PeopleTitle); // Check the extracted PeopleTitle
    console.log("People Email:", PeopleEmail); // Check the extracted PeopleEmail
    console.log("Items response:", itemsResponse); // Check the console for retrieved data
  
    this.setState({ 
      items: itemsWithPeopleTitleAndEmail,
    });
  }
  

  public render(): React.ReactElement<IWebpart3Props> {
    return (
      <>
        {this.state.items.map((item: { Title: string; UserTitle?: string; UserEMail?: string; selectedUsers?: { Title: string; EMail: string } }, index: number) => (
          <div key={index}>
            <h6>Finally.... </h6>
            <h4>Title: {item.Title}</h4>
            {item.selectedUsers && (
              <>
                <h4>UserTitle: {item.UserTitle || item.selectedUsers.Title}</h4>
                <h4>UserEMail: {item.UserEMail || item.selectedUsers.EMail}</h4>
              </>
            )}
          </div>
        ))}
        <div>Hello I'm people picker </div>
      </>
    );
    
    
  }
}
