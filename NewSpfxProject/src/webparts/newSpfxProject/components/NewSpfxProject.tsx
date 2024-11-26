import * as React from 'react';
import type { INewSpfxProjectProps } from './INewSpfxProjectProps';
import { INewSpfxProjectState } from './INewSpfxProjectState';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/items/get-all";
import { Dropdown, IDropdownOption } from '@fluentui/react';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { List } from 'office-ui-fabric-react/lib/List';
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react';

export default class NewSpfxProject extends React.Component<INewSpfxProjectProps, INewSpfxProjectState> {
  constructor(props: INewSpfxProjectProps) {
    super(props);
    this.state = {
      ID: '',
      Title: '',
      Description: '',
      Address: '',
      data: [],
      Choice: '',
      ChoiceOptions: [],
      selectedUsers: [], 
      Lookup: '',
      LookupOptions: [],
    };
  }
  
  public async componentDidMount(): Promise<void> {
    try {
      this.getList();
      await this.fetchChoiceOptions();
      await this.fetchLookupOptions();
    } catch (error) {
      console.log("ComponentDidmount : error ", error);
    }  
  }

  public handlePeoplePickerChange = (people: any) => {
    this.setState({ selectedUsers: people });
  };

  getList = async (): Promise<void> => {
    try {
      const sp = spfi().using(SPFx(this.props.context));
      const list = sp.web.lists.getByTitle("Test");
      
      const items = await list.items.select("ID", "Title", "Description", "Address", "Choice", "selectedUsers/Id", "selectedUsers/Title", "selectedUsers/EMail", "Lookup/Title", "Lookup/Cost")
        .expand("selectedUsers", "Lookup")
        .getAll();
        
      console.log("Retrieved items:", items);

      const PeopleTitle = items.map((item) => item.selectedUsers !== undefined ? item.selectedUsers.Title : null);
      const PeopleEmail = items.map((item) => item.selectedUsers !== undefined ? item.selectedUsers.EMail : null);
  
      const itemsWithPeopleTitleAndEmail: any = items.map((item, index) => {
        return {...item,
          UserTitle: PeopleTitle[index],
          UserEMail: PeopleEmail[index],
        };
      });
  
      console.log("Items Merge:", itemsWithPeopleTitleAndEmail);
  
      this.setState({
        data: itemsWithPeopleTitleAndEmail
      });
    } catch (error) {
      console.error("Error in getList:", error);
    }
  };
  
  public async fetchChoiceOptions(): Promise<void> {
    const sp: any = spfi().using(SPFx(this.props.context));
    const fieldSchema = await sp.web.lists.getByTitle("Test").fields.getByInternalNameOrTitle("Choice")();
    console.log("fieldSchema", fieldSchema);
    if (fieldSchema && fieldSchema.Choices) {
      this.setState({ ChoiceOptions: fieldSchema.Choices });
    }
  }

  public async fetchLookupOptions(): Promise<void> {
    try {
      const sp = spfi().using(SPFx(this.props.context));
      const lookupList = sp.web.lists.getByTitle("Project List");
      const items = await lookupList.items.select("Cost").getAll(); // Fetching all items from the "Test" list
      console.log("Lookup items:", items); // Log retrieved items for debugging
      const options: IDropdownOption[] = items.map(item => {
        return {
          key: item.Cost, 
          text: item.Cost,
        };
      });
      console.log("Lookup options:", options); // Log dropdown options for debugging
      this.setState({ LookupOptions: options });
    } catch (error) {
      console.error("Error fetching lookup options:", error);
      // Handle the error, display a message, or fallback behavior
    }
  }
  
  handleChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const name = event.target.name;
    const value = event.target.value;

    this.setState({
      [name]: value
    } as unknown as Pick<INewSpfxProjectState, keyof INewSpfxProjectState>);
  }

  handleSubmit = async (): Promise<void> => {
    try {
      const { Title, Description, Address, Choice, Lookup } = this.state;
      const sp =  spfi().using(SPFx(this.props.context));
      const AddedList = await sp.web.lists.getByTitle("Test").items.add({
        'Title': Title,
        'Description': Description,
        'Address': Address,
        'Choice': Choice,
        'Lookup': Lookup,
      });
      alert('List item added successfully');
      await this.getList();
    } catch (error) {
      console.error("Error in handleSubmit adding list item:", error);
    }
  }

  handleSelectChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number): void => {
    const value = option ? (option.key as string) : ''; 

    this.setState({
      Choice: value
    });
  }

  handleLookupChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number): void => {
    const value = option ? (option.key as string) : ''; 

    this.setState({
      Lookup: value
    });
  }

  handleDelete = async (Id: number): Promise<void> => {
    const sp: any = spfi().using(SPFx(this.props.context));
    const list = sp.web.lists.getByTitle("Test");
    await list.items.getById(Id).delete();
    alert('List item deleted successfully');
    await this.getList();
  }
  
  handleEdit = (item: any): void => {

    this.setState({
      ID: item.ID,
      Title: item.Title,
      Description: item.Description,
      Address: item.Address,
      Choice: item.Choice,
      Lookup: item.Lookup?.Cost,
      selectedUsers: item.selectedUsers?.EMail,
    });
    console.log("sato...", item.selectedUsers?.EMail);
  };
  
  handleUpdate = async (selectedPerson: []): Promise<void> => {
    const { ID, Title, Description, Address, Choice, data, Lookup, selectedUsers } = this.state;
    const sp = spfi().using(SPFx(this.props.context));

    const matchingIds = data.filter((item: { ID: React.Key }) => item.ID === ID).map((item: { Id: number }) => item.Id);

    const itemId = matchingIds.length > 0 ? matchingIds[0] : undefined;

    const user = selectedUsers.id;
    if (itemId) {
      try {
        const list = await sp.web.lists.getByTitle("Test").items.getById(itemId).update({
                  'Title': Title,
                  'Description': Description,
                  'Address': Address,
                  'Choice': Choice,
                  // 'Lookup': Lookup,
                  'selectedUsersId': user.id,
                });
                console.log(selectedUsers)
                console.log(Lookup)
        this.getList();
        this.setState({ 
          Title: '',
          Description: '',
          Address: '',
          Choice: '',
          // Lookup: '',
          selectedUsers: [],
         });
        alert('Updated Successfully');
      } catch (error) {
        console.error('Error adding item:', error);
        alert('Failed to add item. Please try again.');
      }
    } else {
      alert('Please fill all the fields');
    }
  };
  
  public render(): React.ReactElement<INewSpfxProjectProps> {
    return (
      <>
        <h1>List</h1>
        
        <List items={this.state.data} onRenderCell={(item, index) => (
          <div>
            <ul>
              <li key={item.Id}>{item.Title}</li>
              <li key={item.Id}>{item.Address}</li>
              <li key={item.Id}>{item.Description}</li>
              <li key={item.Id}>{item.Choice}</li>
              <li key={item.Id}>{item.Lookup && item.Lookup.Title}</li>
              <li key={item.Id}>{item.Lookup && item.Lookup.Cost}</li>
              {item.selectedUsers && (
                <>
                  <li> UserTitle: {item.selectedUsers.Title}</li>
                  <li>UserEMail: {item.selectedUsers.EMail}</li>
                </>
              )}
            </ul>
            <DefaultButton onClick={() => this.handleEdit(item)}>Edit</DefaultButton>
            <PrimaryButton text="Delete" onClick={() => this.handleDelete(item.Id)} />
          </div>
        )} />

        <div>
          <form onSubmit={this.handleSubmit}>
            <label htmlFor="Title">Title:
              <TextField
                name='Title'
                value={this.state.Title}
                onChange={this.handleChange}
              />
            </label><br />
            <label htmlFor="Description">Description:
              <TextField
                name='Description'
                value={this.state.Description}
                onChange={this.handleChange}
              />
            </label><br />
            <label htmlFor="Address">Address:
              <TextField
                name='Address'
                value={this.state.Address}
                onChange={this.handleChange}
              />
            </label>
            <Dropdown
              label="Lookup:"
              onChange={this.handleLookupChange}
              selectedKey={this.state.Lookup}
              options={[
                { key: '', text: 'Select Lookup' },
                ...this.state.LookupOptions
              ]}
            />
            <Dropdown
              label="Choice:"
              selectedKey={this.state.Choice}
              onChange={this.handleSelectChange}
              options={[
                { key: '', text: 'Select Choice' },
                ...this.state.ChoiceOptions.map(option => ({ key: option, text: option }))
              ]}
            />
            <PeoplePicker
              context={this.props.context}
              titleText="People Picker"
              personSelectionLimit={1}
              showtooltip={true}
              disabled={false}
              onChange={this.handlePeoplePickerChange}
              principalTypes={[PrincipalType.User]}
              defaultSelectedUsers={[this.state.selectedUsers]}
            />
            <PrimaryButton type="submit">Submit</PrimaryButton>
            <DefaultButton type="button" onClick={() => { this.handleUpdate(this.state.selectedUsers) }}>Update</DefaultButton>
          </form>
        </div>
      </>
    )
  }
}
