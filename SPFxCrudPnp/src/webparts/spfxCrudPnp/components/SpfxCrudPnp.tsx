import * as React from 'react';
import styles from './SpfxCrudPnp.module.scss';
import type { ISpfxCrudPnpProps } from './ISpfxCrudPnpProps';
import type { ISpfxCrudState } from './ISpfxCrudPnpState';
import { escape } from '@microsoft/sp-lodash-subset';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/items/get-all";
import { Dropdown } from 'office-ui-fabric-react';
import { IDropdownOption } from '@fluentui/react';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { List } from 'office-ui-fabric-react/lib/List';
import { DefaultButton } from 'office-ui-fabric-react';
import { PrimaryButton } from 'office-ui-fabric-react';
import { IAddNewList } from './IAddNewList';
import { SPFx, spfi } from '@pnp/sp';

export default class SpfxCrudPnp extends React.Component<ISpfxCrudPnpProps, {}> {
  constructor(props: ISpfxCrudPnpProps) {
    super(props)
  
    this.state = {
       ID : '',
       Title: '',
       Description: '',
       Address: '',
       data: [],
       Choice: '',
       ChoiceOptions: [],
       selectedUsers: [],
       Lookup: '',
    }
  }
  
public async componentDidMount(): Promise <void> {
  try {
    this.getList()
    await this.fetchChoiceOptions()
  } catch (error) {
    console.log("ComponentDidmount : error ", error)
  }  
}

  getList = async () => {
    const sp = spfi().using(SPFx(this.props.context));
    // const list = await sp.web.lists.getByTitle("Test").items.getAll()
    const sp1 = sp.web.lists.getByTitle("Test");
    const items = await sp1.items.select( "Title", "Description", "Address", "Choice", "Lookup/Title", "Lookup/Cost" ).expand("Lookup").getAll();
    console.log("Retrieved items:", items);
    this.setState({
      data: items
    })
  }
  public async fetchChoiceOptions(): Promise<void> {
      const sp: any = spfi().using(SPFx(this.props.context));
      const fieldSchema = await sp.web.lists.getByTitle("Test").fields.getByInternalNameOrTitle("Choice")();
      console.log("fieldSchema",fieldSchema);
      if (fieldSchema && fieldSchema.Choices) {
        this.setState({ ChoiceOptions: fieldSchema.Choices });
      }
  }

  handleChange = (event: React.ChangeEvent <HTMLInputElement>) => {
      const name = event.target.name;
      const value = event.target.value;

      this.setState({
          [name] : value
      } as unknown as Pick<IAddNewList, keyof IAddNewList>) 
  }

  handleSubmit = async(): Promise<void> => {
    try {
      const { Title, Description, Address, Choice, Lookup } = this.state as {
          Title: string;
          Description: string;
          Address:string;
          Choice: string;
          Lookup: string;
      }
      const sp =  spfi().using(SPFx(this.props.context));
      const AddedList = await sp.web.lists.getByTitle("Test").items.add({
        'Title': Title,
        'Description': Description,
        'Address': Address,
        'Choice': Choice,
        'Lookup': Lookup,
      })
      alert('List item added successfully');
      await this.getList();
    } catch (error) {
      console.error("Error in handleSubmit adding list item:", error);
    }
  }

  handleSelectChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number): void => {
    // Extract the selected value from the option parameter
    const value = option ? (option.key as string) : ''; // or whatever type your key is

    // Update the state
    this.setState({
        Choice: value
    });
}

  handleDelete = async (Id: number) => {
    const sp: any = spfi().using(SPFx(this.props.context));
    const list = sp.web.lists.getByTitle("Test");
    await list.items.getById(Id).delete();
    alert('List item deleted successfully');
    await this.getList()
  }
  
  handleEdit = async (item: any) => {
    // Populate form fields with the selected item's details
    this.setState({
      ID : item.ID,
      Title: item.Title,
      Description: item.Description,
      Address: item.Address,
      Choice: item.Choice
    });
  };

  handleUpdate = async (): Promise<void> => {
    const { ID, Title, Description, Address, Choice, data, selectedUsers } = this.state;
    const sp = spfi().using(SPFx(this.props.context)); 
    const itemId = data.find((item: { ID : React.Key }) => item.ID === ID)?.Id;

  
    if (itemId) {
      await sp.web.lists.getByTitle("Test").items.getById(itemId).update({
        Title: Title,
        Description: Description,
        Address: Address,
        Choice: Choice,
        selectedUsers: selectedUsers
      });
      // Clear form fields after update
      this.setState({
        ID: '',
        Title: '',
        Description: '',
        Address: '',
        Choice: '',
        selectedUsers: [],
      });
      // Refresh the list data
      await this.getList();
    } else {
      console.error("Item ID not found for update.");
    }
  };

  handlePeoplePickerChange = (people: []) => {
    this.setState({
        selectedUsers: people
    });
    console.log(this.state.selectedUsers);
}

  public render(): React.ReactElement<ISpfxCrudPnpProps> {
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
              <li key={item.Id}>{item.selectedUser}</li>
              <li key={item.Id}>{item.Lookup && item.Lookup.Title}</li>
              <li key={item.Id}>{item.Lookup.Cost}</li>

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
          <label htmlFor="Lookup">Lookup:
            <TextField
              name='Lookup'
              value={this.state.Lookup}
              onChange={this.handleChange}
            />
          </label>
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
                  personSelectionLimit={3}
                  groupName={""} // Leave it blank to fetch all the users
                  showtooltip={true}
                  disabled={false}
                  onChange={this.handlePeoplePickerChange}
                  principalTypes={[PrincipalType.User]}
                  defaultSelectedUsers={this.state.selectedUsers}
              /> <br />
            <PrimaryButton type="submit">Submit</PrimaryButton>
            <DefaultButton type="button" onClick={this.handleUpdate}>Update</DefaultButton>
          </form>
        </div>
      </>
    )
  }
}