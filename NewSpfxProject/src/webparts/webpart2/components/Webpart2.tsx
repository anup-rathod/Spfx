import * as React from 'react';
// import styles from './Webpart2.module.scss';
import type { IWebpart2Props } from './IWebpart2Props';
import { escape } from '@microsoft/sp-lodash-subset';
import { PrincipalType, SPFx, spfi } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/items/get-all";
import type { IWebpart2State } from './IWebpart2State';
import { ComboBox, DefaultButton, IComboBox, TextField } from '@fluentui/react';
import { IWebpart2Add } from './IWebpart2Add';
import { PrimaryButton } from 'office-ui-fabric-react';
import { DetailsList, SelectionMode, IColumn } from '@fluentui/react/lib/DetailsList';
import { PeoplePicker } from '@pnp/spfx-controls-react/lib/PeoplePicker';



export default class Webpart2 extends React.Component<IWebpart2Props, IWebpart2State> {
  constructor(props: IWebpart2Props){
    super(props);

    const columns: IColumn[] = [
      {
        key: 'column1',
        name: 'Title',
        fieldName: 'Title',
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
      },
      {
        key: 'column2',
        name: 'Lookup Title',
        fieldName: 'Lookup',
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
        onRender: (item) => {
          return <span>{item.Lookup.Title}</span>;
        }
      },
      {
        key: 'column3',
        name: 'Lookup Cost',
        fieldName: 'Lookup',
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
        onRender: (item) => {
          return <span>{item.Lookup.Cost}</span>;
        }
      },
      {
        key: 'column4',
        name: 'Users Title',
        fieldName: 'selectedUsers',
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
        onRender: (item) => {
          return (<span>{item.selectedUsers ? item.selectedUsers.Title : ''}</span>)
        }
      },
      {
        key: 'column5',
        name: 'Users Email',
        fieldName: 'selectedUsers',
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
        onRender: (item) => {
          return (<span>{item.selectedUsers ? item.selectedUsers.EMail : ''}</span>)
        }
      },
      {
        key: 'column6',
        name: 'Edit',
        fieldName: 'Edit',
          minWidth: 100,
          maxWidth: 200,
          isResizable: true,
          onRender: (item) => {
          return (<DefaultButton text='Edit' onClick={()=> this.handleEdit(item.ID)}/>)
        }
      },
      {
        key: 'column7',
        name: 'Delete',
        fieldName: 'Delete',
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
        onRender: (item) => {
          return (<DefaultButton text='Delete' onClick={()=> this.handleDelete(item.ID)}/>)
        }
      },
    ];
    

    this.state = {
      ID: '',
      Title: '',
      Lookup: '',
      data: [],
      lookupOptions:[],
      columns: columns,
      selectedUsers: [],
    }
    
  }

  public async componentDidMount(): Promise<void> {
    try {
      await this.getAll();
      await this.getLookupOptions();

    } catch (error) {
      console.log("error in componentDidMount", error);
    }
  }

  public getAll = async () => {
    try {
      const sp: any = spfi().using(SPFx(this.props.context));
      const sp1 = sp.web.lists.getByTitle("Test");
      const items = await sp1.items.select("ID", "Title", "selectedUsers/Title", "selectedUsers/EMail", "Lookup/Title", "Lookup/Cost" ).expand( "selectedUsers","Lookup" ).getAll();

      console.log("Retrieved items:", items); 

      const UsersTitle = items.map((item: any) => item.selectedUsers !== undefined ? item.selectedUsers.Title : null);
      const UsersEMail = items.map((item: any) => item.selectedUsers !== undefined ? item.selectedUsers.EMail : null);

      const Users = items.map((item: any, index: number)=> {
        return {
          ...item, 
          UsersTitle: UsersTitle[index],
          UsersEMail : UsersEMail[index]
        }
      })
      
      console.log("Items Merge:", Users); 
      console.log("MyUser Title:", UsersTitle); 
      console.log("MyUser Email:", UsersEMail); 
      console.log("Items response:", items);

      this.setState({
        data: items,
      });
    } catch (error) {
      console.log("Error in getAll:", error); 
    }
  }

  public getLookupOptions = async () => {
    try {
      const sp: any = spfi().using(SPFx(this.props.context));
      
      const spList: any[] = await sp.web.lists.getByTitle("Project List").items.select('ID', 'Cost').getAll();
      let temp: any[] = [];  
      spList.forEach((value: any) => {
        temp.push({ key: value.ID, text: value.Cost });
      });

      this.setState({ lookupOptions: temp });

    } catch (error) {
      console.log("Error in getLookupOptions:", error);
    }
  }
  
  handleEdit = async (item : {Title : string , ID : number , Lookup : any, selectedUsers: any}) =>{
    const selectedMyLookup= item.Lookup?.Cost;
        this.setState({
          Title : item.Title,
          ID : item.ID,
          Lookup : item.Lookup,
          selectedUsers: item.selectedUsers,
        })
  }

  public handleSubmit = async (selectedKey: string, selectedPerson : any) : Promise<void> => {
    const { Title, Lookup, selectedUsers } = this.state as {
      Title: string,
      Lookup: string,
      selectedUsers: any,
    };
    // const sp: any = spfi().using(SPFx(this.props.context));
 
    const sp: any = spfi().using(SPFx(this.props.context));
    console.log("MAIN User", selectedPerson)
    console.log("Selected USER", selectedPerson.text);
    console.log("Selected USER", selectedPerson.secondaryText);

    const user = selectedPerson;
    if (user) {
      try {
        const list = await sp.web.lists.getByTitle("Test").items.add({
          'Title': Title,
          // 'Description': description,
          'LookupId': parseInt(selectedKey), 
          'selectedUsersId': user.id,
          // 'LookupColumnId': lookColumnValue, // Add this line
        });
 
        await this.getAll();
        this.setState({ Title: '', selectedUsers: '' });
        alert('Added Successfully');
      } catch (error) {
        console.error('Error adding item:', error);
        alert('Failed to add item. Please try again.');
      }
    } else {
      alert('Please fill all the fields');
    }
  }

  handleChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const name = event.target.name
    const value = event.target.value

    this.setState({
      [name]: value,
    } as Pick<IWebpart2Add, keyof IWebpart2Add>);
  }

  handleChangeLookup = (event: React.FormEvent<IComboBox>, option?: { key: string | number }) => {
    if (option) {
      this.setState({ Lookup: option.key as string });
    } else {
      this.setState({ Lookup: '' }); 
    }
  }

    handleDelete = async (Id: React.Key) => {
      try {
        const sp: any = spfi().using(SPFx(this.props.context));
        const list = sp.web.lists.getByTitle("Test");
        await list.items.getById(Id).delete();
        alert('List item deleted successfully');
        await this.getAll()
      } catch (error) {
        console.error('Error deleting item:', error);
        alert('Failed to delete item. Please try again.');
      }
    }

    handleUpdate = async (selectedPerson: []): Promise<void> => {
      const { ID, Title, data, Lookup, selectedUsers } = this.state;
      const sp = spfi().using(SPFx(this.props.context));
  
      const matchingIds = data.filter((item: { ID: React.Key }) => item.ID === ID).map((item: { Id: number }) => item.Id);
  
      const itemId = matchingIds.length > 0 ? matchingIds[0] : undefined;
  
      const user = selectedUsers.id;
      if (itemId) {
        try {
          const list = await sp.web.lists.getByTitle("Test").items.getById(itemId).update({
                    'Title': Title,
                    'Lookup': Lookup,
                    'selectedUsersId': user.id,
                  });
                  console.log(selectedUsers)
                  console.log(Lookup)
          this.getAll()
          this.setState({ 
            Title: '',
            Lookup: '',
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

    // handleEdit = async (selectedPersonId: any): Promise<void> => {
    //   const { ID, Title, data, lookupOptions, columns, selectedUsers } = this.state;
    //   const sp = spfi().using(SPFx(this.props.context));
  
    //   const matchingIds = data.filter((item: { ID: React.Key }) => item.ID === ID).map((item: { Id: number }) => item.Id);
    //   const itemId = matchingIds.length > 0 ? matchingIds[0] : undefined;
  
    //   const user = selectedPersonId;
    //   if (itemId) {
    //     try {
    //       const list = await sp.web.lists.getByTitle("Test").items.getById(itemId).update({
    //         'Title': Title,
    //         'PersonId': user.id,
    //       })
  
    //       await this.getAll();
    //       this.setState({
    //         ID: '',
    //         Title: '',
    //         Lookup: '',
    //         data: [],
    //         lookupOptions: [],
    //         columns: columns,
    //         selectedUsers: [],
    //       });
    //       alert('Updated Successfully');
    //     } catch (error) {
    //       console.error('Error updating item:', error);
    //       alert('Failed to update item. Please try again.');
    //     }
    //   } else {
    //     alert('Please select an item to edit');
    //   }
    // };
  

    private handlePeoplePickerChange = (selectedItems: any[]) => {
      if (selectedItems.length > 0) {
        this.setState({
          selectedUsers: selectedItems[0], // Assuming you want to select only one person
        });
      } else {
        this.setState({
          selectedUsers: null,
        });
      }
    };
    
    public render(): React.ReactElement<IWebpart2Props> {
      const {selectedUsers } = this.state;
      return (
        <>
          <div>
            {/* Render the DetailsList */}
            <div>
              <DetailsList
                items={this.state.data}
                columns={this.state.columns}
                selectionMode={SelectionMode.none}
              />
            </div>
    
            {/* Render form elements */}
            <div>
              <TextField label="Title" name="Title" onChange={this.handleChange} value={this.state.Title} />
              <br/>
              <ComboBox
                label="Lookup"
                options={this.state.lookupOptions}
                selectedKey={this.state.Lookup}
                onChange={this.handleChangeLookup}
                data-name="Lookup"
              />
              <br/>
            <PeoplePicker
              context={this.props.context}  
              titleText="Select People"
              personSelectionLimit={3}
              showtooltip={true}
              // Use defaultSelectedUsers to set initial selected users
              defaultSelectedUsers={[selectedUsers]}         
              onChange={this.handlePeoplePickerChange}
              ensureUser={true}
              resolveDelay={1000}
            /> <br/>
              <PrimaryButton text="Submit" onClick={() => this.handleSubmit(this.state.Lookup, this.state.selectedUsers)} />
              <DefaultButton type="button" onClick={() => { this.handleUpdate(this.state.selectedUsers) }}>Update</DefaultButton>
            </div>
          </div>
        </>
      );
    }
    
}