import * as React from 'react';
import type { IWebpart4Props } from './IWebpart4Props';
import { escape } from '@microsoft/sp-lodash-subset';
import { IWebpart4State } from './IWebpart4State';
import { SPFx, spfi } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/items/get-all";
import { Dropdown, TextField } from 'office-ui-fabric-react';
import { IDropdownOption } from '@fluentui/react';

export default class Webpart4 extends React.Component<IWebpart4Props, IWebpart4State> {

  constructor(props: IWebpart4Props) {
    super(props);
    this.state = {
      ID: '',
      Title: '',
      Description: '',
      Address: '',
      data: [],
      Lookup: '',
      LookupOptions: [],
    }
  }

  componentDidMount(): void {
    this.getList();
    this.fetchLookupOptions();
  }

  getList = async() => {
    const sp = spfi().using(SPFx(this.props.context))
    const list= sp.web.lists.getByTitle("Test")
    const items = await list.items.select("ID", "Title", "Description", "Address", "Lookup/Title", "Lookup/Cost").expand("Lookup").getAll()
    console.log("mylist", items)
    this.setState({
      data: items
    })
  }

  fetchLookupOptions = async () => {
    const sp = spfi().using(SPFx(this.props.context))
    const lookupList = sp.web.lists.getByTitle("Project List")
    const items = await lookupList.items.select("Cost").getAll()
    const options : IDropdownOption[] = items.map(item => {
      return {
        key: item.Cost,
        text: item.Cost,
      } 
    })

    this.setState({
      LookupOptions: options
    })
  }

  handleLookupChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number): void => {
    const value = option ? (option.key as string) : ''; 

    this.setState({
      Lookup: value
    });
  }

  handleChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const name = event.target.name;
    const value = event.target.value;

    this.setState({
      [name]: value
    } as unknown as Pick<IWebpart4State, keyof IWebpart4State>);
  }

  public render(): React.ReactElement<IWebpart4Props> {
    return (
      <>
        <div>
        {
          this.state.data.map((item: any, index: number) => (
            <div key={index}>
              <ul>
                <li>{item.Title}</li> 
                <li>{item.Address}</li>
                <li>{item.Description}</li>
                <li>{item.Lookup && item.Lookup.Title}</li>
                <li>{item.Lookup && item.Lookup.Cost}</li>
              </ul>
            </div>
          ))
        }
      </div>
      <form>
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
              selectedKey={this.state.Lookup}
              onChange={this.handleLookupChange}
              options={[
                { key: '', text: 'Select Choice' },
                ...this.state.LookupOptions
              ]}
            />
      </form>
      </>
    );
  }
}
