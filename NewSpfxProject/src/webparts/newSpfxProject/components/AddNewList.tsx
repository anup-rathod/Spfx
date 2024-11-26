import * as React from 'react'
import  { Component } from 'react'
import { INewSpfxProjectProps } from './INewSpfxProjectProps'
import { IAddNewList } from './IAddNewList'
import { SPFx, spfi } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";

export default class AddNewList extends Component<INewSpfxProjectProps, IAddNewList> {
    constructor(props: INewSpfxProjectProps){
        super(props);
        this.state = {
            Title: '',
            Description: '',
            Address: '',
            Choice: '',
            ChoiceOptions: []
        }
    }

    public async componentDidMount(): Promise<void> {
        await this.fetchChoiceOptions(); 
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
        const { Title, Description, Address, Choice} = this.state as {
            Title: string;
            Description: string;
            Address:string;
            Choice: string;
        }
        const sp =  spfi().using(SPFx(this.props.context));
        const AddedList = await sp.web.lists.getByTitle("Test").items.add({
            'Title': Title,
            'Description': Description,
            'Address': Address,
            'Choice': Choice,
        })
    }

    handleSelectChange = (event: React.ChangeEvent<HTMLSelectElement>) => {

       
        const value = event.target.value;
        this.setState({
            Choice : value
        })

    }

  render() {
    return (
      <div>
        <form onSubmit={this.handleSubmit}>
            <label htmlFor="Title">Title: 
                <input type="text"  name='Title' value={this.state.Title} onChange={this.handleChange}></input>
            </label><br/>
            <label htmlFor="Description">Description: 
                <input type="text"  name='Description' value={this.state.Description} onChange={this.handleChange}/>
            </label><br/>
            <label htmlFor="Address">Address: 
                <input type="text"  name='Address' value={this.state.Address} onChange={this.handleChange}/>
            </label><br/>
            <label htmlFor="Choice" >Choice: 
                <select value={this.state.Choice} onChange={this.handleSelectChange}>
                <option value="">Select Choice</option>
                {this.state.ChoiceOptions.map((option: string) => {
                    <option key={option} value={option}>                                         
                      {option}
                   </option>
                })}
                </select>
            </label>
           <button>Submit</button>
        </form>
      </div>
    )
  }
}


