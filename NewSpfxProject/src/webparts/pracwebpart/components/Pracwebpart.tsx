import * as React from 'react';
import styles from './Pracwebpart.module.scss';
import type { IPracwebpartProps } from './IPracwebpartProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IPracwebpartState } from './IPracwebpartState';
import { SPFx, spfi } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/items/get-all";

export default class Pracwebpart extends React.Component<IPracwebpartProps, IPracwebpartState> {
  constructor(props: IPracwebpartProps){
    super(props);
    this.state= {
      Title: '',
      Address: '',
      Description: '',
      data: [],
    }
  }

  componentDidMount(): void {
    console.log("Component mounted"); 
    this.getListData();
  }
  
  public getListData = async () => {
    try {
      const sp = spfi().using(SPFx(this.props.context));
      const allItems = await sp.web.lists.getByTitle("Test").items.getAll();
      console.log(allItems);
      this.setState({
        data : allItems,
      });
    } catch (error) {
      console.log(error);
    }
  }
  
  public render(): React.ReactElement<IPracwebpartProps> {
    return (
      <div>
        {
          this.state.data.map((item: any, index: number) => (
            <div key={index}>
              <ul>
                <li>{item.Title}</li> 
                <li>{item.Address}</li>
                <li>{item.Description}</li>
              </ul>
            </div>
          ))
        }
      </div>
    );
  }
}
