import * as React from 'react';
import { ITask1Props } from './ITask1Props';
import { ITask1State } from './ITask1State';
import { spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import "@pnp/sp/webs";
import "@pnp/sp/fields";

export default class Task1 extends React.Component<ITask1Props, ITask1State> {
  private sp: ReturnType<typeof spfi>;

  constructor(props: ITask1Props) {
    super(props);

    this.sp = spfi().using(SPFx(this.props.context));

    this.state = {
      Name: "",
      Amount: 0,
      Country: "",
      CompanyName: "",
      data: []
    };
  }

  public async componentDidMount() {
    await this.getData();
  }

  public getData = async () => {
    try {
      const response = await fetch("https://localhost:7231/api/Employeeapi");
      if (!response.ok) {
        throw new Error('Network response was not ok');
      }
      const data = await response.json();
      console.log(data);

      this.setState({ data: data });
    } catch (error) {
      console.error("Error fetching data:", error);
    }
  }

  public render(): React.ReactElement<ITask1Props> {
    const { data } = this.state;

    return (
      <div>
        {data && data.map((item: any, index: number) => (
          <ul key={index}>
            <li>Name: {item.name}</li>
            <li>Company: {item.companyName}</li>
            <li>Country: {item.countries}</li>
            <li>Amount: {item.amount}</li>
          </ul>
        ))}
      </div>
    );
  }
}
