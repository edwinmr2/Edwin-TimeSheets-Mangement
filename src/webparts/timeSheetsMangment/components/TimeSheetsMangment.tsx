import * as React from 'react';
import styles from './TimeSheetsMangment.module.scss';
import { ITimeSheetsMangmentProps } from './ITimeSheetsMangmentProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { sp, Web, IWeb } from "@pnp/sp/presets/all";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';

export interface ISPListTimeSheetsItem{
    Items: any;
    ID:number;
    Title:string;
    Description:string;
    CreatedBy:string;
    CreatedOn:Date;
    Category:string;
    Hours:number;
    Status:string;
    HTML: any;
}

export default class TimeSheetsMangment extends React.Component<ITimeSheetsMangmentProps, ISPListTimeSheetsItem> {
    constructor(props) {
        super(props);
        this.state = {
            Items: [],
            ID:0,
            Title: "",
            Description: "",
            CreatedBy: "",
            CreatedOn: null,
            Category: "Billable",
            Hours: 0,
            Status: "",
          HTML: []
    
        };
      }
    
      public async componentDidMount() {
        await this.fetchData();
      }
    
      public async fetchData() {
       
        let web = Web(this.props.webURL);
        //Get Items Created by current User
        const items: any[] = await web.lists.getByTitle("TimeSheets").items.filter("Author/EMail eq '${encodeURIComponent(this.context.pageContext.user.email)}'").get();
        console.log(items);
        this.setState({ Items: items });
        let html = await this.getHTML(items);
        this.setState({ HTML: html });
      }
     
    //Method to display current time sheet records
      public async getHTML(items) {
        var tabledata = <table className={styles.table}>
          <thead>
            <tr>
                <th>#</th>
                <th>Title</th>
                <th>Description</th>
                <th>Created By</th>
                <th>Created Date</th>
                <th>Category</th>
                <th>Hours</th>
                <th>Status</th>
            </tr>
          </thead>
          <tbody>
            {items && items.map((item, i) => {
              return [
                <tr>
                  <td>{item.Title}</td>
                  <td>{item.Description}</td>
                  <td>{item.CreatedOn}</td>
                  <td>{item.CreatedBy}</td>
                  <td>{item.Category}</td>
                  <td>{item.Hours}</td>
                  <td>{item.Status}</td>
                </tr>
              ];
            })}
          </tbody>
    
        </table>;
        return await tabledata;
      }
    
      //set default state of the form
      public onchange(value, stateValue) {
        let state = {};
        state[stateValue] = value;
        this.setState(state);
      }

      private async SaveToSPList(){

        //check if user has created more hours that required in a day
        
        let web = Web(this.props.webURL);
        const items: any[] = await web.lists.getByTitle("TimeSheets").items.filter("Author/EMail eq '${encodeURIComponent(this.context.pageContext.user.email)}'").get();
        var currentHoursForTheDay = 0;
        {items && items.map((item, i) => {

          if(item.CreatedOn == new Date()){
            //add number of hours already recorded for the day
            currentHoursForTheDay = currentHoursForTheDay + item.Hours;
          }

        })}

        currentHoursForTheDay = currentHoursForTheDay + this.state.Hours;
        if(currentHoursForTheDay < 8 ){
          this.SaveData("Approved");
          alert("Time Sheet Record Created Successfully");
        }
        else{
          this.SaveData("Pending Approval");
          alert("Time Sheet Record Created Successfully. Please note that you are captured more hours today, your timesheet record has been sent for approval.");

        }
        this.setState({Title:"",Description:"",CreatedOn:null,CreatedBy:"",Category:"",Hours:0,Status:""});
        this.fetchData();
      }
      private async SaveData(status) {
        let web = Web(this.props.webURL);
        await web.lists.getByTitle("TimeSheets").items.add({
    
            Title:this.state.Title,
            Description: new Date(this.state.Description),
            CreatedOn: this.state.CreatedOn,
            CreatedBy:this.state.CreatedBy,
            Category: new Date(this.state.Category),
            Hours: this.state.Hours,
            Status: status,
    
        }).then(i => {
          console.log(i);
        });
       
        
      }
     
     
      public render(): React.ReactElement<ITimeSheetsMangmentProps> {
        return (
          
          <div>
             <h4><strong>Submit Time Sheet Record</strong></h4>
            {this.state.HTML}
            <div>
              <form>
              <div>
                  <Label>Title</Label>
                  <TextField value={this.state.Title}  onChanged={(value) => this.onchange(value, "Title")} />
                </div>
                <div>
                  <Label>Description</Label>
                  <TextField value={this.state.Description} multiline onChanged={(value) => this.onchange(value, "Description")} />
                </div>
                <div>
                  <Label>Category</Label>
                    <select id="sltCategory" value={this.state.Category}  onChange={(value) => this.onchange(value, "Category")} >
                        <option selected value="1">Billable </option>
                        <option value="2">Non-Billable</option>
                        <option value="3">Upskilling </option>
                        <option value="4">Meeting </option>
                    </select>
                </div>
                <div>
                  <Label>Hours</Label>
                  <input  type="number" value={this.state.Hours}  onChange={(value) => this.onchange(value, "Hours")} />
                </div>
              </form>
            </div>
           
            <div className={styles.btngroup}>
              <div><PrimaryButton text="Create" onClick={() => this.SaveToSPList()}/></div>
            </div>
          </div>
        );
      }
      
}