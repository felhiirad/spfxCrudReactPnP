import * as React from 'react';
import styles from './PnPCrud.module.scss';
import { IPnPCrudProps } from './IPnPCrudProps';
import { escape } from '@microsoft/sp-lodash-subset';
import {IPnPCrudState } from './IPnPCrudState';
import {SPOperations} from '../../Services/SPOps';
import {Dropdown,Button}from 'office-ui-fabric-react'


export default class PnPCrud extends React.Component<
IPnPCrudProps,
IPnPCrudState,
 {}
 > 
 {
   private _spServices:SPOperations;
   public selectedListTitle:string;
  constructor(props:IPnPCrudProps){
    super(props)
    this._spServices=new SPOperations();
    this.state={listTitle:[],status:""};
  } 
  public componentDidMount(){
    this._spServices.getListeTitles().then((result)=>{
      this.setState({listTitle:result});
    });
  }
  //getSelectedListItem
  public getSelectedListTitle=(ev,data)=>{
    this.selectedListTitle=data.text;


  }
  public render(): React.ReactElement<IPnPCrudProps> {
    return (
      <div className={ styles.pnPCrud }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to Best Consulting it !</span>
              <p className={ styles.subTitle }>Demo : Sharepoint CRUD Operations using React PnPjs !!</p>
              
             
            </div>
            <div className={styles.myStyles} id="dv_ParentDiv">
              <Dropdown 
              className={styles.myDropDown} 
              options={this.state.listTitle}
              onChange={this.getSelectedListTitle}
               placeholder="*****select your list*****" >

              </Dropdown>
              <Button
               text="Create List Item" 
               onClick={()=>this._spServices
               .CreateListeItem(this.selectedListTitle)
               .then((result)=>{
                 this.setState({status:result})
                 })} >

               </Button>
               <Button
               text="update List Item" 
               onClick={()=>this._spServices
               .UpdateListeItem(this.selectedListTitle)
               .then((result)=>{
                 this.setState({status:result})
                 })} >

               </Button>
               <Button
               text="Delete List Item" 
               onClick={()=>this._spServices
               .DeleteListeItem(this.selectedListTitle)
               .then((result)=>{
                 this.setState({status:result})
                 })} >

               </Button>
              <div className={styles.myStatusBar}>{this.state.status}</div>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
