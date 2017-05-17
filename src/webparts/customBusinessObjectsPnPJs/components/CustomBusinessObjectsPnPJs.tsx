import * as React from 'react';
import styles from './CustomBusinessObjectsPnPJs.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';

// import model
import { MyDocument } from "../model/MyDocument";

// import pnp and pnp logging system
import pnp from "sp-pnp-js";

import { ICustomBusinessObjectsPnPJsProps } from './ICustomBusinessObjectsPnPJsProps';
import { ICustomBusinessObjectsPnPJsState } from './ICustomBusinessObjectsPnPJsState';

export default class CustomBusinessObjectsPnPJs extends React.Component<ICustomBusinessObjectsPnPJsProps, ICustomBusinessObjectsPnPJsState> {

  constructor(props: ICustomBusinessObjectsPnPJsProps) {
    super(props);
    // set initial state
    this.state = {
      myDocuments: [],
      errors: []
    };

    // normally we don't need to bind the functions as we use arrow functions and do automatically the bing
    // http://bit.ly/reactArrowFunction
    // but using Async function we can't convert it into arrow function, so we do the binding here
    this._loadPnPJsLibrary.bind(this);

  }

  public render(): React.ReactElement<ICustomBusinessObjectsPnPJsProps> {
    return (
      <div className={styles.container}>
        <div className={"ms-Grid-row ms-bgColor-themeDark ms-fontColor-white " + styles.row}>
        <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
          <span className="ms-font-xl ms-fontColor-white">Welcome to SharePoint Async Await SP PnP JS Demo!</span>
          <div>
            {this._gerErrors()}
          </div>
          <p className="ms-font-l ms-fontColor-white">List of documents:</p>
          <div>
            <div className={styles.row}>
              <div className={styles.left}>Name</div>
              <div className={styles.right}>Size (KB)</div>
              <div className={styles.clear + " " + styles.header}></div>
            </div>
            {this.state.myDocuments.map((item) => {
              return (
                <div className={styles.row}>
                  <div className={styles.left}>{item.Name}</div>
                  <div className={styles.right}>{(item.Size / 1024).toFixed(2)}</div>
                  <div className={styles.clear}></div>
                </div>
              );
            })}
          </div>
        </div>
      </div>
      </div >
    );
  }

  public componentDidMount(): void {
    const libraryName: string = "Documents";
    console.log("libraryName: " + libraryName);
    this._loadPnPJsLibrary(libraryName);
  }

  private async _loadPnPJsLibrary(libraryName: string): Promise<void> {
    console.log("loadPnPJsLibrary");
    try {

      const myDocument = await pnp.sp
        .web
        .lists
        .getByTitle(libraryName)
        .items
        .getById(1)
        .get();
      // query all item's properties
      console.log(myDocument);


      const myDocumentGetAs = await pnp.sp
        .web
        .lists
        .getByTitle(libraryName)
        .items
        .getById(1)
        .getAs<MyDocument>();
      // query all item's properties, exactly the same result as before
      console.log(myDocumentGetAs);


      const myDocumentGetAsWithSelectExpand = await pnp.sp
        .web
        .lists
        .getByTitle(libraryName)
        .items
        .getById(1)
        .select("Title", "FileLeafRef", "File/Length")
        .expand("File/Length")
        .getAs<MyDocument>();
      // query only selected properties, but ideally should
      // get the props from our custom object
      console.log(myDocumentGetAsWithSelectExpand);


      const myDocumentWithCustomObject: MyDocument = await pnp.sp
        .web
        .lists
        .getByTitle(libraryName)
        .items
        .getById(1)
        // using as("Model") overrides select and expand queries
        .as(MyDocument)
        .getAs<MyDocument>();
      // query only selected properties, using our Custom Model properties
      // but only those that have the proper @select and @expand decorators
      console.log(myDocumentWithCustomObject);


      const myDocumentsWithCustomObject: MyDocument[] = await pnp.sp
        .web
        .lists
        .getByTitle(libraryName)
        .items
        // using as("Model") overrides select and expand queries
        // that´s where the MAGIC happends as even if we are using
        // items (item collection) it will use the proper query
        .as(MyDocument)
        // using MyDocument[] match the type checking for the returned object
        // and avoid javaScript error
        .getAs<MyDocument[]>();
      // query only selected properties, using our Custom Model properties
      // but only those that have the proper @select and @expand decorators
      console.log(myDocumentsWithCustomObject);


      // set our Component´s State
      this.setState({ ...this.state, myDocuments: myDocumentsWithCustomObject });

    } catch (error) {
      // set a new state conserving the previous state + the new error
      console.error(error);
      this.setState({
        ...this.state,
        errors: [...this.state.errors, "Error getting ItemCount for " + libraryName + ". Error: " + error]
      });
    }
  }

  private _gerErrors() {
    return this.state.errors.length > 0
      ?
      <div style={{ color: "orangered" }} >
        <div>Errors:</div>
        {
          this.state.errors.map((item) => {
            return (<div>{JSON.stringify(item)}</div>);
          })
        }
      </div>
      : null;
  }

}
