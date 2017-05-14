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
      myDocuments: null,
      errors: []
    };

    // normally we don't need to bind the functions as we use arrow functions and do automatically the bing
    // http://bit.ly/reactArrowFunction
    // but using Async function we can't convert it into arrow function, so we do the binding here
    this._loadPnPJsLibrary.bind(this);

  }

  public render(): React.ReactElement<ICustomBusinessObjectsPnPJsProps> {
    return (
      <div className={styles.helloWorld}>
        <div className={styles.container}>
          <div className={"ms-Grid-row ms-bgColor-themeDark ms-fontColor-white " + styles.row}>
            <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <span className="ms-font-xl ms-fontColor-white">Welcome to SharePoint!</span>
              <p className="ms-font-l ms-fontColor-white">Customize SharePoint experiences using Web Parts.</p>
              <p className="ms-font-l ms-fontColor-white">{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={styles.button}>
                <span className={styles.label}>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
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
        // Using as("Model") overrides select and expand queries
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
        // Using as("Model") overrides select and expand queries
        // That´s where the MAGIC happends as even if we are using
        // ITEMS (item collection) it will use the proper query
        .as(MyDocument)
        // Using MyDocument[] match the type checking for the returned object
        // and avoid javaScript error
        .getAs<MyDocument[]>();
      // query only selected properties, using our Custom Model properties
      // but only those that have the proper @select and @expand decorators
      console.log(myDocumentsWithCustomObject);

      debugger;

      // Set our Component´s State
      // this.setState({ ...this.state, myDocuments });

    } catch (error) {
      // set a new state conserving the previous state + the new error
      console.error(error);
      this.setState({
        ...this.state,
        errors: [...this.state.errors, "Error getting ItemCount for " + libraryName + ". Error: " + error]
      });
    }
  }

}
