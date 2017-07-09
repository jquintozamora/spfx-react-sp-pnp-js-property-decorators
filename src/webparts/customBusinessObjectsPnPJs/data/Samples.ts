const libraryName = "Documents";
import pnp from "sp-pnp-js";

// import models
import { MyDocument } from "../model/MyDocument";
import { MyDocumentCollection } from "../model/MyDocumentCollection";
// initially we import our custom model which extends from Item class from PnP JS Core
import { MyItem } from "../model/MyItem";

// import custom parsers
import { SelectDecoratorsParser, SelectDecoratorsArrayParser } from "../parser/SelectDecoratorsParsers";


export interface ICodeSample {
  tabTitle: string;
  title: string;
  code: () => void;
  codeText: () => string;
}

// TODO: Investigate webpack custom plugin to get the function text automatically


export const samples: ICodeSample[] = [
  {
    title: "PnP JS Core WITHOUT custom objects",
    tabTitle: "No.Custom.Objects",
    code: async () => {
      const plainItemAsAny: any = await pnp.sp
        .web
        .lists
        .getByTitle("PnPJSSample")
        .items
        .getById(1)
        .select("ID", "Title", "Category", "Quantity")
        .get();
      console.log(plainItemAsAny);
    },
    codeText: () => {
      return `const plainItemAsAny: any = await pnp.sp
  .web
  .lists
  .getByTitle("PnPJSSample")
  .items
  .getById(1)
  .select("ID", "Title", "Category", "Quantity")
  .get();
console.log(plainItemAsAny);`;
    }
  },
  {
    title: "PnP JS Core WITH custom objects",
    tabTitle: "Custom.Objects",
    code: async () => {
      const itemCustomObject: MyItem = await pnp.sp
        .web
        .lists
        .getByTitle("PnPJSSample")
        .items
        .getById(1)
        // we don't need the select
        // .select("ID", "Title", "Category", "Quantity")
        .as(MyItem)
        .get();
      console.log(itemCustomObject);
    },
    codeText: () => {
      return `const itemCustomObject: MyItem = await pnp.sp
  .web
  .lists
  .getByTitle("PnPJSSample")
  .items
  .getById(1)
  // we don't need the select
  // .select("ID", "Title", "Category", "Quantity")
  .as(MyItem)
  .get();
console.log(itemCustomObject);`;
    }
  },
  {
    title: "One document selecting all properties",
    tabTitle: "Single.AllProps",
    code: async () => {
      const myDocument: any = await pnp.sp
        .web
        .lists
        .getByTitle(libraryName)
        .items
        .getById(1)
        .get();
      // query all item's properties
      console.log(myDocument);
    },
    codeText: () => {
      return `const myDocument: any = await pnp.sp
  .web
  .lists
  .getByTitle(libraryName)
  .items
  .getById(1)
  .get();
// query all item's properties
console.log(myDocument);`;
    }
  },
  {
    title: "One document using select, expand and get()",
    tabTitle: "Single.SelectExpand",
    code: async () => {
      const singleSelectExpand: any = await pnp.sp
        .web
        .lists
        .getByTitle(libraryName)
        .items
        .getById(1)
        .select("Title", "FileLeafRef", "File/Length")
        .expand("File/Length")
        .get();
      // query only selected properties, but ideally should
      // get the props from our custom object
      console.log(singleSelectExpand);
    },
    codeText: () => {
      return `const singleSelectExpand: any = await pnp.sp
  .web
  .lists
  .getByTitle(libraryName)
  .items
  .getById(1)
  .select("Title", "FileLeafRef", "File/Length")
  .expand("File/Length")
  .get();
// query only selected properties, but ideally should
// get the props from our custom object
console.log(singleSelectExpand);`;
    }
  },
  {
    title: "One document using as(MyDocument) and get() with Default Parser",
    tabTitle: "Single.As",
    code: async () => {
      const singleAs: MyDocument = await pnp.sp
        .web
        .lists
        .getByTitle(libraryName)
        .items
        .getById(1)
        .as(MyDocument)
        .get();
      // query only selected properties in our Custom Object
      // only those with @select and @expand decorators
      console.log(singleAs);
    },
    codeText: () => {
      return `const singleAs: MyDocument = await pnp.sp
  .web
  .lists
  .getByTitle(libraryName)
  .items
  .getById(1)
  .as(MyDocument)
  .get();
// query only selected properties in our Custom Object
// only those with @select and @expand decorators
console.log(singleAs);`;
    }
  },
  {
    title: "One document using as(MyDocument) and getAs<MyDocument>()",
    tabTitle: "Single.As.Parser",
    code: async () => {
      const singleAsParser: MyDocument = await pnp.sp
        .web
        .lists
        .getByTitle(libraryName)
        .items
        .getById(1)
        .as(MyDocument)
        // it's using getAs from MyDocument which has SelectDecoratorsParser
        .getAs<MyDocument>();
      // query only selected properties, using our Custom Model properties
      // but only those that have the proper @select and @expand decorators
      console.log(singleAsParser);
    },
    codeText: () => {
      return `const singleAsParser: MyDocument = await pnp.sp
  .web
  .lists
  .getByTitle(libraryName)
  .items
  .getById(1)
  .as(MyDocument)
  // getAs from MyDocument has SelectDecoratorsParser
  .getAs<MyDocument>();
// query only selected properties, using @select
// and map the properties using custom parser
console.log(singleAsParser);`;
    }
  },
  {
    title: "Document Collection selecting all properties",
    tabTitle: "Collection.AllProps",
    code: async () => {
      const myDocumentCollection: any = await pnp.sp
        .web
        .lists
        .getByTitle(libraryName)
        .items
        .get();
      console.log(myDocumentCollection);
    },
    codeText: () => {
      return `const myDocumentCollection: any = await pnp.sp
  .web
  .lists
  .getByTitle(libraryName)
  .items
  .get();
console.log(myDocumentCollection);`;
    }
  },
  {
    title: "Document Collection using as(MyDocumentCollection) and get()",
    tabTitle: "Collection.As",
    code: async () => {
      const collectionAs: MyDocument[] = await pnp.sp
        .web
        .lists
        .getByTitle(libraryName)
        .items
        // using as("Model") overrides select and expand queries
        .as(MyDocumentCollection)
        .get();
      // query only selected properties in our Custom Object
      // only those with @select and @expand decorators
      console.log(collectionAs);
    },
    codeText: () => {
      return `const collectionAs: MyDocument[] = await pnp.sp
  .web
  .lists
  .getByTitle(libraryName)
  .items
  // using as("Model") overrides select and expand queries
  .as(MyDocumentCollection)
  .get();
// query only selected properties in our Custom Object
// only those with @select and @expand decorators
console.log(collectionAs);`;
    }
  },
  {
    title: "Document Collection using as(MyDocumentCollection) and get() with Custom Array Parser",
    tabTitle: "Collection.As.ArrayParser",
    code: async () => {
      const collectionAsArrayParser: any[] = await pnp.sp
        .web
        .lists
        .getByTitle(libraryName)
        .items
        // using as("Model") overrides select and expand queries
        .as(MyDocumentCollection)
        .skip(1)
        // this renderer mix the properties and do the match between the props names and the selected if they have /
        .get(new SelectDecoratorsArrayParser<MyDocument>(MyDocument));
      // query only selected properties, using our Custom Model properties
      // but only those that have the proper @select and @expand decorators
      console.log(collectionAsArrayParser);
    },
    codeText: () => {
      return `const collectionAsArrayParser: any[] = await pnp.sp
  .web
  .lists
  .getByTitle(libraryName)
  .items
  // using as("Model") overrides select and expand queries
  .as(MyDocumentCollection)
  .skip(1)
  // we stil can use method chain
  .get(new SelectDecoratorsArrayParser<MyDocument>(MyDocument));
console.log(collectionAsArrayParser);`;
    }
  },
  {
    title: "Document Collection using as(MyDocumentCollection) and get() with Custom Array Parser",
    tabTitle: "Collection.As.Parser.JustOurModel",
    code: async () => {
      const myDocumentsAsParserJustSelect: any[] = await pnp.sp
        .web
        .lists
        .getByTitle(libraryName)
        .items
        // using as("Model") overrides select and expand queries
        .as(MyDocumentCollection)
        // we can't use method chain afterwards...
        .get(new SelectDecoratorsArrayParser<MyDocument>(MyDocument, true));
      // query only selected properties, using our Custom Model properties
      // but only those that have the proper @select and @expand decorators
      console.log(myDocumentsAsParserJustSelect);
    },
    codeText: () => {
      return `const myDocumentsAsParserJustSelect: any[] = await pnp.sp
  .web
  .lists
  .getByTitle(libraryName)
  .items
  .as(MyDocumentCollection)
  // we can't use method chain afterwards...
  .get(new SelectDecoratorsArrayParser<MyDocument>(MyDocument, true));
console.log(myDocumentsAsParserJustSelect);`;
    }
  }


];
