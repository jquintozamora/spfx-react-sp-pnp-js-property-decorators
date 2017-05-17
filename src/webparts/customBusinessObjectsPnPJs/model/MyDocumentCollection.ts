import { Items, ODataEntityArray, ODataParser, FetchOptions } from "sp-pnp-js";
import { select, expand, getSymbol } from "../utils/decorators";
import { SelectDecoratorsParser } from "../parser/SelectDecoratorsParser";


export class MyDocumentCollection extends Items {

  @select()
  public Title: string;

  @select("FileLeafRef")
  public Name: string;

  @select("File/Length")
  @expand("File/Length")
  public Size: number;


  // public CustomProps: string = "Custom Prop to pass";

  // override get to enfore select and expand for our fields to always optimize
  public get(parser?: ODataParser<any>, getOptions?: FetchOptions): Promise<any> {
    // public get(): Promise<MyDocument> {
    this
      ._setCustomQueryFromDecorator("select")
      ._setCustomQueryFromDecorator("expand");
    if (parser === undefined) {
      // parser = ODataEntityArray(MyDocuments);
      parser = new SelectDecoratorsParser<MyDocumentCollection>(MyDocumentCollection);
    }
    return super.get.call(this, parser, getOptions);
  }


  private _setCustomQueryFromDecorator(parameter: string): MyDocumentCollection {
    const sym: string = getSymbol(parameter);
    // get pre-saved select and expand props from decorators
    const arrayprops: { propName: string, queryName: string }[] = this[sym];
    let list: string = arrayprops.map(i => i.queryName).join(",");
    // use apply and call to manipulate the request into the form we want
    // if another select isn't in place, let's default to only ever getting our fields.
    // implement method chain
    return this._query.getKeys().indexOf("$" + parameter) > -1
      ? this
      : this[parameter].call(this, list);
  }
}


