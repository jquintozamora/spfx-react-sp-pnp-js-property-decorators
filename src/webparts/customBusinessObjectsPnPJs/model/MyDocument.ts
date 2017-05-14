import { Item, ODataEntity } from "sp-pnp-js";
import { select, expand, getSymbol } from "../utils/decorators";
import { SelectDecoratorsParser } from "../parser/decoratorsParser";


export class MyDocument extends Item {

  @select()
  public Title: string;

  @select()
  public FileLeafRef: string;

  @select("File/Length")
  @expand("File/Length")
  public Size: string;

  public CustomProps: string = "custom prop";

  // override get to enfore select and expand for our fields to always optimize
  public get(): Promise<MyDocument> {
    this
      ._setCustomQuery("select")
      ._setCustomQuery("expand");
    return super.get.call(this, ODataEntity(MyDocument), arguments[1] || {});
  }

  // override get to enfore select and expand for our fields to always optimize
  // used to solve MyDocument[] type checking
  public getAs<T>(): Promise<T> {
    this
      ._setCustomQuery("select")
      ._setCustomQuery("expand");
    return super.get.call(this, new SelectDecoratorsParser<MyDocument>(MyDocument), arguments[1] || {});
  }

  private _setCustomQuery(parameter: string) {
    // Symbol not supported on IE, maybe try with polyfill
    // const sym: symbol = Symbol.for(parameter);
    const sym = getSymbol(parameter);
    // get pre-saved select and expand props
    const arrayprops: { propName: string, queryName: string }[] = this[sym];
    let list =  arrayprops.map(i => i.queryName).join(",");
    // use apply and call to manipulate the request into the form we want
    // if another select isn't in place, let's default to only ever getting our fields.
    // implement method chain
    return this._query.getKeys().indexOf("$" + parameter) > -1
      ? this
      : this[parameter].call(this, list);
  }
}


