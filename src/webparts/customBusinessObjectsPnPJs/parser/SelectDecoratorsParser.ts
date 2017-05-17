import { ODataParserBase, QueryableConstructor, Util, Logger, LogLevel } from "sp-pnp-js";
import { getEntityUrl } from "sp-pnp-js/lib/sharepoint/odata";
import { getSymbol } from "../utils/decorators";

export class SelectDecoratorsArrayParser<T> extends ODataParserBase<T[]> {

    constructor(protected factory: QueryableConstructor<T>) {
        super();
    }

    public parse(r: Response): Promise<T[]> {
        return super.parse(r).then((d: any[]) => {
            return d.map(v => {
                const o = <T>new this.factory(getEntityUrl(v), null);
                return Util.extend(o, v);
            });
        });
    }
}

export class SelectDecoratorsParser<T> extends ODataParserBase<T> {

  constructor(protected factory: QueryableConstructor<T>) {
    super();
  }

  public parse(r: Response): Promise<T> {
    // we don't need to handleError inside as we are calling directly
    // to super.parse(r) and it's already handled there
    return super.parse(r).then(d => {
      const classDefaults: T = <T>new this.factory(getEntityUrl(d), null);
      const combinedWithResults: any = Util.extend(classDefaults, d);
      const sym: string = getSymbol("select");
      if ("length" in combinedWithResults) {
        return this._processCollection(combinedWithResults, sym);
      } else {
        return this._processSingle(combinedWithResults, sym);
      }
    });
  }

  // get only custom model properties with @select decorator and return single item
  private _processSingle(combinedWithResults: T, symbolKey: string): T {
    const arrayprops: { propName: string, queryName: string }[] = combinedWithResults[symbolKey];
    let newObj: T = {} as T;
    arrayprops.forEach((item) => {
      newObj[item.propName] = this._getDescendantProp(combinedWithResults, item.queryName);
    });
    return newObj;
  }

  // get only custom model properties with @select decorator and return item collection
  private _processCollection(combinedWithResults: T[], symbolKey: string): T[] {
    let newArray: T[] = [];
    const arrayprops: { propName: string, queryName: string }[] = combinedWithResults[symbolKey];
    for (let i: number = 0; i < combinedWithResults.length; i++) {
      const r: T = combinedWithResults[i];
      let newObj: T = {} as T;
      arrayprops.forEach((item) => {
        newObj[item.propName] = this._getDescendantProp(r, item.queryName);
      });
      newArray = newArray.concat(newObj);
    }
    return newArray;
  }

  private _getDescendantProp(obj, objectString: string) {
    var arr: string[] = objectString.split("/");
    if (arr.length > 1 && arr[0] !== "") {
      while (arr.length) {
        var name: string = arr.shift();
        if (name in obj) {
          obj = obj[name];
        } else {
          Logger.log({
            data: {
              name
            },
            level: LogLevel.Warning,
            message: "[getDescendantProp] - " + name + " property does not exists."
          });
          return null;
        }
      }
      return obj;
    }
    if (objectString !== undefined && objectString !== "") {
      return obj[objectString];
    }
    return null;
  }

}
