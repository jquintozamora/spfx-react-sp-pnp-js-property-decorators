import { ODataParserBase, QueryableConstructor, Util } from "sp-pnp-js";
import { getEntityUrl } from "sp-pnp-js/lib/sharepoint/odata";
import { getSymbol } from "../utils/decorators";

export class SelectDecoratorsParser<T> extends ODataParserBase<T> {

  constructor(protected factory: QueryableConstructor<T>) {
    super();
  }

  public parse(r: Response): Promise<T> {
    return super.parse(r).then(d => {

      const classDefaults = <T>new this.factory(getEntityUrl(d), null);
      const combinedWithResults = Util.extend(classDefaults, d);
      const sym = getSymbol("select");


      // TODO: combinedWithResults could be an array
      if ("length" in combinedWithResults) {
        let newArray = [];
        const arrayprops: { propName: string, queryName: string }[] = combinedWithResults[sym];
        for (let i = 0; i < combinedWithResults.length; i++) {
          const r = combinedWithResults[i];
          let newObj = {};
          arrayprops.forEach((item) => {
            newObj[item.propName] = this._getDescendantProp(r, item.queryName);
          });
          newArray = newArray.concat(newObj);
        }
        return newArray;
      } else {
        // only returns the properties with @select decorator
        const arrayprops: { propName: string, queryName: string }[] = combinedWithResults[sym];
        let newObj = {};
        arrayprops.forEach((item) => {
          newObj[item.propName] = this._getDescendantProp(combinedWithResults, item.queryName);
        });
        return newObj;
      }

    });
  }

  private _getDescendantProp(obj, objectString) {
    var arr = objectString.split("/");
    if (arr.length > 1 && arr[0] !== "") {
      while (arr.length) {
        var name = arr.shift();
        if (name in obj) {
          obj = obj[name];
        } else {
          console.warn('[getDescendantProp] - ' + name + ' property does not exists.');
          return undefined;
        }
      }
      return obj;
    }
    if (objectString !== undefined && objectString !== "") {
      return obj[objectString];
    }
    return null;
  }
  // public parse(response: Response): Promise<T> {
  //   return new Promise((resolve, reject) => {
  //     if (this.handleError(response, reject)) {
  //       debugger;
  //       resolve(response);
  //     }
  //   });
  // }
}
