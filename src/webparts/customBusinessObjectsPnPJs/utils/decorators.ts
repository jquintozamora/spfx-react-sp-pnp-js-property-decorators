import { Logger, LogLevel } from "sp-pnp-js";

// symbol emulation as it's not supported on IE
// consider using polyfill as well
import { getSymbol } from "./symbol";

/*
 * Property Decorators
 */
export function select(selectName?: string): PropertyDecorator {
  return function (target: Object, propertyKey: string): void {
    setMetadata(target, "select", propertyKey, selectName);
  };
}
export function expand(expandName: string): PropertyDecorator {
  return function (target: Object, propertyKey: string): void {
    setMetadata(target, "expand", propertyKey, expandName);
  };
}

/*
 * decorators utils
 */
function setMetadata(target: any, key: string, propName: string, queryName: string): void {
  if (queryName === undefined
    || queryName === null
    || queryName === "") {
    queryName = propName;
  }
  const sym: string = getSymbol(key);
  // instead of using Map object, we use an array of objects, as Map is not well supported
  // still by all the browsers, consider using Map compiling TypeScript to ES6 and using Babel to transpile to ES5
  let currentValues: { propName: string, queryName: string }[] = target[sym];
  if (currentValues !== undefined) {
    currentValues = [...currentValues, { propName, queryName }];
  } else {
    currentValues = [].concat({ propName, queryName });
  }
  // property Decorators will store the metadata in its instance ( as a class property)
  // ideally having a symbol as a key, but symbol are not still supported on all the browsers
  // and they will require polyfill, as a sample, I will not use symbols, but please, consider it
  target[sym] = currentValues;
  Logger.log({
    data: {
      propertyKey: propName,
      queryName,
      key,
      target,
    },
    level: LogLevel.Verbose,
    message: "set metadata for property decorator"
  });
}
