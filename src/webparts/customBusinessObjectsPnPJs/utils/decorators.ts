import { Logger, LogLevel } from "sp-pnp-js";


/*
 * Property Decorators
 */
export function select(selectName?: string): PropertyDecorator {
  // console.log("select(): evaluated");
  return function (target: Object, propertyKey: string): void {
    // console.log("select(): called");
    setMetadataMap(target, "select", propertyKey, selectName);
  };
}

export function expand(expandName: string): PropertyDecorator {
  // console.log("expand(): evaluated");
  return function (target: Object, propertyKey: string): void {
    // console.log("expand(): called");
    setMetadataMap(target, "expand", propertyKey, expandName);
  };
}


/*
 * decorators utils
 */
export function getSymbol(key: string): string {
  return "__" + key + "__";
}
function setMetadataMap(target: any, key: string, propName: string, queryName: string) {
  if (queryName === undefined) {
    queryName = propName;
  }
  // Symbol not supported on IE, maybe try with polyfill
  // const sym: symbol = Symbol.for(key);
  const sym = getSymbol(key);
  let currentValues: { propName: string, queryName: string }[] = target[sym];
  if (currentValues !== undefined) {
    currentValues = [...currentValues, { propName, queryName }];
  } else {
    currentValues = [].concat({ propName, queryName });
  }
  target[sym] = currentValues;
}


// export function annotation(tag: string): PropertyDecorator {
//   console.log("annotation(): evaluated");
//   return function (target: Object, propertyKey: string): void {
//     const sym: symbol = Symbol(tag);
//     target[sym] = true;
//     console.log("annotation(): called");
//     console.log(target);
//     console.log(propertyKey);
//   };
// }

// export function logProperty(target: Object, key: string): void {
//   // property value
//   var _val: any = target[key];

//   // property getter
//   var getter = function () {
//     console.log(`Get: ${key} => ${_val}`);
//     return _val;
//   };

//   // property setter
//   var setter = function (newVal) {
//     console.log(`Set: ${key} => ${newVal}`);
//     _val = newVal;
//   };

//   // delete property.
//   if (delete target[key]) {

//     // create new property with getter and setter
//     Object.defineProperty(target, key, {
//       get: getter,
//       set: setter,
//       enumerable: true,
//       configurable: true
//     });
//   }
// }
