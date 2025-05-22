// export enum CRUD {
//   Create,
//   Read,
//   Update,
//   Delete,
// }

// export enum ListNames {
//   ansList = "AnswersList",
//   quesList = "Questions",
//   rootList = "QuestionnaireRootList",
// }
// export const listNameArr = [
//   ListNames.ansList,
//   ListNames.quesList,
//   ListNames.rootList,
// ];

// /**
//  * Checks if value is empty. A value is considered empty unless it’s an arguments object, array, string, or
//  * jQuery-like collection with a length greater than 0 or an object with own enumerable properties.
//  *
//  * @param args The array of values to inspect.
//  * @return Returns true if any value is empty, else false.
//  */
// export const itemIsEmpty = (...args: any[]): boolean => {
//   for (const val of args) {
//     if (
//       val === "" ||
//       val === null ||
//       val === "null" || // helpful when an item is to be checked after JSON.stringify
//       val === "undefined" || // helpful when an item is to be checked after JSON.stringify
//       val === undefined ||
//       (typeof val === "object" && !Object.keys(val).length)
//     ) {
//       return true; // If any argument is empty, return true
//     }
//   }
//   return false; // If none of the arguments are empty, return false
// };

// const crudSpList = (
//   siteAbsoluteUrl: string,
//   apiGetType: string,
//   listname: string,
//   queryApiArgs?: string, // like filter, select etc, to append to the queryURL
//   addToHeaders?: RequestInit, // obj to add to the headers
//   method = CRUD.Read,
//   requestDigest?: string, // requied in create, edit and delete
//   id?: number, // required in edit and delete
//   requestdata?: any // required in edit and delete
// ): Promise<any> => {
//   const baseUrl = `${siteAbsoluteUrl}/_api/web/lists/${apiGetType}('${listname}')`;
//   let attachedObj: RequestInit;
//   let queryUrl = "";

//   switch (method) {
//     case CRUD.Update: {
//       queryUrl = baseUrl + `/Items(${id})`;
//       attachedObj = {
//         method: "POST",
//         credentials: "same-origin",
//         headers: {
//           Accept: "application/json",
//           "Content-Type": "application/json;odata=verbose",
//           "X-RequestDigest": requestDigest || "",
//           "IF-MATCH": "*",
//           "X-HTTP-Method": "MERGE",
//         },
//         body: requestdata,
//       };
//       break;
//     }

//     case CRUD.Delete: {
//       queryUrl = baseUrl + `/Items(${id})`;
//       attachedObj = {
//         method: "POST",
//         credentials: "same-origin",
//         headers: {
//           Accept: "application/json",
//           "If-Match": "*",
//           "X-HTTP-Method": "DELETE",
//           "X-RequestDigest": requestDigest || "",
//         },
//       };
//       break;
//     }

//     case CRUD.Create: {
//       queryUrl = baseUrl + "/Items";
//       attachedObj = {
//         method: "POST",
//         credentials: "same-origin",
//         headers: {
//           Accept: "application/json",
//           "Content-Type": "application/json",
//           "X-RequestDigest": requestDigest || "",
//         },
//         body: requestdata,
//       };
//       break;
//     }

//     case CRUD.Read:
//     default: {
//       queryUrl = baseUrl + `/Items`;
//       attachedObj = { headers: { Accept: "application/json" } };
//       break;
//     }
//   }

//   return new Promise<any>((resolve, reject) => {
//     const conditionToCheckIfEmptyId =
//       method !== CRUD.Read && method !== CRUD.Create && itemIsEmpty(id);
//     if (conditionToCheckIfEmptyId) reject({ status: 500 });

//     if (queryApiArgs) queryUrl = queryUrl + `?${queryApiArgs}`;

//     try {
//       fetch(queryUrl, { ...attachedObj, ...addToHeaders })
//         // .then((res) => res.json())
//         .then((data) => resolve(data))
//         .catch((error) => reject(error));
//     } catch (e) {
//       reject(e);
//     }
//   });
// };

// export default crudSpList;
