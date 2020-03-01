import * as _ from "lodash";
import { intersectionWith } from "lodash";

export const log = (value: any, optionalParams?: string) => {
  if (
    window.location.search &&
    window.location.search.length != 0 &&
    window.location.search.toLocaleLowerCase().indexOf("loadSPFX") < 0
  ) {
    if (optionalParams) console.log(value, optionalParams);
    else console.log(value);
  }
};

export const getSitesStartingWith = (
  context: any,
  configurations: any
): Promise<string[]> => {
  const startingUrl =
    window.location.protocol + "//" + window.location.hostname;
  const getPathsFromResults = (results: any): string[] => {
    let urls: string[] = [];
    let pathIndex = null;

    for (let result of results.PrimaryQueryResult.RelevantResults.Table.Rows) {
      // Stores the index of the "Path" cell on the first loop in order to avoid finding the cell on every loop
      if (!pathIndex) {
        let pathCell = result.Cells.filter(cell => {
          return cell.Key == "Path";
        })[0];
        pathIndex = result.Cells.indexOf(pathCell);
      }
      urls.push(result.Cells[pathIndex].Value.toLowerCase().trim());
    }
    return urls;
  };

  const getSearchResults = (
    webUrl: string,
    queryParams: string
  ): Promise<any> => {
    return new Promise<any>((resolve, reject) => {
      let endpoint = `${webUrl}/_api/search/query?${queryParams}`;

      context.spHttpClient
        .get(endpoint, configurations)
        .then(response => {
          if (response.ok) {
            resolve(response.json());
          } else {
            reject(response.statusText);
          }
        })
        .catch(error => {
          reject(error);
        });
    });
  };

  const getSearchResultsRecursive = (
    webUrl: string,
    queryParams: string
  ): Promise<any> => {
    return new Promise<any>((resolve, reject) => {
      // Executes the search request for a first time in order to have an idea of the returned rows vs total results
      getSearchResults(webUrl, queryParams)
        .then((results: any) => {
          // If there is more rows available...
          let relevantResults = results.PrimaryQueryResult.RelevantResults;
          let initialResults: any[] = relevantResults.Table.Rows;

          if (
            relevantResults.TotalRowsIncludingDuplicates >
            relevantResults.RowCount
          ) {
            // Stores and executes all the missing calls in parallel until we have ALL results
            let promises = new Array<Promise<any>>();
            let nbPromises = Math.ceil(
              relevantResults.TotalRowsIncludingDuplicates /
                relevantResults.RowCount
            );

            for (let i = 1; i < nbPromises; i++) {
              promises.push(getSearchResults(webUrl, queryParams));
            }

            // Once the missing calls are done, concatenates their results to the first request
            Promise.all(promises).then(values => {
              for (let recursiveResults of values) {
                initialResults = initialResults.concat(
                  recursiveResults.PrimaryQueryResult.RelevantResults.Table.Rows
                );
              }
              results.PrimaryQueryResult.RelevantResults.Table.Rows = initialResults;
              results.PrimaryQueryResult.RelevantResults.RowCount =
                initialResults.length;
              resolve(results);
            });
          }
          // If no more rows are available
          else {
            resolve(results);
          }
        })
        .catch(error => {
          reject(error);
        });
    });
  };

  return new Promise<string[]>((resolve, reject) => {
    let queryProperties = `querytext='Path:${startingUrl}/* AND contentclass:STS_Site'&selectproperties='Path'&trimduplicates=false&rowLimit=500&Properties='EnableDynamicGroups:true'`;

    getSearchResultsRecursive(startingUrl, queryProperties)
      .then((results: any) => {
        resolve(getPathsFromResults(results));
      })
      .catch(error => {
        reject(error);
      });
  });
};
