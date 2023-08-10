import { Injectable } from '@angular/core';
import { HttpClient, HttpHeaders, HttpErrorResponse } from '@angular/common/http';
import { Observable, throwError } from 'rxjs';
import { catchError, map, mergeMap } from 'rxjs/operators';

@Injectable()
export class Spcrudnew {

  jsonHeader = 'application/json; odata=verbose';
  headers = new HttpHeaders({ 'Content-Type': this.jsonHeader, 'Accept': this.jsonHeader });
  options = { headers: this.headers };
  baseUrl: any ='';
  apiUrl: string='';
  currentUser: string='';
  login: string='';

  constructor(private http: HttpClient) {
      this.setBaseUrl();
  }

  private handleError(error: HttpErrorResponse | any) {
    let errMsg: string;
    if (error instanceof HttpErrorResponse) {
      errMsg = `${error.status || ''} - ${error.statusText || ''} ${error.message}`;
    } else {
      errMsg = error.message ? error.message : error.toString();
    }
    console.error(errMsg);
    return throwError(errMsg);
  }

  private endsWith(str: string, suffix: string) {
    return str.indexOf(suffix, str.length - suffix.length) !== -1;
  }

  setBaseUrl(webUrl?: string) {
    if (webUrl) {
      this.baseUrl = webUrl;
    } else {
      //const ctx = window['_spPageContextInfo'];
     const ctx = (window as any)._spPageContextInfo;
      if (ctx) {
        this.baseUrl = ctx.webAbsoluteUrl;
      }
    }

    this.apiUrl = `${this.baseUrl}/_api/web/lists/GetByTitle('{0}')/items`;

    const el:any = document.querySelector('#__REQUESTDIGEST');
    if (el) {
      this.headers = this.headers.set('X-RequestDigest', el.nodeValue);
    }
  }

  refreshDigest(): Observable<any> {
    return this.http.post<any>(`${this.baseUrl}/_api/contextinfo`, null, this.options)
      .pipe(
        map(res => {
          const formDigestValue = res.d.GetContextWebInformation.FormDigestValue;
          this.headers = this.headers.set('X-RequestDigest', formDigestValue);
        }),
        catchError(this.handleError)
      );
  }
 
// send email

sendMail(to: string, ffrom: string, subj: string, body: string): Observable<any> {
    const tos: string[] = to.split(',');
    const recip: string[] = (tos instanceof Array) ? tos : [tos];
    const message = {
      'properties': {
        '__metadata': {
          'type': 'SP.Utilities.EmailProperties'
        },
        'To': {
          'results': recip
        },
        'From': ffrom,
        'Subject': subj,
        'Body': body
      }
    };
    const url = `${this.baseUrl}/_api/SP.Utilities.Utility.SendEmail`;
    const data = JSON.stringify(message);

    return this.http.post<any>(url, data, this.options).pipe(
      catchError(this.handleError)
    );
  }

// ----------SHAREPOINT USER PROFILES----------
  getCurrentUser(): Observable<any> {
    const url = `${this.baseUrl}/_api/web/currentuser?$expand=Groups`;

    return this.http.get<any>(url, this.options).pipe(
      map(res => res),
      catchError(this.handleError)
    );
  }

  // get my profile 
  getMyProfile(): Observable<any> {
    const url = `${this.baseUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties?select=*`;

    return this.http.get<any>(url, this.options).pipe(
      map(res => res),
      catchError(this.handleError)
    );
  }

  // Look Up any sharepoint profile 
  getProfile(login: string): Observable<any> {
    const url = `${this.baseUrl}/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='${login}'&select=*`;

    return this.http.get<any>(url).pipe(
      map((res: any) => res),
      catchError(this.handleError)
    );
  }


  // Get Any UserInfo
  getUserInfo(id: string): Observable<any> {
    const url = `${this.baseUrl}/_api/web/getUserById(${id})`;

    return this.http.get<any>(url).pipe(
      map((res: any) => res),
      catchError(this.handleError)
    );
  }
   //check User Exists

  ensureUser(login: string): Observable<any> {
    const url = `${this.baseUrl}/_api/web/ensureuser`;

    return this.http.post<any>(url, login, this.options).pipe(
      map((res: any) => res),
      catchError(this.handleError)
    );
  }
    ///-------------LIST CORE -----------//

    // ----------SHAREPOINT LIST AND FIELDS----------

  // Create list
    createList(title: string, baseTemplate: string, description: string): Observable<any> {
        const data = {
          '__metadata': { 'type': 'SP.List' },
          'BaseTemplate': baseTemplate,
          'Description': description,
          'Title': title
        };
        const url = `${this.baseUrl}/_api/web/lists`;
    
        return this.http.post<any>(url, data, this.options).pipe(
          map((res: any) => res),
          catchError(this.handleError)
        );
      }
      //CREATE FIELD
      createField(listTitle: string, fieldName: string, fieldType: string): Observable<any> {
        const data = {
          '__metadata': { 'type': 'SP.Field' },
          'Type': fieldType,
          'Title': fieldName
        };
        const url = `${this.baseUrl}/_api/web/lists/GetByTitle('${listTitle}')/fields`;
    
        return this.http.post<any>(url, data, this.options).pipe(
          map((res: any) => res),
          catchError(this.handleError)
        );
      }

    // ----------SHAREPOINT FILES AND FOLDERS----------
     
    // Create folder
    createFolder(folderUrl: string): Observable<any> {
        const data = {
          '__metadata': {
            'type': 'SP.Folder'
          },
          'ServerRelativeUrl': folderUrl
        };
        const url = `${this.baseUrl}/_api/web/folders`;
    
        return this.http.post<any>(url, data, this.options).pipe(
          map((res: any) => res),
          catchError(this.handleError)
        );
      }
   // Upload file to folder
  // https://kushanlahiru.wordpress.com/2016/05/14/file-attach-to-sharepoint-2013-list-custom-using-angular-js-via-rest-api/
  // http://stackoverflow.com/questions/17063000/ng-model-for-input-type-file
  // var binary = new Uint8Array(FileReader.readAsArrayBuffer(file[0]));
      uploadFile(folderUrl: string, fileName: string, binary: any): Observable<any> {
        const url = `${this.baseUrl}/_api/web/GetFolderByServerRelativeUrl('${folderUrl}')/files/add(overwrite=true, url='${fileName}')`;
    
        return this.http.post<any>(url, binary, this.options).pipe(
          map((res: any) => res),
          catchError(this.handleError)
        );
      }
    
    // Upload Attachment 
    uploadAttach(listName: string, id: string, fileName: string, binary: any, overwrite?: boolean): Observable<any> {
        let url = `${this.baseUrl}/_api/web/lists/GetByTitle('${listName}')/items(${id})`;
        const headers = this.options.headers;
    
        if (overwrite) {
          // Append HTTP header PUT for UPDATE scenario
          headers.append('X-HTTP-Method', 'PUT');
          url += `/AttachmentFiles('${fileName}')/$value`;
        } else {
          // CREATE scenario
          url += `/AttachmentFiles/add(FileName='${fileName}')`;
        }
    
        return this.http.post<any>(url, binary, { headers }).pipe(
          map((res: any) => res),
          catchError(this.handleError)
        );
      }

       // Get attachment for item
       getAttach(listName: string, id: string): Observable<any> {
        const url = `${this.baseUrl}/_api/web/lists/GetByTitle('${listName}')/items(${id})/AttachmentFiles`;
    
        return this.http.get<any>(url, this.options).pipe(
          map((res: any) => res),
          catchError(this.handleError)
        );
      }

      //// Copy file

      copyFile(sourceUrl: string, destinationUrl: string): Observable<any> {
        const url = `${this.baseUrl}/_api/web/GetFileByServerRelativeUrl('${sourceUrl}')/copyto(strnewurl='${destinationUrl}',boverwrite=false)`;
    
        return this.http.post<any>(url, null, this.options).pipe(
          map((res: any) => res),
          catchError(this.handleError)
        );
      }

     // ----------SHAREPOINT LIST CORE----------

  // CREATE item - SharePoint list name, and JS object to stringify for save
  create(listName: string, jsonBody: any): Observable<any> {
    const url = this.apiUrl.replace('{0}', listName);

    // append metadata
    if (!jsonBody.__metadata) {
      jsonBody.__metadata = {
        'type': 'SP.ListItem'
      };
    }

    const data = JSON.stringify(jsonBody);

    return this.http.post<any>(url, data, this.options).pipe(
      map((res: any) => res),
      catchError(this.handleError)
    );
  }

    // Build URL string with OData parameters

  readBuilder(url: string, options?: any): string {
    if (options) {
      const queryParams = [];

      if (options.filter) {
        queryParams.push(`$filter=${options.filter}`);
      }
      if (options.select) {
        queryParams.push(`$select=${options.select}`);
      }
      if (options.orderby) {
        queryParams.push(`$orderby=${options.orderby}`);
      }
      if (options.expand) {
        queryParams.push(`$expand=${options.expand}`);
      }
      if (options.top) {
        queryParams.push(`$top=${options.top}`);
      }
      if (options.skip) {
        queryParams.push(`$skip=${options.skip}`);
      }

      if (queryParams.length > 0) {
        url += (url.includes('?') ? '&' : '?') + queryParams.join('&');
      }
    }
    
    return url;
  }
// READ entire list - needs $http factory and SharePoint list name

  read(listName: string, options?: any): Observable<any> {
    let url = this.apiUrl.replace('{0}', listName);
    url = this.readBuilder(url, options);

    return this.http.get<any>(url, this.options).pipe(
      map((resp: any) => resp),
      catchError(this.handleError)
    );
  }

  // READ single item - SharePoint list name, and item ID number

  readItem(listName: string, id: string): Observable<any> {
    let url = this.apiUrl.replace('{0}', listName) + '(' + id + ')';
    url = this.readBuilder(url);

    return this.http.get<any>(url, this.options).pipe(
      map((resp: any) => resp),
      catchError(this.handleError)
    );
  }

  // UPDATE item - SharePoint list name, item ID number, and JS object to stringify for save

    update(listName: string, id: string, jsonBody: any): Observable<any> {
        const localOptions = ({
          headers: this.options.headers.append('X-HTTP-Method', 'MERGE').append('If-Match', '*')
        });
    
        if (!jsonBody.__metadata) {
          jsonBody.__metadata = {
            'type': 'SP.ListItem'
          };
        }
    
        const data = JSON.stringify(jsonBody);
        const url = this.apiUrl.replace('{0}', listName) + '(' + id + ')';
        
        return this.http.post<any>(url, data, localOptions).pipe(
          map((resp: any) => resp),
          catchError(this.handleError)
        );
      }

     // DELETE item - SharePoint list name and item ID number
      del(listName: string, id: string): Observable<any> {
        const localOptions =({
          headers: this.options.headers.append('X-HTTP-Method', 'DELETE').append('If-Match', '*')
        });
    
        const url = this.apiUrl.replace('{0}', listName) + '(' + id + ')';
        
        return this.http.post<any>(url, '', localOptions).pipe(
          map((resp: any) => resp),
          catchError(this.handleError)
        );
      }
      
   // JSON blob read from SharePoint list - SharePoint list name
   
      jsonRead(listName: string): Observable<any> {
        return this.getCurrentUser().pipe(
          mergeMap((res: any) => {
            const svc = this;
            svc.currentUser = res.d;
            svc.login = res.d.LoginName.toLowerCase();
    
            if (svc.login.indexOf('\\')) {
              svc.login = svc.login.split('\\')[1];
            }
    
            const url = svc.apiUrl.replace('{0}', listName) + `?$select=JSON,Id,Title&$filter=Title eq '${svc.login}'`;
            
            return svc.http.get<any>(url, svc.options).pipe(
              map((res2: any) => {
                const d2 = res2.d;
                if (d2.results.length) {
                  return d2.results[0];
                } else {
                  return null;
                }
              }),
              catchError(svc.handleError)
            );
          })
        );
      }

        // JSON blob upsert write to SharePoint list - SharePoint list name and JS object to stringify for save
      jsonWrite(listName: string, jsonBody: any): Observable<any> {
        return this.refreshDigest().pipe(
          mergeMap((res: any) => {
            return this.jsonRead(listName).pipe(
              mergeMap((item: any) => {
                if (item) {
                  // update if found
                  item.JSON = JSON.stringify(jsonBody);
                  return this.update(listName, item.Id, item);
                } else {
                  // create if missing
                  item = {
                    '__metadata': {
                      'type': 'SP.ListItem'
                    },
                    'Title': this.login,
                    'JSON': JSON.stringify(jsonBody)
                  };
                  return this.create(listName, item);
                }
              }),
              catchError(this.handleError)
            );
          })
        );
      }

   
}
