import { Injectable } from '@angular/core';
import { HttpClient, HttpHeaders } from '@angular/common/http';
import * as hello from 'hellojs/dist/hello.all.js';
import { map } from 'rxjs/operators';
import { Observable } from 'rxjs';

@Injectable()
export class GraphService {
    constructor(private http: HttpClient) {}

    getAccessToken() {
        const msft = hello('msft').getAuthResponse();
        const accessToken = msft.access_token;
        return accessToken;
    }

    executeQuery(queryType, query, postBody?, requestHeaders?: HttpHeaders): Observable<{}> {
        let newHeaders: HttpHeaders = new HttpHeaders();
        if (requestHeaders === undefined) {
            newHeaders = new HttpHeaders({ 'Authorization': 'Bearer ' + this.getAccessToken() });
        } else {
            newHeaders = requestHeaders.append('Authorization', 'Bearer ' + this.getAccessToken()) ||
            newHeaders.append('Authorization', 'Bearer ' + this.getAccessToken());
        }
        switch (queryType) {
            case 'GET':
                return this.http.get(query, { headers: newHeaders });
            case 'GET_BINARY':
                return this.http.get(query, { responseType: 'arraybuffer', headers: newHeaders });
            case 'PUT':
                return this.http.put(query, postBody, { headers: newHeaders });
            case 'POST':
                return this.http.post(query, postBody, { headers: newHeaders });
            case 'PATCH':
                return this.http.patch(query, postBody, { headers: newHeaders });
            case 'DELETE':
                return this.http.delete(query, { headers: newHeaders });
        }
    }
}
