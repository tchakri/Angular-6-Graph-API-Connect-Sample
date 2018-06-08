import { Injectable, NgZone, ChangeDetectorRef } from '@angular/core';
import { Router } from '@angular/router';
import * as hello from 'hellojs/dist/hello.all.js';

import { of, Observable, forkJoin } from 'rxjs';
import { Configs, explorerValues } from '../shared/configs';
import { GraphService } from './graph.service';

@Injectable()
export class GraphAuthService {
    email: string;
    constructor(
        private zone: NgZone,
        private router: Router
    ) {}

    initAuth(graphService: GraphService, changeDetectorRef: ChangeDetectorRef) {
        setInterval(this.refreshAccessToken, 1000 * 60 * 10);
        hello.init({
            msft: {
                id: Configs.appId,
                oauth: {
                    version: 2,
                    auth: Configs.graphAuthUrl
                },
                scope_delim: ' ',
                form: false
            },
        },
            { redirect_uri: '/' }
        );

        hello.on('auth.login', function (auth) {
            let accessToken;
            if (auth.network === 'msft') {
                const authResponse = hello('msft').getAuthResponse();
                accessToken = authResponse.access_token;
            }
            if (accessToken) {
                explorerValues.authentication.status = 'authenticating';
                changeDetectorRef.detectChanges();
                explorerValues.authentication.user = {};
                const params = '?$select=displayName,mail,userPrincipalName,jobTitle,mobilePhone,department,officeLocation';

                const userData = graphService.executeQuery('GET', Configs.graphV1Url + params);
                const userPhoto = graphService.executeQuery('GET_BINARY', Configs.graphBetaUrl + 'photo/$value');

                forkJoin(userData, userPhoto)
                    .subscribe((result: any) => {
                        explorerValues.authentication.user.displayName = result[0].displayName;
                        explorerValues.authentication.user.mail = result[0].mail || result[0].userPrincipalName;
                        explorerValues.authentication.user.jobTitle = result[0].jobTitle;
                        explorerValues.authentication.user.mobilePhone = result[0].mobilePhone;
                        explorerValues.authentication.user.officeLocation = result[0].officeLocation;
                        explorerValues.authentication.user.department = result[0].department;

                        const blob = new Blob([result[1]], { type: 'image/jpeg' });
                        const imageUrl = window.URL.createObjectURL(blob);
                        explorerValues.authentication.user.profileImageUrl = imageUrl;
                        explorerValues.showImage = true;

                        explorerValues.authentication.status = 'authenticated';
                        changeDetectorRef.detectChanges();
                    }, e => {
                        console.log(e);
                    });
            }
        });
        explorerValues.authentication.status = this.haveValidAccessToken() ? 'authenticating' : 'anonymous';
    }

    login() {
        const loginProperties = {
            display: 'popup',
            mkt: Configs.Language,
            scope: Configs.scope
        };
        hello('msft').login(loginProperties).then(
            () => {
                this.zone.run(() => {
                    this.router.navigate(['/home']);
                });
            },
            e => console.error(e.error.message)
        );
    }

    logout() {
        hello('msft').logout().then(
            () => window.location.href = '/',
            e => console.error(e.error.message)
        );
        explorerValues.authentication.status = 'anonymous';
        explorerValues.authentication.user = {};
    }

    localLogout() {
        explorerValues.selectedOption = 'GET';
        if (typeof hello !== 'undefined') {
            hello('msft').logout(null, { force: true });
        }
        explorerValues.authentication['status'] = 'anonymous';
        explorerValues.authentication.user = {};
    }

    refreshAccessToken() {
        if (explorerValues.authentication.status !== 'authenticated') {
            console.log('Not refreshing access token since user is logged out or currently logging in.', new Date());
            return;
        }
        explorerValues.authentication.status = 'authenticating';
        const loginProperties = {
            display: 'none',
            response_type: 'token',
            response_mode: 'fragment',
            nonce: 'graph_explorer',
            prompt: 'none',
            scope: Configs.scope,
            login_hint: explorerValues.authentication.user.mail,
            domain_hint: 'organizations'
        };
        const silentLoginRequest = hello('msft').login(loginProperties);
        silentLoginRequest.then(function() {
            console.log('Successfully refreshed access token.', new Date());
        }, function(e) {
            console.error('Error refreshing access token', e, new Date());
            this.checkHasValidAuthToken();
        });
    }

    checkHasValidAuthToken() {
        if (!this.haveValidAccessToken() && this.isAuthenticated()) {
            console.log('App says user is authenticated, but doesn\'t have a valid access token.', new Date());
            this.logout();
        }
    }

    haveValidAccessToken() {
        const session = hello('msft').getAuthResponse();
        if (!session) {
            return false;
        }
        const currentTime = (new Date()).getTime() / 1000;
        return session && session.access_token && session.expires > currentTime;
    }

    isAuthenticated() {
        return explorerValues.authentication['status'] !== 'anonymous';
    }

}
