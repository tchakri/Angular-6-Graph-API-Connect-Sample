import { Component, OnInit, ViewChild, ElementRef, DoCheck } from '@angular/core';
import { GraphAuthService } from '../services/graph-auth.service';
import { GraphService } from '../services/graph.service';
import { Configs, explorerValues } from '../shared/configs';
import { Subscription } from 'rxjs';
import { HttpHeaders } from '@angular/common/http';
import { DomSanitizer, SafeUrl } from '@angular/platform-browser';

@Component({
    selector: 'app-home',
    templateUrl: 'home.component.html',
    styleUrls: ['home.component.css']
})
export class HomeComponent implements OnInit, DoCheck {
    subsSendMail: Subscription;
    me = false;
    userPrincipalName = '';
    displayName: string;
    mail = '';
    emailSent = false;
    postMessage = {};
    profilePicUrl: SafeUrl;
    isImageLoaded = false;
    profilePic: SafeUrl;
    dataLoaded = false;
    _explorerValues = explorerValues;

    constructor(
        private graphAuthService: GraphAuthService,
        private graphService: GraphService,
        private sanitizer: DomSanitizer
    ) {
        this.getMe();
        this.getPhoto();
    }

    ngOnInit() { }

    ngDoCheck() {
        if (explorerValues.authentication.status === 'authenticated') {
            if (explorerValues.showImage) {
                this.profilePic = this.sanitizer.bypassSecurityTrustUrl(explorerValues.authentication.user.profileImageUrl);
                this.dataLoaded = true;
            }
        }
    }

    getHeaders() {
        const headers = new HttpHeaders({
            'Content-Type': 'application/json',
            'Access-Control-Allow-Origin': '*'
        });
        return headers;
    }

    onSendMail() {
        this.postMessage = {
            message: {
                subject: 'Angular 6, RxJs 6 and the Microsoft Graph API Connect Sample',
                toRecipients: [{
                    emailAddress: {
                        address: this.mail || this.userPrincipalName
                    }
                }],
                body: {
                    content: '<html><head> <meta http-equiv=\'Content-Type\' content=\'text/html; charset=us-ascii\'> <title></title>'
                        + '</head><body style=\'font-family:calibri\'> <p>Congratulations ' + this.displayName +
                        `,</p> <p>This is a sample message from the Microsoft Graph Connect sample.</body> </html>`,
                    contentType: 'html'
                }
            }
        };

        this.subsSendMail = this.graphService.executeQuery('POST', Configs.graphV1Url + 'sendMail',
            this.postMessage, this.getHeaders()).subscribe(response => {
                this.emailSent = true;
            });
    }

    onLogout() {
        this.graphAuthService.logout();
    }

    onLogin() {
        this.graphAuthService.login();
    }

    getMe() {
        this.graphService.executeQuery('GET', Configs.graphV1Url)
            .subscribe((result: any) => {
                this.userPrincipalName = result.userPrincipalName;
                this.displayName = result.displayName;
                this.mail = result.mail || result.userPrincipalName;
            },
            error => console.log(error),
            () => this.me = true);
    }

    getPhoto() {
        this.graphService.executeQuery('GET_BINARY', Configs.graphBetaUrl + 'photo/$value')
            .subscribe((result: any) => {
                const blob = new Blob([result], { type: 'image/jpeg' });
                const imageUrl = window.URL.createObjectURL(blob);
                this.profilePicUrl = this.sanitizer.bypassSecurityTrustUrl(imageUrl);
            },
            error => console.log(error),
            () => this.isImageLoaded = true);
    }
}
