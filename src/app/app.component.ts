import { Component, ChangeDetectorRef } from '@angular/core';
import { GraphAuthService } from './services/graph-auth.service';
import { GraphService } from './services/graph.service';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  title = 'Angular 6, RxJs 6 and the Microsoft Graph API Connect Sample';
  isLogin = false;
  constructor(private graphAuthService: GraphAuthService,
    private graphService: GraphService,
    private cdRef: ChangeDetectorRef) {
    this.graphAuthService.initAuth(graphService, cdRef);
  }
}
