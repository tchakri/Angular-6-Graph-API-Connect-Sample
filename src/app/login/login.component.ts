import { Component } from '@angular/core';

import { GraphAuthService } from '../services/graph-auth.service';

@Component({
  selector: 'app-login',
  templateUrl: 'login.component.html'
})
export class LoginComponent {
  constructor(private graphAuthService: GraphAuthService) {}

  onLogin() {
    this.graphAuthService.login();
  }
}
