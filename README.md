# ANGULAR 11, GRAPH EXPLORER API CONNECT

First Register your app in Microsoft Azure AD. For reference: 
https://developer.microsoft.com/en-us/graph/docs/concepts/auth_register_app_v2

You can make use of sample queries to Graph API using this url:
https://developer.microsoft.com/en-us/graph/graph-explorer

This project was generated with [Angular CLI](https://github.com/angular/angular-cli) version 11.2.11.

## Run this project

1. Clone or Download the code to your folder.
2. Open `configs.ts` file under location: `\ANGULAR6-GRAPH-API-CONNECT\src\app\shared`.
3. Update your `appId` (example: in place of 'YOUR-CLIENTID-OF-REGISTERED-APP' to 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx) and then save the file.
4. Open Command Prompt (Windows + R => type 'cmd' => Enter) and change directory to the downloaded location, say, c:\yourpath\ANGULAR6-GRAPH-API-CONNECT.
5. Type command: `npm install` and then press `Enter`; node_modules gets downloaded.
6. Type command: `ng serve -o` and then press `Enter` to run the project.
7. Browser will open and click on `Connect` button. A popup gets open for authentication.
8. Login to your personal account.

## Development server

Run `ng serve` for a dev server. Navigate to `http://localhost:4200/`. The app will automatically reload if you change any of the source files.

## Code scaffolding

Run `ng generate component component-name` to generate a new component. You can also use `ng generate directive|pipe|service|class|guard|interface|enum|module`.

## Build

Run `ng build` to build the project. The build artifacts will be stored in the `dist/` directory. Use the `--prod` flag for a production build.

## Running unit tests

Run `ng test` to execute the unit tests via [Karma](https://karma-runner.github.io).

## Running end-to-end tests

Run `ng e2e` to execute the end-to-end tests via [Protractor](http://www.protractortest.org/).

## Further help

To get more help on the Angular CLI use `ng help` or go check out the [Angular CLI README](https://github.com/angular/angular-cli/blob/master/README.md).
