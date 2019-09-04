import { BrowserModule } from '@angular/platform-browser';
import { NgModule, ErrorHandler } from '@angular/core';

import { AppRoutingModule } from './app-routing.module';

import { AppComponent } from './app.component';

import { TrackJsErrorHandler } from '../trackjs.handler';

@NgModule({
  declarations: [
    AppComponent
  ],
  imports: [
    BrowserModule,
    AppRoutingModule
  ],
  providers: [
    { provide: ErrorHandler, useClass: TrackJsErrorHandler } // Log errors to TrackJS
  ],
  bootstrap: [AppComponent]
})
export class AppModule { }
