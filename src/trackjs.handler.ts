///<reference path="../node_modules/trackjs/tracker.d.ts"/>
import { ErrorHandler, Injectable } from '@angular/core';

@Injectable()
export class TrackJsErrorHandler implements ErrorHandler {
  handleError(error: any) {
    // Add the error message to the telemetry timeline.
    // It can occasionally have useful additional context.
    console.warn(error.message);

    // Assumes we have already loaded and configured TrackJS
    if (trackJs) {
      trackJs.track(error.originalError); // Send the native error object to TrackJS
    }
  }
}
