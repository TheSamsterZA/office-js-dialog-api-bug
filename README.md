# OfficeJs Dialog Api in Outlook 2016 Bug

```
npm install
npm start
```
Then navigate to https://localhost:3000.

## Context

Within a dialog created by the Dialog API, one has [a very limited set of OfficeJS APIs available](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/dialog-api-in-office-add-ins#send-information-from-the-dialog-box-to-the-host-page). Given this, the dialog page is extremely simple, and delegates any given action to the parent/host page.

The way it works is:
1. The host page `index.html` checks early on if it is in a dialog.
2. When the add-in is loaded in the taskpane, this check fails and the add-in continues on.
3. The add-in opens a dialog. For the sake of illustration, it opens a `dialog.html` page, but this could be a third-party authentication page. It is also at this point that event handlers are attached to the dialog.
4. The `dialog.html` page only does one thing: redirect back to the host page, `index.html`. This mirrors how most third party authentication works: login, send a token, and redirect back to the caller.
5. `index.html` checks if it is in a dialog, which passes. This code then uses [Office.context.ui.messageParent](https://docs.microsoft.com/en-us/javascript/api/office/office.ui?view=office-js#messageparent-message-) to send a message to the host page in the taskpane.
6. The event handlers setup in step #3 are fired. After the message is processed, the host page calls [dialog.close()](https://docs.microsoft.com/en-us/javascript/api/office/office.dialog?view=office-js#close--) to close the dialog.

## The Bug

After sending a message back to the parent, I am seeing the following error in the console:

```
Error: AgaveSink could not find a suitable sink to use
 at e.prototype.failToInitialize (https://appsforoffice.microsoft.com/lib/1.1/hosted/telemetry/oteljs_agave.js:1:71169)
 at e.prototype.initialize (https://appsforoffice.microsoft.com/lib/1.1/hosted/telemetry/oteljs_agave.js:1:70452)
 at e (https://appsforoffice.microsoft.com/lib/1.1/hosted/telemetry/oteljs_agave.js:1:70227)
 at e.createInstance (https://appsforoffice.microsoft.com/lib/1.1/hosted/telemetry/oteljs_agave.js:1:70361)
 at c.create (https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js:18:16634)
 at e (https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js:18:17056)
 at Anonymous function (https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js:18:17204)
 at t.prototype.invoke (https://outlookapp.test.titlex.com.au/polyfills.77e5fe1dedbbcaec49d6.bundle.js:1:88956)
 at e.prototype.run (https://outlookapp.test.titlex.com.au/polyfills.77e5fe1dedbbcaec49d6.bundle.js:1:84204)
 at Anonymous function (https://outlookapp.test.titlex.com.au/polyfills.77e5fe1dedbbcaec49d6.bundle.js:1:95475)
```

I spent some time trying to debug this, and noticed in `https://appsforoffice.microsoft.com/lib/1.1/hosted/telemetry/oteljs_agave.debug.js`:

```
3806    AgaveSink.prototype.initialize = function() {
3807        if (!this.isTelemetryEnabled()) {
3808            this.failToInitialize(E_TELEMETRY_DISABLED);
3809        }
3810        if (OutlookSink_OutlookSink.isSupported()) {
3811            this.connectOutlookSink();
3812        } else if (SdxWacSink_SdxWacSink.isSupported()) {
3813            this.connectSdxWacSink();
3814        } else if (isWacAgave() || typeof OfficeExtension === "undefined") {
3815            if (canSendToAria()) {
3816                this.connectAriaSink();
3817            } else {
3818                this.failToInitialize();
3819            }
3820        } else {
3821            this._awaitingInitialization = true;
3822            getRichApiSink(false, this.onGetRichApi.bind(this));
3823        }
3824    };
```

I believe the bug is in here or in one of the functions called within here. There are assumptions about retrieving the host information which I believe cannot be done when in an OfficeJS dialog.

## Versions

```
Outlook 2016 on Windows Version 1908 (Build 11911.20000 Click-to-Run), Office Insider [Microsoft Outlook for Office 365 MSO (16.0.11911.20000)]

Windows 10 Version 1903 (OS Build 18362.239)

Microsoft Edge browser engine: EdgeHTML 18.18362 (so we are not using Chromium), and Microsoft Edge DevTools Preview
```

