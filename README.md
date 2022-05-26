# Google Workspace (Docs) Templates

These templates are meant to be used with Google Docs. To use them, please import the .odt files to your own drive, and
add the corresponding AppScript and HTML files (you can do so in the Google Docs extension menu).

There is also an AppScript file that is to be similarly inserted into a Google Spreadsheet, it is used for parsing
several/all RRA files you create into a table, for querying purposes.

You may need to customize the files to your own specific needs.

## RRA versions

Each RRA template is under it's own version directory (e.g. `3.0.2`). The associated code is licensed under the MPL.

## Vendor questionnaire

This is the original Mozilla Vendor Questionnaire which you may use for collecting information about various vendors
used by your organization. It is licensed under the MPL.


## Installing the templates

Because these templates use Google Workspace and this project does not own a Google Workspace environment, you should set them up manually. In the future, these may be provided automatically through an add-on instead.

### RRA Document template with RRA Utilities

This is the main RRA template, basically the document you use to run RRAs. You can make your own, but this is a good way to get you started.

1. Go to [Google Drive](https://drive.google.com/) and create a new directory, e.g. `RRA Drive`.
2. Select new and "File Upload". Select the RRA template you want to use (e.g. `RRA 3.2.0.odt`.
3. Open the new document, it should have been imported into a native Google Docs document. Feel free to customize.
4. Select `Extensions>App Script` in the document menu.
5. Copy-paste the RRA utilities (e.g. `RRA.gs`) into the new script and save it (floppy disk icon).
6. Click `Run` with the function `onInstall` selected to check everything works. You should see a permission prompt and warnings if you've never authorized your account to run AppScript. Verify you can see the "RRA Utilities" menu in the Google Doc.
7. That's it! You can now copy this doc for any new RRA.

### RRA Parser

If you use the included RRA Document template, this parser will let you parse several RRA documents and record data into a Google Spreadsheet. This spreadsheet can then be used as a risk register directly, or queried externally to include data into your own risk register.

1. Create a new Google Spreadsheet in your [Google Drive](https://drive.google.com/) and name it (e.g. `RRA Parser`). Name the first tab `RRA3`.
2. Select `Extensions>App Script` in the spreadsheet menu.
3. Copy-paste the RRA parser (e.g. `RRA_Parsing.gs`) into the new script and save it (floppy disk icon).
4. If you use the inline guide, click the `+` to add a new file and select `HTML`, then repeat the copy-paste process for the HTML guide (e.g. with `guide.html` - you will want to customize it). Make sure this document is called `guide` in the App Script interface so that it can be referenced by the script (full name will be `guide.html`).
5. Find your RRA Drive ID, e.g. if your drive URL is `https://drive.google.com/drive/u/0/folders/13iLaDd8BGpVZE7O9p_tKNmNifPLw2qEc` then your Drive ID is `13iLaDd8BGpVZE7O9p_tKNmNifPLw2qEc`, then replace it in the parsing script (replace `YOUR-DRIVE-ID-HERE`).
6. Repeat the process for your RRA template ID if it's in the same drive directory, e.g. if your RRA template document is `https://docs.google.com/document/d/1Hi8W1-VGggDdd84PeIsrA-05SI2RihybCestP-_1ACQ/edit` your RRA template ID is `1Hi8W1-VGggDdd84PeIsrA-05SI2RihybCestP-_1ACQ` (replace `YOUR-TEMPLATE-ID-HERE`).
7. Click `Run` with the function `refreshRRA` selected to check everything works. Similarly to the utilities you should get a permission prompt before the script actually runs.
8. That's it! You can re-run that function every time you want to update your RRA database. Note that if you have an empty table, you will need to create some RRAs first.


Optional: Add automatic refresh trigger.

In order to refresh the RRA database automatically, you'll want to install an AppScript trigger.

1. In the previously added App Script spreadsheet, click the Clock icon (`Triggers`).
2. Select `+ Add Trigger` and keep the defaults, but change `Select Event Source` to `Time-Driven` and select a timer, e.g. every hour.
3. Done!

## License

See the associated `LICENSE` file.

This Source Code Form is subject to the terms of the Mozilla Public
License, v. 2.0. If a copy of the MPL was not distributed with this
file, You can obtain one at https://mozilla.org/MPL/2.0/.
