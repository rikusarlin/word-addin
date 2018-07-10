# Word Web Add-in project 
Starting point was [Word Tutorial](https://docs.microsoft.com/en-us/office/dev/add-ins/tutorials/word-tutorial)

The original contents have been mostly commented out, and content controls related to Finnish "goverment proposal" (hallituksen esitys; a proposal to parliament for new legislation) have been added. This in "Start" folder.

Techlonogies used
* Node.js
* npm
* jquery
* Word Javascript API, mostly content controls

## Prerequisites

* An Office 365 account.
- Word 2016, version 1711 (Build 8730.1000 Click-to-Run) or later. You might need to be an Office Insider to get this version. For more information, see [Be an Office Insider](https://products.office.com/en-us/office-insider?tab=tab-1).
- [Node and npm](https://nodejs.org/en/) 
- [Git Bash](https://git-scm.com/downloads) (Or another git client.)

## To use the project

This sample is meant to accompany the tutorials that begin at: [Word Tutorial](https://docs.microsoft.com/en-us/office/dev/add-ins/tutorials/word-tutorial).

Install packages, update npm and the relevant packages as follows:
* npm install
* npm install npm@latest -g
* npm audit

Then compile and run:
* npm run build
* npm start

After this, you need to copy the manifest to a place trusted by Word and "sideload" the add-in. Further information via the link above.

## Additional resources

* [Office add-in documentation](https://docs.microsoft.com/en-us/office/dev/add-ins/)
* [Office Dev Center](https://developer.microsoft.com/office)
* More Office Add-in samples at [OfficeDev on Github](https://github.com/officedev)

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Copyright
Copyright (c) 2018 Microsoft Corporation. All rights reserved.

