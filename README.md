DRIVEPM
=======

Here we have some scripts and VBA code that works together to integrate project clients with drive spreadsheets servers for easy reporting and collaboration.

Acknowledgements
------ 

+ we reuse some nice piece of code found at desktop liberation by Bruce McPherson, http://ramblings.mcpher.com

The code on this repo is free to use/modify. You may re-use or distribute anything as you wish, with appropriate acknowledgement under a creative commons share-alike license.
If you wish to contribute code, improvements, bugfixes, articles, comments to this repo, or simply to connect, I would be delighted to hear from you.

Tools available
------ 


### Sync project client with google drive server ###

This tool is able to sync main project objects and properties with a google spreadsheet acting as a EPM server.
Users may work in colaboration and share data between themselves and between project and google spreadsheets.
At actual development sync routines are not fully written and may have data loss during push and pull rounds with concurrent data. 

TODO: screenshots and examples

How to use
------ 

+ open project file with macros enabled and at VBA for application references add these: Microsoft Internet Controls, Microsoft XML
+ add classes: cBrowser, cJobject, cOauth2, cSheetsV4 and cStringChunker
+ add modules: functions, drivEPM
+ config a project at google developer console and create oauth client_id and client_secret
+ setup VBA procedure RunOnce at drivEPM module with SheetID, ClientID, ClientSecret and uncomment these lines to run at least once

TODO: write more stuff here

Wiki
------ 
Take a look at our wiki at https://github.com/vanaware/drivepm/wiki for more information
