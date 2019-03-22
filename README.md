# IEEE_Google_Script_Merge
Bound Google Script to merge Sheets fields into Docs and Slides - bulk awards generation

## Background
For at least thirty years the IEEE Northern VA Section awarded organizational prizes at four regional fairs to the best projects related to electrical engineering.  Up to six projects per fair could be awarded a prize, and team projects could have two or three students.  Since 2014, each student received a certificate and cover letter.

At a roundtable discussion in 2016, with the science coordinators from all the fairs agreed it was better to recognize more students than to give larger prizes, so we started giving one "Technology Innovation", up to five "Technology Excellence", and up to three "Technology Achievement" awards beginning in 2017, and using Microsoft Office's built-in mail merge to generate cover letters tailor to each award from three separate spreadsheets.

To comply with IEEE's GDPR guidelines and their Guidelines for Working with Children introduced in 2018, winners of the 2019 fair (including a fifth regional event) were entered in a Google sheet.  There are several Add-Ons (i.e. AutoCrat) that allow merging sheet data with documents and slides and that will probably not violate GDPR guidelines, but without code review and/or tightened access restrictions it's impossible to be certain.

I wrote this bound Google Script to bulk generate cover letters and certificates from data in a single winners spreadsheet.  It can also be used to generate award certificates for the banquet and other events.  You're encouraged to modify it to suit your section's needs.  If there's interest from other sections perhaps IEEE could support it as an Add-On.

## Preparation
To use this script:
1. Create the certificate template on your Google Drive: If you plan to create certificates, create a template.  Data from the spreadsheet will be substituted for strings of the form "<<keyword>>".  You can create the template using Microsoft PowerPoint, but after uploading to the Google Drive open it to convert it to a Google slide that omits the ".pptx" or similar extension.  There can be only one certificate template for all awards, and this can be omitted if you just want to create personalized letters.
2. Create the cover letter template(s) on your Google Drive: Create a template using the same substitution indicators as certificates.  If using Word upload and open on Google Drive to create a Google Doc without the .docx or similar extension.  Each row in the spreadsheet can use a different template.
3. Create the spreadsheet containing the data to merge on your Google Drive: The first row must containing column headings.  One column must be labeled "CoverTemplate"; create a column label "Last" for the last name if you want it to appear in the name of the individual cover letters generated.
4. Import the script: From the sheets application, select "Tools" and then "<> Script Editor".  Paste this script into the editor field and then use the "Run" pull-down menu to execute the "onOpen()" function.  (You generally only need to do this once.)  Go back to the sheet and within a few seconds a "Merge" pull-down menu option will appear.
5. Select whether you want to generate cover letters or certificates.  The first time this is invoked you may be prompted to authorize the script to access your drive and google API's.  It's okay to do this since the code can be inspected to verify no data will be modified or leaked.
  
## How It Works
To be added.
// Required enabling the Slides API!
// No need to open or close the presentation! var preso = SlidesApp.openById(copyId);


## Resources
Some of the URL's referenced:
* https://yagisanatode.com/2017/12/13/google-apps-script-iterating-through-ranges-in-sheets-the-right-and-wrong-way/
* https://opensourcehacker.com/2013/01/21/script-for-generating-google-documents-from-google-spreadsheet-data-source/
* https://developers.google.com/apps-script/reference/drive/drive-app#getFilesByName(String)
* https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/String/padStart
* https://developers.google.com/apps-script/reference/document/body#replacetextsearchpattern-replacement
* https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet-app
* https://developers.google.com/apps-script/guides/dialogs
* https://developers.google.com/slides/how-tos/merge
