# IEEE_Google_Script_Merge
Bound Google Script to merge Sheets fields into Docs and Slides for bulk award generation

## Background
For at least thirty years the IEEE Northern VA Section awarded organizational prizes at four regional fairs to the best projects related to electrical engineering.  Up to six projects per fair could be awarded a prize, and team projects could have two or three students each.  Starting in 2014 each student receives a certificate and cover letter.

At a roundtable discussion in 2016, with the science coordinators from all the fairs agreed it was better to recognize more students than to give larger prizes, so we started giving one "Technology Innovation", up to five "Technology Excellence", and up to three "Technology Achievement" awards beginning in 2017, and using Microsoft Office's built-in mail merge to generate cover letters tailor to each award from three separate spreadsheets.

To comply with IEEE's GDPR guidelines and their Guidelines for Working with Children introduced in 2018, for the 2019 season (which also includes a fifth regional event) I wanted to generate all awards off a single Google Sheet and print them without ever touching my local disk.  There are several Add-Ons (i.e. AutoCrat) that allow merging sheet data with documents and slides and that will probably not violate GDPR guidelines, but without code review and/or tightened access restrictions it's impossible to be certain.

I wrote this bound Google Script to bulk generate cover letters and certificates from data in a single winners spreadsheet.  It can also be used to generate award certificates for the banquet and other events; you are welcome to modify it to suit your needs.  If there's interest from other sections perhaps IEEE could support it as an Add-On.  Note that OS X downloads the files to disk before printing anyway and Windows may also, so it doesn't completely satisfy the original goals but it does consolidate all winners in a single place and reduce cutting and pasting which reduces the changes of errors.

## Preparation
To use this script:
1. Create the certificate template on your Google Drive: If you plan to create certificates, create a template.  Data from the spreadsheet will be substituted for strings of the form "<<keyword>>".  You can create the template using Microsoft PowerPoint, but after uploading to the Google Drive open it to convert it to a Google slide that omits the ".pptx" or similar extension.  There can be only one certificate template for all awards, and this can be omitted if you just want to create personalized letters.
2. Create the cover letter template(s) on your Google Drive: Create a template using the same substitution indicators as certificates.  If using Word upload and open on Google Drive to create a Google Doc without the .docx or similar extension.  Each row in the spreadsheet can use a different template.
3. Create the spreadsheet containing the data to merge on your Google Drive: The first row must containing column headings.  One column must be labeled "CoverTemplate"; create a column label "Last" for the last name if you want it to appear in the name of the individual cover letters generated.
4. Import the script: From the sheets application, select "Tools" and then "<> Script Editor".  Paste this script into the editor field and then use the "Run" pull-down menu to execute the "onOpen()" function.  (You generally only need to do this once.)  Go back to the sheet and within a few seconds a "Merge" pull-down menu option will appear.
5. Select whether you want to generate cover letters or certificates.  The first time this is invoked you may be prompted to authorize the script to access your drive and google API's.  It's okay to do this since the code can be inspected to verify no data will be modified or leaked.
  
## How It Works
This was not my first Javascript/ECMAscript-like project, but it was the first time writing in the Google Script environment, so there are definitely things can could be cleaned up and improved, and there is probably a better overall architecture - please feel free to write it!  The script has are three types of functions:

### onOpen()
This is the function that adds pulldown menu options to Google Sheets.  It's invoked when the script is Run in step 4 above.

### coverLetters(), certificates(), addressLabels()
Each of these three functions are invoked by one of the pulldown menu options added to the sheet.  All three loop over rows of the spreadsheet beginning with the second to the column headings.  Both coverLetters() and certificates() create one new file for each visited row, naming the file with both a zero-padded row number and the winner's last name.  Note that Google sheet allows multiple files in the same directory to have the same name (and we have had multiple winners in the same year with the same last name... from the same fair!), but the row number helps associate the file with the entry.

For each visited row, coverLetters() loops over every column heading and searches for it bracketed by "<<" and ">>.  If found it replaces it with that column's value in the current row using the body.replaceText() function, allowing users to define their own substitution values.  The addressLabels() function uses the same replaceText() call replaces the words "Last", "First", "Address1", and "Address2" with the corresponding cell values.  It also presume there are six mailing labels per sheet so it does not generate as many

The API's to manipulate Google Slides are very powerful, allowing text to be modified, graphics to be inserted, slides to be added, and more, but accomplished by creating a JSON object and passing it to the batchUpdate() method.  It's not as easy to read the JSON.stringify code that builds the object, but it seems to work faster than the changes to Docs.

### getIndex(), createDuplicateDocument()
These functions are called by two or three of the above menu functions.  The getIndex() function takes the array of column headers, the number of column headers, and the name of a specific header and returns the index of the column in which the specific header is found or -1 if it's not found.  The createdDuplicateDocument() function takes the name of a file to be copied and the name of a copy to make in the same directory as the original.

## Enhancements
As the first iteration of this script there are numerous possibilities for improvement in several categories.

### Reducing Redundancy
The name of the award and the template are redundant - it would be easy to define a table in another sheet that maps the award name to the award letter template file, then create a pulldown field so the user can select the award.  The value is somewhat mitigated because autocomplete reduces the changes of incorrectly entering values already present.  The same approach could be used for the date and fair names.

There are two other obvious sources of redundancy: The check amount could also be computed automatically based on the award and the number of students with that project name.  The salutation (Mr. or Ms.) and the pronoun (his or her) are also redundant and could be replaced with a single gender ID field.

### Selective Generation
We may not want to (re)generate all files every time - i.e. we might want to print in batches (some fairs are held much earlier and we don't want the winners to wait), or a mistake might be caught after printing.  This could be accommodated with a dialog that allows the user to select which rows to use.  Alternatively there could be an extra column in the spreadsheet that indicates whether to process the given row, and running the script would populate that column with the date and time the letter was generated.  To reprint a specific letter just remove the entry from the row.

Note that it appears creating a duplicate document with the same name as an existing document does NOT overwrite the original; Google supports multiple files in a directory with the same name.  It's also fair to point out this is mainly a time saver, since it's easy to selectively print only the newly generated files.


error checking?
making label generation more versatile - supporting different format for bulk letters.
it would be nice if we could put all the letters and certs into a single file for easier printing!


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
