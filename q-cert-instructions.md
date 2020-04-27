# Q Certificate PDF Creation and Mailing Automation
Hey all, here are the instructions for automatically generating the PDFs and emails to distribute for the Q awards. 


Note: You will want all docs, spreadsheets, and folders to be owned by the account you'll be emailing from.


## Step 1: Format the spreadsheet
Open up the Google Spreadsheet you're using to track responses from awardees, it will look like the first sheet of [this spreadsheet](https://docs.google.com/spreadsheets/d/1ZiBLvL39-youUb3qSK-hub4-NQ_wL0HJeKWa7u753ek/edit#gid=1515223909). 
Start by duplicating the form responses onto the next sheet, for safety's sake.
Next, change the headers in the top row to match the following:
| Old Heading | New Heading |
| ------ | ------ |
| Email address | EMAIL |
| First Name | FIRSTNAME |
| Your name to print on the certificate:  We copy this field onto your certificate, so please specify exactly how you would like your name to appear: | CERTNAME |
>Note: We're doing this so that the code knows where to look for the `EMAIL` address to send the certificate to, the `FIRSTNAME` of the person to be addressed in the email, and the name to write on the certificate (`CERTNAME`).

You will also want to make a few new columns:
| New Column | Formula to put in it | Purpose |
| ------ | ------ | ------ |
| String Length | `=arrayformula(LEN(F2:F))` | This counts the number of characters in the name they want to be used on their certificate. This isn't necessary for the code, but is useful for checking very long names to make sure the certificate looks good before sending. |
| FILENAME | `=arrayformula(D2:D&"_"&LEFT(C2:C,1))` | This creates the filename, which should be formatted as [lastname]_[first initial]. |
| PDF | (Leave this empty.) | This is where the link to the pdf will go when the `createPDF` script is run.|
When you're done, it should look like Sheet 2 of [this spreadsheet](https://docs.google.com/spreadsheets/d/1ZiBLvL39-youUb3qSK-hub4-NQ_wL0HJeKWa7u753ek/edit#gid=380205012).


## Step 2: Gather your materials
You will need to have the following materials ready to go:
1. The spreadsheet you were working on above.
2. The award template (in Google Docs). [Here's an example.](https://drive.google.com/open?id=1SD_EFNC-E-yjYl_MEnxCeVVa0Br4P16vdxPRbGXbxnM)
3. The email template (in Google Docs) formatted *exactly* the way you want the email to be formatted. [Here's an example.](https://drive.google.com/open?id=1K-ta6OEBmQtEM1nqHelCJC6k5pCx6pIp8CiHmQfR4Oo)
4. Create a folder where the finished PDF certificates should live.

## Step 3: Code!
### Paste your code
Once you have your documents and folders open, go to the LL Dev Github and copy the code that Lauren wrote [here](https://github.com/learninglab-dev/ll-apps-scripts/commit/d1e0ab3e9169ea345795aca9bcd6d19dcffe3657).

Go back to your spreadsheet and click `Tools` > `Script Editor`. This will open a window with the beginning address `script.google.com`. Start by deleting anything that's written in the box on that page. Then, paste Lauren's code in the box and click the `Save` icon. You can save it as `Q code` or whatever title you'd like. 

The only part of the code you will need to interact with is the very top:
```sh
// Replace with your document IDs before running the script:
// 1. Create a GDoc certificate template and put the ID here
var TEMPLATE_ID = '1SD_EFNC-E-yjYl_MEnxCeVVa0Br4P16vdxPRbGXbxnM'
// 2. Create a folder for the certificates and put the ID here
var CERTIFICATES_FOLDER_ID = '1tqW7wgzHEBt1dYyxJHRN7fgc5opjbokH'
// 3. Put the ID for the GDoc with the email body here and add a subject line
var EMAIL_TEMPLATE_ID = '1K-ta6OEBmQtEM1nqHelCJC6k5pCx6pIp8CiHmQfR4Oo'
var EMAIL_SUBJECT = 'Q Award Certificate'
```
### Find your document IDs
The code above is asking for the IDs of the templates that you gathered. To find those IDs, you'll need to look in the URLs of each template.
If the whole URL looks like this: `https://docs.google.com/document/d/1SD_EFNC-E-yjYl_MEnxCeVVa0Br4P16vdxPRbGXbxnM/edit`
The ID is the string of letters and numbers between `/d/` and `/edit`. So in that example this ID is `1SD_EFNC-E-yjYl_MEnxCeVVa0Br4P16vdxPRbGXbxnM`.

### Paste in your document IDs
`// 1.` in the code is looking for the ID of the certificate template.
`// 2.` in the code is looking for the ID of the folder where the PDFs of the certificates should be saved.
`// 3.` in the code is looking for the ID of the email template.
You can also change the `var EMAIL_SUBJECT` to whatever you'd like. That is the text that will show up in the subject line of your email.

### Save it!
Save your work by pressing the `Save` button again.

### Run the code (part 1)
To the right of the little bug icon in the file menu of the coding screen there will be a dropdown menu to select which section of code to run. Click that dropdown and select `onOpen`, then press the `Play` triangle button to the left of the little bug icon.
By running this section of code, you're creating a little menu for yourself on the main Google spreadsheet.
Once it's run, you can exit out of the coding window.

## Step 4: Create the PDFs
The script is programmed to run on only the rows that you select. Before you go crazy selecting all of the rows and creating thousands of PDFs, let's start with just one test:
* Add a row at the top of your spreadsheet with your name and email address (make sure that your filename and string length formulas didn't break).
* Highlight the row with your info.
* In the File menu, scroll over to the menu item called `Robot Stuff` and click `create PDF`.
* A green box should appear at the top of your screen that says `Running script`. Wait a second :) good things take time.
* When it's done it will say `Finished script` and the PDF column should fill with a link.
* Check your test pdf out for errors.

Once you've tested it out on one, feel free to generate more by highlighting the columns and going back up to `Robot Stuff` > `create PDF`. Note: This script takes some time to run. You can watch its progress as the PDF boxes fill in and feel satisfied that your robot helper is doing all of the work. :)

**Remember to check the PDFs with particularly long names, just to make sure they haven't broken the template. You might want to highlight the `String Length` column and make a conditional formating rule to color any value over 40.**

## Step 5: Email the recipients
Once again, the script is programmed to run on only the rows that you select. Test out the script on the same tester row you created in Step 4:
* Highlight the tester row
* In the File menu, scroll over to the menu item called `Robot Stuff` and click `send PDF`.
* A green box should appear at the top of your screen that says Running script. Wait a second :) good things take time.
* When it’s done it will say Finished script. You should recieve an email within a few minutes. You can also check your sent mail from the inbox of the account you're doing this with.
* Check your test email for errors.

Once you've tested it out, feel free to send out the rest of your emails. 













 
