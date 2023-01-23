# prefillqr
### Create QR codes of generated Google Forms prefill links

## To Download:
[Click here](https://github.com/IonAir1/prefillqr/releases/tag/release1.1.1) to go to downloads page and select the appropriate version.
For Windows and Linux, executables are available but for MacOS, you will have to use the source code.


## How to use

### Codes and Code Tab

To set up the program, open the Google Forms link, click the **3 dots button** in the top right, then click **Get pre-filled Link**  
For **Short answer text** and **Long answer text** type in a unique **code** (e.g *name1* or *email1*)  
Enter a random year and time for **Year** and **Time** entries. If you have multiple, do not pick the same year and time  
For **Multiple Choice** and **Dropdown** Select any choice. If you are going to prefill multiple of these, you need to select a different choice for each. If you are going to select **Other Option**, simply enter a unique code.  
This program is unable to handle **Checkboxes** and **Grids** so you cannot prefill them.  
Only prefill the answers that need to be prefilled. dont prefill any unnecessary questions.  
After that, scroll down and click **Get Link** then click **Copy Link** and then in the program, in settings, paste it in the **G Forms Prefill Link** Entry  


In the program, open the Codes tab. On the top, the left entry is for your **Codes** and the right is for the columns in the excel file.  
For **Short and Long answer text** codes, type the unique code you gave as is. (e.g **name1** or **email1**)  
For **Year**, Type the year you entered in the format YYYY-MM-DD (e.g 2022-06-23)  
For **Time**, type the time you entered in 24hr format (e.g 17:35)  
For the **Multiple Choice** and **Dropdown**, type the choice you selected with the spaces replaced with **+** (e.g Option+1 or Option+2). If **Other Option** was selected, type the unique code as is.  

For the 2nd entry, you can type the excel column letter of the appropriate answers. You can also use **+** to combine multiple columns, useful for when first names and last names are in a different column. (e.g A or B or AO or A+B or C+A)

If you are satisfied, you can click **Add**  
You can repeat this proccess until you have added all the prefilled entries  

To remove a code, you can select the code you want to remove and click **Remove**  
To remove all, you can click **Select All** then **Remove**

***Important:*** You must have a code named **filename** with any column from the excel file to specify the filename for the qr codes. (e.g filename      B+A)  

### Settings Tab

**Excel File** - Select the excel file containing all the prefill answers  
**G Forms Prefill Link** - Enter the prefill link copied from earlier  
**Use Bitly** - Check to enable bitly. Afterards, the **Bitly Tokens** tab will be enabled  
**Invert Color** - For some devices, the colors might appear inverted. Check to revert the colors back.  
**Filter "=" Sign** - If enabled, Only codes after an equal sign within the link will be detected. Enable for Prefill G Forms link but disable for other uses.  
**Export as CSV** - Will export to a csv file compatible with me-qr.com instead (note that me-qr.com does not accept commas in CSVs. use lower quotes/U+201A/‚ instead).  
**Starting Cell** - Determine the starting point in your excel file. note that the columns entered in **codes** is relative to the starting cell (e.g If starting cell is B1; filename=A will be interpreted as filename=B, so ideally you should only change the row but keep the column as **A** in order to not get confusion)  
**Box Size** - determine the size (in pixels) of each pixel/box in qr code (must be a value of 1-100)  
**Border Size** - determine the border size of the qr codes (must be a value of 1-100)  
**Output Folder** - Enter the folder for the output path of the generated qr codes

### Bitly Tokens Tab

*Bitly can be used to shorten the url and make the qr codes smaller*

In the settings, click the **Use Bitly** checkbox to enable Bitly
Once **Use Bitly** is enabled, the **Bitly Tokens** tab will be enabled.

To get a token, you must open the bitly website and log in. Then go to the [**Bitly Settings Page**](https://app.bitly.com/settings/profile/). On the left of the webpage, click API which is under the Developer settings.  
In the **Access Token** section, enter your password and click generate. You can now select and copy the given token. (note: Do not share your tokens as it can give someone access to your account)

In the program, paste the token in the entry below where it says 'Bitly Tokens' then click add.

You can use a maximum of 5 tokens at the same time. If you are going to be creating more than **50** qr codes in a month and you do not have a premium account, it is recommended to add multiple tokens.  
If tokens are entered, the usage will appear in the bottom left corner. If it exceeds 100% you need new tokens.

As you may have noticed by now, all of the tokens are hidden as asterisks, click the **Show Tokens** checkbox in the bottom left corner to show tokens.

To remove tokens, simply select the tokens you want to remove then click **Remove**.  
To remove all, you can click **Select All** then **Remove**
