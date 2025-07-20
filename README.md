# App Script - Insert Thumbnails from GDrive into Google Sheet

This Google Apps Script allows users to insert images from a specified Google Drive folder into a Google Sheet.

###  Features

- **Preview Sort Order**: Before inserting images, you can preview the order in which files will be sorted.
- **Batch Processing**: The script processes images in batches to handle large numbers of files efficiently, 10 images at a time.
- **Progress Tracking**: A checklist shows the current stage of the process.
- **Resizing**: Images are automatically resized to maintain quality while keeping file sizes manageable.

<br>
<img src="./Docs/Images/Main UI.png" width="450" />

## Installation

### Step 1: Create a Google Sheet

1. Open [Google Sheets](https://sheets.google.com).
2. Create a new Google Sheet or open an existing one.

### Step 2: Open the Script Editor

1. In your Google Sheet, click on `Extensions` in the menu bar.
2. Select `Apps Script` from the dropdown menu. This will open the Apps Script editor in a new tab.

### Step 3: Add the Main Code

1. In the Apps Script editor, delete any code in the `Code.gs` file under the `Files` section.
2. Copy and paste the code from the [Code.gs](Code.gs) file in this repo into the `Code.gs` file.
3. Save the script by clicking the disk icon or by pressing `Ctrl + S` (`Cmd + S` on Mac).

### Step 4: Create the HTML UI File

1. In the Apps Script editor, click on the `+` button next to the `Files` section.
2. Select `HTML` and name the file "`index`" (must be lowercase, do not include an extension).
3. Delete all the automatically created code in the `index.html` file.
4. Copy and paste the code from the [index.html](Index.html) file in this repo into the `index.html` file in Google App Scripts.
5. Save the HTML file by clicking the disk icon or by pressing `Ctrl + S` (`Cmd + S` on Mac).

### Step 5: Add the Manifest File

**Important:** This step is required to allow the app permission to read your Google Drive images and edit the Google Sheet.

1. In the Apps Script editor, go to `Project Settings` (gear icon on the left sidebar).
2. Check the box for `Show 'appsscript.json' manifest file in editor`.
3. Go back to the `Editor` tab and you should now see an `appsscript.json` file in the file list.
4. Click on the `appsscript.json` file and replace its contents with the code from the [appsscript.json](appsscript.json) file in this repo.
4. Not necessary, but if you wish to change the timezone you can find the correct timezone code [here](https://en.wikipedia.org/wiki/List_of_tz_database_time_zones#List).
6. Save the manifest file by clicking the disk icon or by pressing `Ctrl + S` (`Cmd + S` on Mac).

### Step 6: Run the Add-on
1. Go back to your Google Sheet, you can close the Apps Script tab if you wish.
2. Refresh the web page by clicking the refresh icon or by pressing `Ctrl + R` (`Cmd + R` on Mac).
3. Click on the new `Custom Tools` button in the menu bar (it may take a few seconds to appear).
4. Select `Insert Thumbnails from GDrive`.
 - On first run, in that Sheet, you may be prompted to authorise access to your Google Drive by the script (see [Security Prompt](#Security-Prompt)).
5. In the "Folder URL" field, enter the URL of the Google Drive folder containing the images.
 - For example: `https://drive.google.com/drive/folders/1bRNBR55227gdQc-3_jhfsef_5R80RjAt` < Not a real URL
6. In the "Start Row" and "Start Column" fields, enter the starting row and column for the images to be inserted.
7. Click `Insert Images` to begin the process.
- Once the script starts running, you'll be able to press Stop to stop the script.
- Additionally as the script runs it'll show a checklist of progress under the Stop button.

**NOTE: The script will NOT stop running if you close the window, only if you press the Stop button.**

## Security Prompt
#### On first run, in that Sheet, you may be prompted to authorise access to your Google Drive by the script.
Up to you if you feel like the security is too permissive, the full code is here to see and audit.

| 1. Click Show Advanced<br>2. Click "Go to Insert Thumbnails from GDrive" or "Untitled  | <img src="./Docs/Images/SecPrompt_Step1.png" alt="Screenshot of Google Drive Authorisation - Step 1" width="400"> |
| -------- | ------- |
| This page explains that the script will need access to the following:<br>1. **Google Drive**<br>To access the folder, and **ONLY** the folder, containing the images<br><br>2. **Google Sheets**<br>To insert the images into the sheet, and **ONLY** the sheet you run the script in<br><br>3. **"Allow this app to run when you are not present"**<br>is needed because the process for inserting images can take a long time<br>Potentially 10+ minutes for hundreds of images.<br>So the tool runs the process in the background and in batches.<br><br>**To Continue**<br>1. Make sure to check "Select All"<br>2. Click Continue | <img src="./Docs/Images/SecPrompt_Step2.png" alt="Screenshot of Google Drive Authorisation - Step 2" width="400"> |

## Troubleshooting

### OAuth Scope Errors

If you encounter error messages about insufficient permissions or invalid scopes, such as:

```
https://www.googleapis.com/auth/script.scriptapp
You are receiving this error either because your input OAuth2 scope name is invalid or it refers to a newer scope that is outside the domain of this legacy API.
```

Or:

```
Exception: Specified permissions are not sufficient to call Ui.showModalDialog. Required permissions: `https://www.googleapis.com/auth/script.container.ui`
```

**Solution:** Make sure you've completed **Step 5** above (adding the manifest file). These errors occur because Google has updated their OAuth requirements and the script needs explicit permission declarations in the `appsscript.json` manifest file.

### Other Issues

If you encounter any other issues:
1. Make sure you have the necessary permissions for both the Google Sheet and the Google Drive folder.
2. Check that the folder URL is correct and accessible.
3. If the script stops unexpectedly, try running it again. There may be temporary issues with Google's services.

## License

This project is licensed under the GPLv3 License - see the [LICENSE](LICENSE) file for details.
