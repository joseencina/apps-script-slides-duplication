# Apps Script Slides Duplication

## Overview

This repository contains a basic example of Apps Script code for managing the automatic duplication of a _Google Slides_ file.

The script is designed to simplify the process of creating and sharing a copy of a slide document, ensuring that each copy maintains the formatting and structure

## Apps Script Configuration

To utilize this Apps Script code, you need to meet the following prerequisites:

### Google Slides Folder

Ensure you have an existing folder where Google Slides files are collected. The **Folder Id** will be needed during the logic execution.

> [!NOTE]
> This logic analyzes the existing files in the specified folder and uses the most recently created file to make a copy.

### Apps Script Add-on

Add the Apps Script to your Google Slides file.
   - Open your Google Slides file.
   - Click on `Extensions` in the menu bar.
   - Select `Apps Script`. This will open the Apps Script editor.
### Build Code

Insert (and edit) Apps Script code at the editor.

Save the script by clicking on the disk icon or pressing `Ctrl+S`.

> [!IMPORTANT]
> Google Slides and Drive files will ask for authorization at this step.

### Triggers

Add a trigger to run automatically at specific intervals or events.
   - In the Apps Script editor, click on the clock icon to open the `Triggers` menu.
   - Click on `+ Add Trigger` at the bottom to configure a trigger.
   - Select main function to be triggered*.
   - Select the time or the frequency this code will be triggered.
   - Save the trigger settings.

> [!NOTE]
> This code is using the main function `duplicate()` to be executed from the trigger configuration.

## Apps Script Code

### Data Configuration Resources

This object holds various configurations and data structures used throughout the script. The most important information to be added is email and folder id.

```javascript
let dataConfig = {
    email: 'team-email@example.com',
    folder: {
        id: 'team_folder_id',
        // (...)
    },
    // Additional resources
}
```

> [!WARNING]
> Please refer to the complete `dataConfig` object in the main `slides-duplication.js` file attached to this repository to ensure all resources are taken into account.

### Main Logic Functionalities

This section contains the main logic of the script. As described, this code takes into account the folder id where a the file that is required for the copy is located.

#### `init()`
Initializes the script with the data configuration, sets up necessary data and execute the different methods in order.
```javascript
init: function(d) {
    _main.data = d;

    _main.data.date.nextDate = _main.getNextDate(_main.data.date.additional);
    _main.data.date.formatted = Utilities.formatDate(_main.data.date.nextDate, 'ETC/GMT', 'yyyy-MM-dd');
    
    // Rest of the methods in order
  },
```
#### `getNextDate()`
Calculates the next date based on the current date and an additional number of days.
```javascript
getNextDate: function(addDays) {
    const date = new Date();
    date.setDate(date.getDate() + addDays);
    return date;
}
```
#### `getOrdinal()`
Returns the ordinal suffix for a given date.
```javascript
getOrdinal(date) {
  let d = date.getDate(), suffix = ['th', 'st', 'nd', 'rd'][(d > 3 && d < 21) || d % 10 > 3 ? 0 : d % 10];
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "d'" + suffix + "'");
}
```
#### `getLatestFile()`
Retrieves the latest created file from the specified folder.
```javascript
getLatestFile: function(folderContent) {
    _main.data.files = folderContent.getFiles();
    let tempLatestFile = null;
    while(_main.data.files.hasNext()) {
        _main.data.file.content = _main.data.files.next();
        if(_main.data.file.latestDate == null || _main.data.file.content.getDateCreated() > _main.data.file.latestDate) {
        _main.data.file.latestDate = _main.data.file.content.getDateCreated();
            tempLatestFile = _main.data.file.content;
        }
    }
    return tempLatestFile;
}
```
> [!NOTE]
> This is the file that will be used as template for the copy.

#### `getCopy()`
Creates and returns a copy of the latest file.
```javascript
getCopy: function(latest, name) {
    return latest.makeCopy(name);
}

```
#### `setDateOnCover()`
Updates the date on the first slide of the copied presentation.

```javascript
setDateOnCover: function(fileId, newText) {
    let newSlide = SlidesApp.openById(fileId).getSlides()[0];
    let text = newSlide.getPageElements()[2].asShape().getText();
    text.setText(newText);
}
```
> [!TIP]
> This method is used as an example of what is possible to be edited inside the copied file.
#### `sendEmail()`
Configures and sends an email notification with a link to the newly created copy.
```javascript
sendEmail: function(date,fileId) {
    _main.data.emailConfig = _main.emailConfig(
        _main.data.email,
        `[REMINDER] Team Meeting ${date}`,
        `Hello, Team!<br>This is a reminder for your next Team Meeting.<br><br>Please fill in your notes and update your slide in the document:<br>https://docs.google.com/presentation/d/${fileId}/ <br><br>Previous files are available here:<br>https://drive.google.com/drive/folders/${getFolderConfig()}<br><br>Enjoy your day &#x1F9E1;!<br><span style="font-weight: 500;font-size: 11px;">This is an automatic message &#x1F916;.</span>`
    );
    GmailApp.sendEmail(_main.data.emailConfig.add, _main.data.emailConfig.sub, "", { htmlBody: _main.data.emailConfig.body });
}
```

>[!IMPORTANT]
> `emailConfig()` method is added as an auxiliar function on the complete code for email configuration to be built easily.


### External Function Initiator

The `duplicate()` function is the entry point for Apps Script configuration to be able to set-up a trigger.

```javascript
function duplicate() {
    _main.init(dataConfig);
}
```

It is possible to make this function to be triggered bi-weekly by adding a condition on the week number.

```javascript
function duplicate() {
  let weekNumber = Utilities.formatDate(new Date(), "GMT", "w") -1;
  if(weekNumber % 2 == 0) {
    _main.init(dataConfig);
  }
}
```
> [!NOTE]
> Google Apps Script triggers are only configurable by Weekly basics. It is possible to configure a trigger driven by calendar event too.
