/**
 * Data Configuration Resources
 */
let dataConfig = {
    email: 'team-email@example.com',
    folder: {
        id: 'team_folder_id',
        content: null
    },
    file: {
        content: null,
        latest: null,
        latestDate: null
    },
    copy: {
        id: null,
        content: null,
        newName: '',
        coverText: ''
    },
    files: null,
    date: {
        additional: 7,
        nextDate: null,
        formatted: null,
        ordinal: null,
        obj: {
        y: null,
        m: null,
        d: null
        }
    },
    monthName: {
        "00": "NA",
        "01": "Jan",
        "02": "Feb",
        "03": "Mar",
        "04": "Apr",
        "05": "May",
        "06": "Jun",
        "07": "Jul",
        "08": "Aug",
        "09": "Sep",
        "10": "Oct",
        "11": "Nov",
        "12": "Dec"
    }
}

/**
 * Main Logic Functionalities
 */
let _main = {
    init: function(d) {
        // Save data configuration object
        _main.data = d;

        // Set and get new date
        _main.data.date.nextDate = _main.getNextDate(_main.data.date.additional);
        _main.data.date.formatted = Utilities.formatDate(_main.data.date.nextDate, 'ETC/GMT', 'yyyy-MM-dd');

        // Set date object for functionality usage
        _main.data.date.obj.y = _main.data.date.formatted.split('-')[0];
        _main.data.date.obj.m = _main.data.date.formatted.split('-')[1];
        _main.data.date.obj.d = _main.data.date.formatted.split('-')[2];

        // Set cover new date reference with -> Month name abbrevation + Ordinal number of the current day + Year
        _main.data.copy.coverText = `${_main.data.monthName[_main.data.date.obj.m]} ${_main.getOrdinal(_main.data.date.nextDate)} ${_main.data.date.obj.y}`;
        
        // Set new file name
        _main.data.copy.newName = `${_main.data.date.formatted} - Team Weekly`;
        
        // Get all the folder files
        _main.data.folder.content = DriveApp.getFolderById(_main.data.folder.id);
        
        // Get the latest file created available
        _main.data.file.latest = _main.getLatestFile(_main.data.folder.content);

        // Set and get the copy of the latest file available
        _main.data.copy.content = _main.getCopy(_main.data.file.latest, _main.data.copy.newName);
        // Set copy file id
        _main.data.copy.id = _main.data.copy.content.getId();

        // Change data on the first page
        _main.setDateOnCover(_main.data.copy.id, _main.data.copy.coverText);

        // Send reminder after file creation
        _main.sendEmail(_main.data.date.formatted, _main.data.copy.id);
    },

    // Set new Date for the name
    getNextDate: function(addDays) {
        // Create current date
        const date = new Date();
        // Set current day + additional days date
        date.setDate(date.getDate() + addDays);
        // Get formatting of the date
        return date;
    },

    // Set and get ordinal number from current day for cover
    getOrdinal(date) {
        let d = date.getDate(), suffix = ['th', 'st', 'nd', 'rd'][(d > 3 && d < 21) || d % 10 > 3 ? 0 : d % 10];
        return Utilities.formatDate(date, Session.getScriptTimeZone(), "d'" + suffix + "'");
    }, 

    // Get the last created file inside the folder
    getLatestFile: function(folderContent) {
        // Get all the files available at the folder
        _main.data.files = folderContent.getFiles();

        // Set null variable for saving latest file
        let tempLatestFile = null;

        // Check all the files contained at the folder
        while(_main.data.files.hasNext()) {
        _main.data.file.content = _main.data.files.next();
        if(_main.data.file.latestDate == null || _main.data.file.content.getDateCreated() > _main.data.file.latestDate) {
            _main.data.file.latestDate = _main.data.file.content.getDateCreated();
            tempLatestFile = _main.data.file.content;
        }
        }
        return tempLatestFile;
    },

    // Create and return the copy of the latest file available
    getCopy: function(latest, name) {
        return latest.makeCopy(name);
    },

    // Set new date on first page
    setDateOnCover: function(fileId, newText) {
        // Get first slide
        let newSlide = SlidesApp.openById(fileId).getSlides()[0];
        // Get text containing the date info
        let text = newSlide.getPageElements()[2].asShape().getText();
        text.setText(newText);
    },

    // Configure and send email
    sendEmail: function(date,fileId) {
        // Configure the email object for easier use during the email sending
        main.data.emailConfig = main.reminder.emailConfig(
            _main.data.email,
            `[REMINDER] Team Meeting ${date}`,
            `Hello, Team!<br>This is a reminder for your next Team Meeting.<br><br>Please fill in your notes and update your slide in the document:<br>https://docs.google.com/presentation/d/${fileId}/ <br><br>Previous files are avialable here:<br>https://drive.google.com/drive/folders/${getFolderConfig()}<br><br>Enjoy your day &#x1F9E1;!<br><span style="font-weight: 500;font-size: 11px;">This is an automatic message &#x1F916;.</span>`
        );
        // Send email for each contact
        GmailApp.sendEmail(main.data.emailConfig.add, main.data.emailConfig.sub, "", { htmlBody: main.data.emailConfig.body });
    },
    // Aux email configuration object for email structure
    emailConfig: function(add,sub,body) {
        return { "add": add, "sub": sub, "body": body }
    },

    // Aux data object 
    data: null
}

/**
 * External Function Initiator
 */
function duplicate() {
    _main.init(dataConfig);
}
