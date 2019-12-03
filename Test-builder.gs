var TASKS_FOLDER_NAME = "Tasks";
var EXCLUDED_FOLDERS = [];
var EXCLUDED_FILES = [];

function load_data()
{
    var ss = SpreadsheetApp.getActive();
    var this_file_id = ss.getId();
    var this_file = DriveApp.getFileById(this_file_id);
    var parents = this_file.getParents();
    var this_folder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
    var tasks_folder = this_folder.getFoldersByName(TASKS_FOLDER_NAME).next();
    var theme_folders = tasks_folder.getFolders();

    var column = 65;
    while (theme_folders.hasNext())
    {
        var theme = theme_folders.next();
        if (EXCLUDED_FOLDERS.indexOf(theme.getName()) != -1)
            continue;

        var row = 1;
        ss.getRange(String.fromCharCode(column) + row).setValue(theme.getName());
        row++;

        var tasks = theme.getFiles();
        while(tasks.hasNext())
        {
            var task = tasks.next();

            ss.getRange(column + row).setValue(task.getName());
            row++;
        }

        column++;
    }
}
