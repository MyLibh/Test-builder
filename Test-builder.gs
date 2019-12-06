var TASKS_FOLDER_NAME = "Tasks";
var EXCLUDED_FOLDERS;
var EXCLUDED_FILES =
{
//  "ThemeFolder": ["file1.txt", "file2.pdf"]
    "Theme1": [],
    "Theme2": ["2_2.txt"],
    "Theme3": []
};

var THEME_FREQUENCY =
{
//  "ThemeFolder": frequency
    "Theme1": 1,
    "Theme2": 1,
    "Theme3": 2
};

var NUMBER_OF_VARIANTS = 2;
var VARIANTS_FOLDER_NAME = "Result";

var FORM_TITLE = "FormName";

function add_variant(form, page_break, index)
{
    page_break.setTitle("Вариант " + (index + 1));

    // do stuff
    form.addTextItem().setTitle("Задача").setRequired(true);

    return form.addPageBreakItem()
               .setGoToPage(FormApp.PageNavigationType.SUBMIT);
}

function add_variant_chooser(form, variant_pages)
{
    var page_break = variant_pages[variant_pages.length - 1].setTitle("Выбор варианта");

    var variant_chooser = form.addListItem()
                              .setTitle("Номер варианта")
                              .setRequired(true);
    var choices = [];
    for (var i = 0; i < NUMBER_OF_VARIANTS; ++i)
        choices.push(variant_chooser.createChoice(String(i + 1), variant_pages[i]));

    variant_chooser.setChoices(choices);

    form.addPageBreakItem();

    return page_break;
}

function add_general_questions(form)
{
    var email_validation = FormApp.createTextValidation()
                                  .setHelpText("Неправильный формат почты")
                                  .requireTextIsEmail()
                                  .build();
    form.addTextItem()
        .setTitle("Email")
        .setRequired(true)
        .setValidation(email_validation);

    form.addTextItem()
        .setTitle("Фамилия")
        .setRequired(true);

    form.addTextItem()
        .setTitle("Имя")
        .setRequired(true);

    form.addListItem()
        .setTitle("Класс")
        .setRequired(true)
        .setChoiceValues([7, 8, 9, 10]);

    return form.addPageBreakItem();
}

function create_form()
{
    var forms = DriveApp.getRootFolder().getFilesByName(FORM_TITLE);
    var form = forms.hasNext() ? FormApp.openById(forms.next().getId()) : FormApp.create(FORM_TITLE);
    form.setTitle("Тест")
        .setAllowResponseEdits(false)
        .setAcceptingResponses(true)
        .setShowLinkToRespondAgain(false)
        .setConfirmationMessage("Тестирование закончено")
        .setCustomClosedFormMessage("Тестирование окончено")
        // .setIsQuiz(true)
        // .setDestination();

    var items = form.getItems();
    for (var item in items)
        form.deleteItem(items[item]); // TODO: Check

    var general_page = add_general_questions(form);
    var variant_pages = [];
    var prev_page = general_page;
    for (var i = 0; i < NUMBER_OF_VARIANTS; ++i)
    {
        prev_page = add_variant(form, prev_page, i);

        variant_pages.push(prev_page);
    }

    general_page.setGoToPage(add_variant_chooser(form, variant_pages));

    return form.shortenFormUrl(form.getPublishedUrl());
}

function get_random(array, n)
{
    if (n > array.length)
        throw new RangeError("get_random: More elements taken than available");

    if (n == 1)
        return [array[Math.floor(Math.random() * array.length)]];

    var shuffled = array.sort(function (a, b) { return 0.5 - Math.random(); });

    return shuffled.slice(0, n);
}

function generate()
{
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName("result").clearContents();
    var this_file_id = ss.getId();
    var this_file = DriveApp.getFileById(this_file_id);
    var parents = this_file.getParents();
    var this_folder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
    var res_folders = this_folder.getFoldersByName(VARIANTS_FOLDER_NAME);
    var res_folder = res_folders.hasNext() ?
        this_folder.removeFolder(res_folders.next()).createFolder(VARIANTS_FOLDER_NAME) :
        this_folder.createFolder(VARIANTS_FOLDER_NAME);
    var tasks_folder = this_folder.getFoldersByName(TASKS_FOLDER_NAME).next();

    var set = load_data();
    for (var i = 0, j = 1; i < NUMBER_OF_VARIANTS; ++i, j = 1)
    {
        var var_folder_name = "Variant" + (i + 1);
        var var_folders = res_folder.getFoldersByName(var_folder_name)
        var var_folder = var_folders.hasNext() ?
          res_folder.removeFolder(var_folders.next()).createFolder(var_folder_name) :
          res_folder.createFolder(var_folder_name);

        sheet.getRange(String.fromCharCode(65 + i) + j++).setValue(var_folder_name);
        for (var theme in set)
        {
            var tasks = set[theme];
            var freq = THEME_FREQUENCY[theme];
            var res = get_random(tasks, freq);

            for (var task in res)
                sheet.getRange(String.fromCharCode(65 + i) + j++).setValue(res[task]);

            var theme_folder = tasks_folder.getFoldersByName(theme).next();
            for (var task_file in res)
                theme_folder.getFilesByName(res[task_file]).next().makeCopy(var_folder);
        }
    }

    sheet.activate();
}

function load_data()
{
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName("data").clearContents();

    var this_file_id = ss.getId();
    var this_file = DriveApp.getFileById(this_file_id);
    var parents = this_file.getParents();
    var this_folder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
    var tasks_folder = this_folder.getFoldersByName(TASKS_FOLDER_NAME).next();
    var theme_folders = tasks_folder.getFolders();

    var column = 65;
    var tasks_set = {};
    while (theme_folders.hasNext())
    {
        var theme = theme_folders.next();
        var theme_name = theme.getName();
        if (EXCLUDED_FOLDERS.indexOf(theme_name) != -1)
            continue;

        tasks_set[theme_name] = [];

        var row = 1;
        sheet.getRange(String.fromCharCode(column) + row).setValue(theme.getName());
        row++;

        var tasks = theme.getFiles();
        while (tasks.hasNext())
        {
            var task = tasks.next();
            var task_name = task.getName();
            if (EXCLUDED_FILES[theme.getName()].indexOf(task_name) != -1)
                continue;

            tasks_set[theme_name].push(task_name);

            sheet.getRange(String.fromCharCode(column) + row).setValue(task.getName());
            row++;
        }

        column++;
    }

    sheet.activate();

    return tasks_set;
}
