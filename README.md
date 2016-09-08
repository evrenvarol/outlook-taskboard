# Outlook Taskboard
Outlook Taskboard is a kanban board style view for Outlook Tasks.

It uses the main "Tasks" folder as *Back Log* and utilizes 4 individual subfolders (InProgress, Next, Focus and Waiting) as each task lanes for kanban workflow.

There are 2 ways to use the taskboard.

  1. Outlook Folder Home Page (recommended)
  2. Directly from Internet Explorer

### Screenshots

![Outlook Taskboard] (http://evrenvarol.github.io/outlook-taskboard/img/outlook-taskboard.png)

### Platforms supported
Only tested with Outlook 2013 and 2016 running on Windows 8.1/10.

It may also work with earlier Outlook versions, and possibly work with Windows 7.

The taskboard can also be opened in Internet Explorer. At the moment, only IE9 / IE10 and IE11 are supported.

### How to setup

1. Download the latest release zip file and extract it to a folder in your local hard drive.

2. In Outlook, create following folder structure under your Tasks folder:

    ![Tasks Folders] (http://evrenvarol.github.io/outlook-taskboard/img/task-folders.png)

3. For Outlook Home page:

  * Create another folder (of any type) and name it something like "Taskboard" or "Kanban", etc. (Alternatively you can use the main account folder as a home page as well)

  * Right-click the folder, and then click **Properties**. Select the *Home Page* tab in the <folder name> Properties dialog box.

  * In the *Address box*, browse to the folder you have just extracted the Taskboard files and select the **kanban.html** file.

  * Click to select the *Show home page by default for this folder* check box and then click **OK**.

      ![Folder Home Page Offline Warning] (http://evrenvarol.github.io/outlook-taskboard/img/folder-home-page-offline-warning.png)

      <sub>*If you receive above warning, simply click X icon to close both warning prompt and the Properties window.*</sub>

4. For Internet Explorer:

  * Open Internet Explorer and go to *Tools > Internet Options > Security tab*. Select the **Local Intranet Zone** and click on the **Custom Level** button. Ensure the "Initialize and script ActiveX controls not marked as safe for scripting" option is set to **Enabled**

  ![IE Local Intranet Zone Setting] (http://evrenvarol.github.io/outlook-taskboard/img/ie-localintranet-activexscript-enable.png)

  * Double-click on the **kanban-ie.html** file to open the page in Internet Explorer.

    <sub>*On Win10, you will need to right click on the file, select Open With -> Internet Explorer to open the page in IE11. Otherwise it tries to open in Edge which is not supported.</sub>








