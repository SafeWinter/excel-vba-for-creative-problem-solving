# Grade Manager - Problem Statement

Problem Statement

In this project, you will create a tool that will consolidate and combine student grades from multiple class section roster grade books.  The class is split up into multiple sections, with each teaching assistant (TA) responsible for the grades of his or her section.  It is quite a challenge to sync all the grades from the TAs into one master grades spreadsheet.  Obviously, many of you will not be using this for teaching, but this tool can easily be modified to consolidate information from multiple users or sources of data, given a consistent starting format.  Furthermore, when a new column (i.e., grade) is added to one of the individual sheets, that change is made to all sheets in the directory.



![img](../assets/5-1.png)

Overview of Grade Manager project.  The Grade Manager will take section files (File 1 through File 5 in this example) and consolidate/sync the information from those files into the main file (main grade sheet).  Furthermore, the Grade Manager should be able to automatically add a column (grade) or delete a column from all files.  See introductory screencast for all the capabilities and requirements.

The directory that the user creates can be a cloud-based one (e.g., Dropbox folder), making this a powerful tool for consolidating information from multiple users.

For this project, you will create a user form that has several functionalities. 

1) Initialization
2) Synchronization of section files with main grade book
3) A tool to back up the current file, amended with today’s date in the file name
4) The ability to add a new column to or delete an entire column from all grade sheets
5) A search tool that will enable the user to search for a student’s grades on an assignment OR allow the user to modify a grade on an assignment



![img](../assets/5-2.png)





![img](../assets/5-3.png)



First, if the course is new and has not yet been initialized, a directory for the course needs to be set up and the class roster file separated into section roster files (“section files”).  For example, a course might have 5 recitation/breakout sections (101 to 105), so your program needs to create a new .xlsx file for and separate students from the main roster file into each of those 5 sections (therefore, there will be 5 individual section files).

Second (a continuation of the initialization process), the user will be able to select the number of assignments in the course from drop-down boxes (combo boxes).  For example, the user can decide that there will be 10 homework assignments, 2 exams, and 4 labs.  Don’t worry, later on the user can delete or add assignments if they wish.

Third, your program/user form should have the ability to synchronize (“sync”) the information found in those 5 section files to the main grade file.  The individual TAs will make changes to their section files but periodically you need to sync the information to update the main grade sheet.  In this way, the grades from multiple sheets are combined into a single sheet.

Fourth, your project should have a way to make a backup file of the current file, saving it with the current date in the file name.

Fifth, there should be a way for the user to add or delete entire grade columns, and these changes would automatically be made to ALL individual section files.  

Finally, your user form should have a way to search through by student name (selected via a drop-down list) for a certain grade on an assignment (selected from a drop-down list), and this would be displayed in a message box.  The user should also be able to change a grade in the spreadsheet.  All changes would be permanent (i.e. placed in the actual cells on the spreadsheet).

On the course website I provide many hints for the individual steps along the way.



This is a very challenging and open-ended project!

This is by far the most challenging project available to you.  Because of the online nature of this course, it will be difficult for you to get help with this project.  If you like to work on challenging, open-ended projects and like to solve problems on your own, this project is well-suited for you.  You will likely need to refer back to previous parts of the course to review various concepts.  You will likely be frustrated at times and “spinning your wheels in the mud.”  I encourage you to reach out to others on the discussion forums to both share some thoughts and to obtain help from others.

On the positive side, this project resembles a real-world project, one in which you must create, synthesize, and design something from the ground up.  This will provide you with a ton of relevant, career-related experience that employers will greatly value.  Learners who complete this project successfully should be very proud with what they accomplish!  Very challenging, yet very rewarding!



Helpful Files

Available on the course site are the following files:

- “Grade Manager – STARTER.xlsm” à just a starting point!
- “Roster.xlsm” à a sample roster with students in sections 101-105 that you can use for your project

These files (in alphabetical order) are products of many of the screencasts presented on the course website; feel free to use them as reference and/or as starting material for the various steps along the way:

- “Add Column Files.zip” à compressed folder with example files
- “AddEntireColumn.xlsm”
- “CreateSections.xlsm”
- “ImportRoster.xlsm”
- “InitializeSheets.xlsm”
- “MakeBackup.xlsm”
- “NewFolder.xlsm”
- “SearchTool.xlsm”
- “Section Files.zip” à compressed folder with example files
- “SyncingFiles.xlsm”

 

Assumptions/Premises

- User will always place names and SID in first 2 columns, and we’ll assume that those will never be misspelled.
- New assignment headings/labels will always be placed in the top row.
- TAs will be able to modify their individual sheets and these will overwrite existing data.
- Lab sections will always be in increments of 1 and continuous (no gaps).
- Grades will always be on Sheet1 of the roster files
- No new students will need to be added



More details

1. (INITIALIZATION) From the START PAGE, clicking on the button will create a new course, which involves the following (see introductory screencast for a demonstration):

- You should create a file called “Grade Manager – SETUP.xlsm” that will perform the tasks below.  There are only 2 files that the teacher needs to set up the course: 1) “Grade Manager – SETUP.xlsm” and 2) the course roster.  These are the two files that your reviewers (other learners) will start with when they peer review your project.
- Ask the user for the name of the course (e.g., “ECON 1001”) and creates a new directory (folder) for this course
- Asks the user to navigate to and select the preexisting folder where the new directory will be placed
- Saves the file as something like “Grade Manager ECON 1001.xlsm” (NOTE: The “Grade Manager – SETUP.xlsm” file should therefore be unchanged, and can be used again for a different course.)
- The START PAGE should be deleted (record a macro for how this is done, you may want to use Application.DisplayAlerts = False)
- The file path should be placed on the “Path” sheet of the main file, and this sheet should be hidden (you can start with it hidden; the subroutine does not need to hide it as long as it starts hidden)
- Asks the user to navigate to the roster file (in the correct format) for the course, which will then be imported into the “Roster” sheet of the Grade Manager.
- Individual section files will then be created in a sub directory of the directory in which the Grade Manager file exists (these files might exist in a folder named something like “Section Rosters” or “Section Files”); the students are separated by section number into the section files.
- IMPORTANT: Both the A and B columns (name and SID numbers) should be imported into the section grade files!
- You will need to do some tweaking around with the numbers in the CreateSection subroutine (HINT: You might want to iterate from minSection to maxSection for index i instead of 1 to nSections in the For… Next loops)
- Each of the section files is then formatted; the user (teacher) can decide which assignments to add to each of the sheets
- This is the end of the initialization process



2. (SYNCING) TAs should then be able to add grades to their individual section grade files.  When needed, the teacher can import grades from the section files into the main Grade Manager file (this is known as syncing).  Syncing involves the following:

- Importing the first column of the files (the assignment headings) – this is necessary because adding grade items (see below) will change the headings in the section files but NOT in the main Grade Manager.
- Grades from the individual students will be transferred, matched up with, and placed into the corresponding row in the main Grade Manager file.
- Any preexisting data will be overwritten without warning (you can add warnings if you wish, but this is not a requirement).



3. (BACKUP)  You should have the ability to back up your roster file and your program should save the file as a new name that includes the current date (e.g., “Grade_Manager_ECON_1001_2-18-2018.xlsm” or similar).
4. (ADD GRADE ITEM)  Your Grade Manger must have the capability to add new grade items (columns) but it doesn’t need to be able to add students.

- If a new grade item is added, then your file should place the new column in the proper place in all of the section files (but not the central Grade Manager project – that will be updated during the next sync).
- (DELETE GRADE ITEM)  Your Grade Manager project needs to be able to delete columns of data from all sheets.  I have not provided a screencast for doing this, but in Week 4 of Part 2 of the course, you will have done something similar.  Also, the deleting columns is very similar to (but much easier than) the Add Grade Item portion (above).
- (SEARCH AND REPLACE) Finally, your Grade Manager must have a search/replace tool that does the following

- Allows the teacher to select the student from a drop-down list (combo box) as well as an assignment of interest (also selected from a combo box), and the program should display that grade.
- Allows the teacher to select the student and assignment and replace the current grade with a new/updated score.
- IMPORTANT:  I did not show you this in any of the screencasts, but if the teacher changes a score in the main Grade Manager file, the program also needs to make that change to the individual roster grade sheet.  You will need to come up with a solution for this part!

 

Hints

- When syncing, also import the first row (grade items/labels) in the import.
- Whenever a combobox is used on a user form, it should be unloaded and not hidden at the end of the subroutine.  Otherwise, duplicate items are added to the comboboxes.
- When troubleshooting, make sure to delete all preexisting files and folders that you are trying to create; otherwise, you might have an error if you try to write over a preexisting folder or file.
- All of the code for everything (including user forms) need to be in the “Grade Manager – SETUP.xlsm” file to begin with

 

Requirements/Grading Rubric

Things that your user form MUST do for full credit (these are the criteria that your project is graded on):

- (3 pts, 0.5 pt. each) Initialization

​           a.      Course directory created in user-defined location on local computer

​           b.       Roster file correctly imported to main Grade Manager file

​           c.       Section roster files are correctly created and students separated into those sections automatically upon initialization

​           d.       User-selected assignments and number of assignments are added successfully and correctly to all of the individual section files

​           e.       Be adaptable to section files with differing number of students (each of the 5 section files have different number of students).

​           f.        Be adaptable to different total number of section files (instead of 5 there may be more).

- (2 pts) Syncing of grades - When synced/updated, new grades are successfully imported from the individual section files to the main Grade Manager file
- (1 pt) File backup – There is a way to back up the main file; the backup file has the current date added to the file name.
- (2 pts, 1 pt. each)  Search and Replace Tool

​            a.       User is able to search for an individual student by name; names are automatically populated into a drop-down menu (combo box)

​            b.       User is able to search for an individual (drop-down box) and select an assignment, then change that student’s grade on that assignment

- (1.5 pts) Your program should allow the user to add a new column (assignment) using a drop-down menu.  This assignment should be inserted in the correct spot on the grade spreadsheet, should be numbered properly (one more than the previous maximum of that assignment before adding the new column), and should be added to all section files.
- (0.5 pts.) Your program should allow the user to delete an entire column of information, selected by the user from a drop-down list.  This should make permanent changes to all section files and there should be a confirmation message box before deleting that assignment.
- You must have quit buttons and reset buttons on all user forms!

 

 Other (not necessary for grading but include to enhance your project!):

- Input validation – Your form should be protected against errors, etc., as much as you can. 
- OPTIONAL (would add to your project but *not required*): If a TA has changed a grade item from the last time the grades were imported/synced, it’d be awesome if your form would detect this and let the user know that *n* items have been modified since the last update.