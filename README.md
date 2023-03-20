<img src="images/box-dev-logo-clip.png" 
alt= “box-dev-logo” 
style="margin-left:-10px;"
width=40%;>
# FolderInfo
Read a spreadsheet distributed by Box with FolderIDs, and determine the folder path for each FolderID

Box distributed Excel spreadsheets to some admins that contained information about some files that had been corrupted. One column of that report was the FolderID that contained the affected file. This script determines the path to that folder from the uploader's point of view and adds a column to an output spreadsheet that contains that path.  It adds a second addtional column containing the user's display name, for help in composing communications.

This project depends on the Box Command Line tool, which can be downloaded from here: https://developer.box.com/docs/box-cli

To run this project, install the Box CLI according to their instructions, including creating a Box Application and downloading credentials. 

You will need to make sure that the following Perl libraries are installed on your system:

    JSON
    Spreadsheet::XLSX
    Excel::Writer::XLSX

Then you can run this program as follows:

    ./FolderInfo <input.xls> <output.xls>

It can take a while to look up all those folders. Be patient.
