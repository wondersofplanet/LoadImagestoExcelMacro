Image Loader and Arranger Readme

Description:
This VBA script is designed to load images from a selected folder on your desktop and arrange them neatly in an Excel worksheet. It allows for easy organization and presentation of multiple images for various purposes such as reports, presentations, or data visualization.

Instructions:

Open Excel.
Press ALT + F11 to open the Visual Basic for Applications (VBA) editor.
In the editor, go to Insert > Module to create a new module.
Copy and paste the provided VBA code into the module window.
Close the VBA editor.
Back in Excel, go to the Developer tab. If you don't see the Developer tab, you may need to enable it in Excel's options.
Click on Macros.
In the Macro dialog box, select LoadAndArrangeImages and click Run.
A folder picker dialog box will appear. Select the folder containing the images you want to load and arrange.
Click OK.
The script will load the images into the active worksheet, arranging them neatly according to the specified settings.
Settings:

columnOffset: Adjusts the width between columns of images.
rowSpacing: Adjusts the vertical spacing between rows of images.
maxColumns: Sets the maximum number of columns of images per row.
maxRows: Sets the maximum number of rows of images.
Notes:

Supported image formats: JPG and PNG. You can modify the code to support additional formats if needed.
Images will be inserted starting from the top-left corner of the worksheet.
The script will overwrite any existing content in the active worksheet.
Ensure that your images are appropriately sized for the layout you desire.
Customize the settings according to your preference and requirements.
Compatibility:

This script is compatible with Excel versions that support VBA macros.
Ensure that macros are enabled in your Excel settings for the script to run.
