The provided VBA (Visual Basic for Applications) script performs a specific task involving data extraction from multiple sheets within a workbook and saving that extracted data into separate new files based on predefined criteria. Here's a summary of the project's functionality:

Purpose: The script's main objective is to extract specific data from various sheets within the current workbook and save this extracted data into separate files, each tailored to contain information relevant to a particular category.

Extraction Criteria: The script targets specific sheets in the workbook, identified by their names (e.g., "Sheet 1," "Sheet 2," "Sheet 3," etc.), and extracts data from predetermined ranges within these sheets.

Custom File Creation: For each targeted sheet, a new file is created with a custom file name assigned from an array of predefined names ("Sheet 1," "Sheet 2," etc.).

Data Processing: Once the data is extracted, it is pasted into a sample format file ("Sample files round.xlsx") at predetermined target ranges within specific sheets matching the extracted data's nature.

Formatting and Cleanup: The script ensures that the extracted data aligns with the designated format by copying values and formats from the sample format file. It clears specific columns in the new files and adjusts column widths for better readability.

Saving Files: After the extracted data is properly formatted and processed, the script saves the modified sample format file and the newly created files in a user-selected folder, facilitating organized storage and easy access.

Efficiency and Automation: By utilizing VBA, the script streamlines and automates the process of extracting, formatting, and saving data from multiple sheets, reducing manual effort and ensuring consistency in the output files.

This project demonstrates an automated workflow using VBA, enabling efficient extraction and organization of data from diverse sheets into separate files based on predefined criteria, aiding in data management and analysis tasks.
