
### AutoKeyer

**Description:** tool to automate physical count keying in based off excel sheet to increase efficency. Updated to include visual UI instead of terminal prompts.

**<u> Files: </u>**

- AutoKey.py

  - Asks for an input excel file, sheet name, and starting cell of the table you are inputing

- minus1.py

  - If a mistake was made in the key program, to reset the list, replaces all boxes with "-1" as it is a number that can be replaced without popping up a error message.
  - Once minus1.py has run, you can use the AutoKey.py script again.

- PICountKeyer_v2.py (2025)

  - Version 2 of the program which combines both functionality of "AutoKey.py" and "minus1.py" into a singular program which can be navigated with operational UI elements.
