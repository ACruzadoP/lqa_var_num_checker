This app was developed by me while I was waiting for my employer-specific work permit to be approved so that I could go back to Montreal and keep working as Spanish Localization Quality Assurance Tester.<br /><br />

This app has two main functionalities:<br />

<code>1. Spotting numeric-related issues that could have appeared during the Localization process. For example, if there's a String that says "Kill 50 spiders", maybe there was a translator who had entered an incorrect character, causing the string to appear with a wrong number - "Tuez 40 araignées". This could be a serious issue because only Functionality QA Testers are allowed to spend all day long killing monsters in order to make sure that the String is referring to the right amount of spiders. Needless to say that they are not going to play on French.</code><br /><br />
<code>2. Spotting missing variables as well as variables that were inconsistently translated.</code><br /><br />

Things to keep in mind:<br />

&nbsp;&nbsp;&nbsp;&nbsp;- You may need to add a Reference to the project. Feel free to browse "Microsoft.Office.Interop.Excel.dll" - you should be able to locate it in the folder called "Features", as well as a Fake TextFile.<br />
&nbsp;&nbsp;&nbsp;&nbsp;- When indicating the columns where the translations are located, you will have some freedom. For example, you can use spacebar, comma or dash in between every column. Also, you can enter a single column, a sequence of columns as well as an interval. <br />
&nbsp;&nbsp;&nbsp;&nbsp;- If you enter the type of tags surrounding every variable, but only check the Numerical checkbox, the app will filter all the numeric characters that belong to any variable.<br />
&nbsp;&nbsp;&nbsp;&nbsp;- Do not open any Excel File while running the app, otherwise the program will crash. You can launch this program with Excel files already opened.
