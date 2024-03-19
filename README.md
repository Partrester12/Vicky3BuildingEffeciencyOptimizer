# Vicky3 BuildingEfficiencyOptimizer (BEO for short)
Inspired by Generalist Gaming's spreadsheet, here's a program which can mathematically optimize the prices of goods instead of guessing numbers by hand in the spreadsheet!

# If you don't want to edit the program or touch the code, then download the latest version of the program here and read the instructions included

Download for Windows: https://www.mediafire.com/file/cpuxnf8r6emvv28/Vicky3_BEO_Windows.zip/file

Download for Linux (only Ubuntu tested): https://www.mediafire.com/file/i0kl1a85a2e7ysc/Vicky3_BEO_Ubuntu.zip/file

If you download the program, then I still recommend checking steps 3, 4, and 5 from this tutorial as they contain more information than the instructions included with the program!

**Also you will need some software to open Excel sheets with. I recommend OpenOffice (https://www.libreoffice.org/) as it's completely free and what I personally use**

# Now, let's jump into the tutorial

Here's an example which optimizes for construction the imaginary economy with only Logging camp on Sawmills, Iron and Coal mines on Condensing, Tooling workshop with Steel tools, and Steel mill with Blistering.

![kuva](https://github.com/Partrester12/Vicky3BuildingEffeciencyOptimizer/assets/49076600/4800dfa3-b6da-4e8d-bb52-85a7fe3cef92)

...And a comparison between the results given by this program compared to the example in Generalist Gaming's spreadsheet. As you can see, you can make the numbers go up! (Equi. Effic. aren't exactly equal since I rounded them to one decimal)

![kuva](https://github.com/Partrester12/Vicky3BuildingEffeciencyOptimizer/assets/49076600/85f3cf9f-b56d-4d8e-8eaf-53a05fe2d731)


# In order for you to use this, you need to install Jupyter Notebook!!!
But don't worry, it's not too complicated. Here's a link to a youtube tutorial if you're on Windows and have never used conda before: https://www.youtube.com/watch?v=WUeBzT43JyY
Note that you only need to watch like the first 5 minutes or so if you only care about running this program.

**You will also need to install some Python packages, or the program will not run.**

On Windows you can do this by searching for a program named 'Anaconda Prompt', launching it, and typing in the following commands (and accepting the installation conditions of each package):
'conda install pip'
'pip install pulp'
'pip install mystic'
'pip install openpyxl'

If you already have these packages installed, then no worries. There are also some other packages the program uses, but those come preinstalled with Conda if you follow the video tutorial I have linked above.

Now, let's see how to use this program

# Step 1 Getting the necessary files
Download this repository by clicking the button that says '<> Code' at the top right corner and selecting 'Download ZIP'.
Once you've downloaded the .ZIP, extract the contents to a folder of your choosing. I recommend creating a folder like 'Vicky3 Optimization' on your desktop and dumping the files there.
Note that for this program to work, we only need the .xlsx file and the .ipynb file!

# Step 2 Editing the excel sheet
Open up 'BuildingSheet.xlsx'. In there you will find the few lines included by default and written by yours truly as examples.

The sheet consists of a few important columns:
- Building, PM, and the three Optional PMs are the heart and soul of this program. You can click on the fields, choose the PMs you want, and the sheet will automatically calculate labor, inputs, and outputs for that building!
- Included should be 1 if you want to include this building in your calculations, 0 if not
- TBonus is the throughput bonus affecting a type of building. This can be from companies, economies of scale etc. This WILL affect the optimal prizes
- ConBonus is the construction bonus affecting a type of building. Usually useful if you want to optimize when having specific companies. This WILL affect the optimal prizes
- Construction and Labor are self explanatory. Automatically calculated by choosing the PMs so don't edit manually!

Then there are the most important columns
- Inp-prefix represents all the inputs for a type of building with selected PMs. These will be updated automatically depending on PMs. Note that you cannot leave any cell in a row empty for this program to work!!!!!
- Out-prefix represents all the outputs for a type of building with selected PMs. These will be updated automatically depending on PMs. Note that you cannot leave any cell in a row empty for this program to work!!!!!

The input- and output goods are ordered alphabetically. Keep in mind that the output and input goods MUST be in the same exact order as shown in the examples. Otherwise you're gonna get wonky results!!

# Step 3 Using the program
Once you've edited the spreadsheet to match your game (remember to include all the buildings you want to have in your economy!), open Jupyter Notebook and locate the folder where you have the .ipynb file and the 'BuildingSheet.xlsx' file

Click 'OptimizeBuildings.ipynb' to open the file. You should be greeted with a bunch of python code and comments. 

If you don't want to change any variables, select 'Run' from the menu at the top, select 'Run all cells', and scroll to the bottom of the notebook.

**NEVER CLICK THE RUN-BUTTON WITH THE ICON**
This button will simply run the currently selected cell instead of the whole program. Confusing, I know, but it is what it its...

![kuva](https://github.com/Partrester12/Vicky3BuildingEffeciencyOptimizer/assets/49076600/6fd92488-e038-4168-8f01-f43c5af35ab4)

**Depending on the version of Juptyer Notebook, you might want to select 'Cell' from the top row instead of 'Run'**

You should be presented with two tables. The first one is optimized for construction whilst the latter is optimized for number of labor.

If you want to edit the variables, then the only ones worth editing is MAX_NUMBER_ITERARTIONS, which controls how long the program will try to calculate optimal prices in cases where optimization is difficult (or close to impossible with common PCs)
I have tested different values and the existing one is optimal for most cases, but feel free to edit it if you feel like testing performance!

# Step 4 What do I do with this information?

The equilibrium prices given by this program roughly equate to prices you want to have in your game, for your buildings to be optimally efficient.
So for example if the optimal price for steel is +75%, then that means you don't want to build any more buildings producing steel than you have to.
This doesn't mean it's bad to build more steel, since construction will be cheaper with cheaper steel, but it means that the buildings producing steel will be less profitable.

...Or this is how I understand it. Check out Generalist Gaming's video on the topic and figure for yourself what to do with this information!
(https://www.youtube.com/watch?v=Jn1RbW4t-M8)

Also take note that this program maximizes the **Net equilibrium efficiency** of all the buildings. This means that there are other equilibrium points, but the ones given by this program optimize the **net value** of all the selected buildings!

Also take into consideration that MAPI is not accounted for in these calculations, but regardless the prices given by this program should be 'close' to optimal

# Step 5 Notes and bugs

Be patient with getting results. If you think it's taking way too long to get your 'OptimizedPrices.xlsx' consistantly, then either follow the instructions and edit the .ipynb -file or let me know and I can create a version of the program that uses significantly less computing power (though keep in mind that the results will also be worse!)

Windows might have some problems where the OptimizedPrices sheet will have unnecessary comments added. I have not yet dove into why this happens, but as it's not breaking anything, it's a feature for now :P

# (Link to Generalist Gaming's spreadsheet: https://docs.google.com/spreadsheets/d/1NDUwUWHlNuQTWMejK0Agv8gC3fn6ujiXKE9WfVmjFBc/edit#gid=0)
