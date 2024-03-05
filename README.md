# Vicky3 BuildingEfficiencyOptimizer (BEO for short)
Inspired by Generalist Gaming's spreadsheet, here's a program which can mathematically optimize the prices of goods instead of guessing numbers by hand in the spreadsheet!

# If you don't want to edit the program or touch the code, then download the latest version of the program here and read the instructions included (https://www.mediafire.com/file/i0kl1a85a2e7ysc/Vicky3_BEO.zip/file)

Here's an example which optimizes for construction the imaginary economy with only Logging camp on Sawmills, Iron and Coal mines on Condensing, Tooling workshop with Steel tools, and Steel mill with Blistering.

![kuva](https://github.com/Partrester12/Vicky3BuildingEffeciencyOptimizer/assets/49076600/4800dfa3-b6da-4e8d-bb52-85a7fe3cef92)

...And a comparison between the results given by this program compared to the example in Generalist Gaming's spreadsheet. As you can see, you can make the numbers go up! (Equi. Effic. aren't exactly equal since I rounded them to one decimal)

![kuva](https://github.com/Partrester12/Vicky3BuildingEffeciencyOptimizer/assets/49076600/85f3cf9f-b56d-4d8e-8eaf-53a05fe2d731)


# In order for you to use this, you need to install Jupyter Notebook!!!
But don't worry, it's not too complicated. Here's a link to a youtube tutorial if you're on Windows and have never used conda before: https://www.youtube.com/watch?v=WUeBzT43JyY
Note that you only need to watch like the first 5 minutes or so if you only care about running this program.

Now, let's see how to use this program

# Step 1 Getting the necessary files
Download this repository by clicking the button that says '<> Code' at the top right corner and selecting 'Download ZIP'.
Once you've downloaded the .ZIP, extract the contents to a folder of your choosing. I recommend creating a folder like 'Vicky3 Optimization' on your desktop and dumping the files there.
Note that for this program to work, we only need the .xlsx file and the .ipynb file!

# Step 2 Editing the excel sheet
Open up 'BuildingSheet.xlsx'. In there you will find the few lines included by default and written by yours truly as examples.

The sheet consists of a few important columns:
- Building, PM, and the three Optional PMs ar all cosmetic to make it easier to know what each row represents. You can edit these freely
- Included should be 1 if you want to include this building in your calculations, 0 if not. I recommend only having one building for each PM active (so don't include both Logging camp with Steam Donk and Logging camp without Steam Donk at the same time). The program will work regardless, but it's up to you to figure out what you're optimizing at that point!
- TBonus is the throughput bonus affecting a type of building. This can be from companies, economies of scale etc. This WILL affect the optimal prizes
- ConBonus is the construction bonus affecting a type of building. Usually useful if you want to optimize when having specific companies. This WILL affect the optimal prizes
- Construction and Labor are self explanatory. You can find all these values from Generalist Gaming's spreadsheet

Then there are the most important columns
- Inp-prefix represents all the inputs for a type of building with selected PMs. You can find these values from Generalist Gaming's spreadsheet. Note that you cannot leave any cell in a row empty for this program to work!!!!!
- Out-prefix represents all the outputs for a type of building with selected PMs. You can find these values from Generalist Gaming's spreadsheet. Note that you cannot leave any cell in a row empty for this program to work!!!!!

The input- and output goods are ordered alphabetically. Keep in mind that the output and input goods MUST be in the same exact order as shown in the examples. Otherwise you're gonna get wonky results!!

And as you might guess, adding all your buildings to this by yourself is a herculean task. 
Thus in the future if the community has begun using this, please expand the table and contact me on discord (username: partrester12) so that I can update the default table in the repo!

# Step 3 Using the program
Once you've edited the spreadsheet to match your game (remember to include all the buildings you want to have in your economy!), open Jupyter Notebook and locate the folder where you have the .ipynb file and the 'BuildingSheet.xlsx' file

Click 'OptimizeBuildings.ipynb' to open the file. You should be greeted with a bunch of python code and comments. 

If you don't want to change any variables, select 'Run' from the menu at the top, select 'Run all cells', and scroll to the bottom of the notebook.

**NEVER CLICK THE RUN-BUTTON WITH THE ICON**
This button will simply run the currently selected cell instead of the whole program. Confusing, I know, but it is what it its...

![kuva](https://github.com/Partrester12/Vicky3BuildingEffeciencyOptimizer/assets/49076600/6fd92488-e038-4168-8f01-f43c5af35ab4)

**Depending on the version of Juptyer Notebook, you might want to select 'Cell' from the top row instead of 'Run'**

You should be presented with two tables. The first one is optimized for construction whilst the latter is optimized for number of labor.

If you want to edit the variables, then the only one worth editing is the amount of 'per pops' you're optimizing for as the default is per 100 pops.

# Step 4 What do I do with this information?

The equilibrium prices given by this program roughly equate to prices you want to have in your game, for your buildings to be optimally efficient.
So for example if the optimal price for steel is +75%, then that means you don't want to build any more buildings producing steel than you have to.
This doesn't mean it's bad to build more steel, since construction will be cheaper with cheaper steel, but it means that the buildings producing steel will be less profitable.

...Or this is how I understand it. Check out Generalist Gaming's video on the topic and figure for yourself what to do with this information!
(https://www.youtube.com/watch?v=Jn1RbW4t-M8)

Also take note that this program maximizes the **Net equilibrium efficiency** of all the buildings. This means that there are other equilibrium points, but the ones given by this program optimize the **net value** of all the selected buildings!

Also take into consideration that MAPI is not accounted for in these calculations, but regardless the prices given by this program should be 'close' to optimal

# Step 5 Notes and bugs

Here's a few tips and tricks for editing the spreadsheet whilst the program isn't finished. 

Currently in the process of migriting from the linear solver to a non-linear solver. A known massive issue atm is that if you include more than 20 buildings, since there's like hundreds of constraints, some of them result in non-linearities that break the current optimizer (Lookign at you Railways....).

The current program still works for any set of 10-20 buildings, just don't include Railways, Urban centers or Powerplants. Also highly recommend not having all plantations with auto irrigation, since the engine consumption is so massive compared to outputs (or so my theory goes)

These issues **should** be fixed after I dabble into other optimizers and do some more maths!

# (Link to Generalist Gaming's spreadsheet: https://docs.google.com/spreadsheets/d/1NDUwUWHlNuQTWMejK0Agv8gC3fn6ujiXKE9WfVmjFBc/edit#gid=0)
