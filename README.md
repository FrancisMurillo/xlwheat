wheat
=====

A code export and import tool for VBA since VB has none.

So my first project with VBA Excel left me wanting to use Git. But since the code stayed inside the built-in IDE and exporting the modules one at a time is tedious and tiring, our team just used notepad to copy and paste code segments. Just thinking back again, it gives me chills.

So Wheat was born this exports and imports the code for you, it recreates the same folder structure to make it look the code seamless. Once you've exported the code, open the terminal and commit those changes and push it to the repo; on the flip side, you can pull first the code then import the code. This is better than copy pasting so to say.

This is called *wheat* because of two words: weak and Git. The project idea was it's not a VCS, it exports and imports the code for you, let the real VCS handle the legwork. 

**To God I plant these seed, may it turn into a forest**

quick start
====

This is a <a href="https://github.com/FrancisMurillo/chip">chip</a> project, so you can download this via *Chip.ChipOnFromRepo "Vase"* or if you want to install it via importing module. Just import these three modules in your project.

1. <a href="https://raw.githubusercontent.com/FrancisMurillo/wheat/master/Modules/Vase.bas">Wheat.bas</a>
2. <a href="https://raw.githubusercontent.com/FrancisMurillo/wheat/master/Modules/VaseLib.bas">WheatLib.bas</a>
3. <a href="https://raw.githubusercontent.com/FrancisMurillo/wheat/master/Modules/VaseAssert.bas">WheatConifg.bas</a>

And include in your project references the following.

1. **Microsoft Visual Basic for Applications Extensibility 5.3** - Any version would do but it has been tested with version 5.3
2. **Microsoft Scripting Runtime**

So before we see anything, we just need to alter one line in *WheatConfig.bas* . We need to change the PROJECT_REPO constant.

```
Public Const PROJECT_REPO As String = "wheat-src"
```

Change the PROJECT_REPO constant to a file directory, not a path. This module exports the code in a folder next to your project, so give it a meaningful name. Once you're satisfied with the name, just run in the Intermediate Window or what I like to think of as the *terminal* the following procedure.

```
Wheat.Setup
```

You should see some output saying the repository is completed. Check out the folder where the project is and you should see a new folder with PROJECT_REPO as its name. Now setup your VCS here and execute the following command to export the code.

```
Wheat.Export
```

With that code is exported, and you can check your VCS to see the diffs. Same thing with import, just run the command.

```
Wheat.Import
```

This will copy the code from each of the module to its corresponding modules. I do advise caution on import as it overwrites and adds modules, you might want to check the import source before importing. But in any case, that's the whole process in an nutshell. Setup, Export, Import, Export, Import and so forth and so on.


import and export filtering
====

There is one section I'd like to elaborate before going trigger happy with importing and exporting code. This is the export and import filtering. Check out the default snippet for the configuration file.

```
Public Sub InitializeVariables()
    ' Sample modules to ignore, a reasonable default is provided
    IgnoreExportModules = Array( _
        "Chip*", "Vase*", _
        "Sheet*", "ThisWorkbook", _
        "Sandbox", "ModuleIgnore")
    IgnoreExceptExportModules = Array( _
        "ChipInfo", "ChipInit", "WheatConfig", _
        "ModuleIgnoreNot")
    
    ' Same restriction as exporting
    ' Modify this when to your specific needs
    PassImportModules = IgnoreExportModules
    PassExceptImportModules = IgnoreExceptExportModules
End Sub
```

In particular, the first section with Ignores and the next section with Pass tell how Wheat should filter out the Export and Import. The string pattern uses the Like operator, so you can take advantage of simple wildcards.

From the implementation above can be explained easier with its meaning. IgnoreExport states that is should ignore all Chip and Vase modules, as these are other libraries not related to the main code, and also ignore all Sheet and ThisWorkbook module, as these are document modules you might not want to export upfront; and lastly ignore the Sandbox and ModuleIgnore modules as these are just for playing around with.

Now with IgnoreExcept, these allow modules that are filtered by Ignore to be exported nonetheless. It states the ChipInfo and ChipInit be exported along with code as this provides the Chip configuration file and update module respectively; and also export the wheat configuration file so that people can have a uniform setting; and finally export the sample module, ModuleIgnoreNot, because I feel like it.

The same goes with Import and Pass. However, I set the default to deny importing what you export as import might destroy your work if you're not careful. So if you want to receive changes from others, tweak this like Ignore.

So in way, this is just a simple filtering mechanism that you should tweak for your own needs. 

what's next
====

I'm pretty happy with the tool right now but the one thing I can say is that this hasn't been tested on Word or Access. Although I would not develop there, it might be a feasible use case. The options are good enough
