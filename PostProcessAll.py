#Author-Tim Paterson
#Description-Post process all CAM setups, using the setup name as the output file name.

import adsk.core, adsk.fusion, adsk.cam, traceback, shutil, json, os, os.path, time, re, pathlib, enum, tempfile

# Version number of settings as saved in documents and settings file
# update this whenever settings content changes
version = 7

# Initial default values of settings
defaultSettings = {
    "version" : version,
    "post" : "",
    "units" : adsk.cam.PostOutputUnitOptions.DocumentUnitsOutput,
    "output" : "",
    "sequence" : True,
    "twoDigits" : False,
    "delFiles" : False,
    "delFolder" : False,
    "splitSetup" : False,
    "fastZ" : False,
    "toolChange" : "M9 G30",
    "fileExt" : ".nc",
    "numericName" : False,
    "homeEndsOp" : False,
    # Groups are expanded or not
    "groupOutput" : True,
    "groupPersonal" : True,
    "groupPost" : True,
    "groupAdvanced" : False,
    # Retry policy
    "initialDelay" : 0.2,
    "postRetries" : 3
}

# Constants
constCmdName = "Post Process All"
constCmdDefId = "PatersonTech_PostProcessAll"
constCAMWorkspaceId = "CAMEnvironment"
constCAMActionsPanelId = "CAMActionPanel"
constPostProcessControlId = "IronPostProcess"
constCAMProductId = "CAMProductType"
constAttrGroup = constCmdDefId
constAttrName = "settings"
constSettingsFileExt = ".settings"
constPostLoopDelay = 0.1
constBodyTmpFile = "gcodeBody"
constOpTmpFile = "8910"   # in case name must be numeric
constRapidZgcode = 'G00 Z{} (Changed from: "{}")\n'
constRapidXYgcode = 'G00 {} (Changed from: "{}")\n'
constFeedZgcode = 'G01 Z{} F{} (Changed from: "{}")\n'
constFeedXYgcode = 'G01 {} F{} (Changed from: "{}")\n'
constFeedXYZgcode = 'G01 {} Z{} F{} (Changed from: "{}")\n'
constAddFeedGcode = " F{} (Feed rate added)\n"
constMotionGcodeSet = {0,1,2,3,33,38,73,76,80,81,82,84,85,86,87,88,89}
constEndMcodeSet = {5,9,30}
constLineNumInc = 5

# Tool tip text
toolTip = (
    "Post process all setups into G-code for your machine.\n\n"
    "The name of the setup is used for the name of the output "
    "file adding the .nc extension. A colon (':') in the name indicates "
    "the preceding portion is the name of a subfolder. Multiple "
    "colons can be used to nest subfolders. Spaces around colons "
    "are removed.\n\n"
    "Setups within a folder are optionally preceded by a "
    "sequence number. This identifies the order in which the "
    "setups appear. The sequence numbers for each folder begin "
    "with 1."
    )

# Global list to keep all event handlers in scope.
# This is only needed with Python.
handlers = []

# Global settingsMgr object
settingsMgr = None

def run(context):
    global settingsMgr
    ui = None
    try:
        settingsMgr = SettingsManager()
        app = adsk.core.Application.get()
        ui  = app.userInterface
        InitAddIn()

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


class SettingsManager:
    def __init__(self):
        self.default = None
        self.path = None
        self.fMustSave = False
        self.inputs = None

    def GetSettings(self, docAttr):
        docSettings = None
        attr = docAttr.itemByName(constAttrGroup, constAttrName)
        if attr:
            try:
                docSettings = json.loads(attr.value)
                if docSettings["version"] == version:
                    return docSettings
            except Exception:
                pass
            
        # Document does not have valid settings, get defaults
        if not self.default:
            # Haven't read the settings file yet
            file = None
            try:
                file = open(self.GetPath())
                self.default = json.load(file)
                # never allow delFiles or delFolder to default to True
                self.default["delFiles"] = False;
                self.default["delFolder"] = False;
                if self.default["version"] != version:
                    self.UpdateSettings(defaultSettings, self.default)
            except Exception:
                self.default = dict(defaultSettings)
                self.fMustSave = True
            finally:
                if file:
                    file.close
        
        if not docSettings:
            docSettings = dict(self.default)
        else:
            self.UpdateSettings(self.default, docSettings)
        return docSettings

    def SaveDefault(self, docSettings):
        self.fMustSave = False
        self.default = dict(docSettings)
        # never allow delFiles or delFolder to default to True
        self.default["delFiles"] = False;
        self.default["delFolder"] = False;
        try:
            strSettings = json.dumps(docSettings)
            file = open(self.GetPath(), "w")
            file.write(strSettings)
            file.close
        except Exception:
            pass

    def SaveSettings(self, docAttr, docSettings):
        if self.fMustSave:
            self.SaveDefault(docSettings)
        docAttr.add(constAttrGroup, constAttrName, json.dumps(docSettings))
            
    def UpdateSettings(self, src, dst):
        for item in src:
            if not (item in dst):
                dst[item] = src[item]
        dst["version"] = src["version"]

    def GetPath(self):
        if not self.path:
            pos = __file__.rfind(".")
            if pos == -1:
                pos = len(__file__)
            self.path = __file__[0:pos] + constSettingsFileExt
        return self.path


def InitAddIn():
    ui = None
    try:
        app = adsk.core.Application.get()
        ui  = app.userInterface

        # Create a button command definition.
        cmdDefs = ui.commandDefinitions
        cmdDef = cmdDefs.addButtonDefinition(constCmdDefId, constCmdName, toolTip, "resources/Command")
        
        # Connect to the commandCreated event.
        commandEventHandler = CommandEventHandler()
        cmdDef.commandCreated.add(commandEventHandler)
        handlers.append(commandEventHandler)
        
        # Get the Actions panel in the Manufacture workspace.
        workSpace = ui.workspaces.itemById(constCAMWorkspaceId)
        addInsPanel = workSpace.toolbarPanels.itemById(constCAMActionsPanelId)
        
        # Add the button right after the Post Process command.
        cmdControl = addInsPanel.controls.addCommand(cmdDef, constPostProcessControlId, False)
        cmdControl.isPromotedByDefault = True
        cmdControl.isPromoted = True

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


def CountOutputFolderFiles(folder, limit, fileExt):
    cntFiles = 0
    cntNcFiles = 0
    for path, dirs, files in os.walk(folder):
        for file in files:
            if file.endswith(fileExt):
                cntNcFiles += 1
            else:
                cntFiles += 1
        if cntFiles > limit:
            return "many files that are not G-code"
        if cntNcFiles > limit * 1.5:
            return "many more G-code files than are produced by this design";
    return None


# Event handler for the commandCreated event.
class CommandEventHandler(adsk.core.CommandCreatedEventHandler):
    def __init__(self):
        super().__init__()

    def notify(self, args):
        try:
            eventArgs = adsk.core.CommandCreatedEventArgs.cast(args)
            cmd = eventArgs.command

            # Get document attributes that will set initial values
            app = adsk.core.Application.get()
            docSettings  = settingsMgr.GetSettings(app.activeDocument.attributes)

            # Add inputs that will appear in a dialog
            inputs = cmd.commandInputs

            # output folder
            inputGroup = inputs.addGroupCommandInput("groupOutput", "Output Folder")
            input = inputGroup.children.addStringValueInput("output", "", docSettings["output"])
            input.isFullWidth = True
            input.tooltip = "Output Folder"
            input.tooltipDescription = (
                "Full path name of the output folder. Any subfolders, as denoted "
                "by colons in the setup name, will be relative this folder.")

            input = inputGroup.children.addBoolValueInput("browseOutput", "Browse", False)
            input.resourceFolder = "resources/Browse"
            input.tooltip = "Browse for Output Folder"
            inputGroup.isExpanded = docSettings["groupOutput"]

            # check box to delete existing files
            input = inputs.addBoolValueInput("delFiles", 
                                             "Delete existing files", 
                                             True, 
                                             "", 
                                             docSettings["delFiles"])
            input.tooltip = "Delete Existing Files in Each Folder"
            input.tooltipDescription = (
                "Delete all files in each output folder before post processing. "
                "This will help prevent accumulation of G-code files which are "
                "no longer used."
                "<p>For example, you could decide to add sequence numbers after "
                "already post processing without them. If this option is not "
                "checked, you will have two of each file, a newer one with a "
                "sequence number and older one without. With this option checked, "
                "all previous files will be deleted so only current results will "
                "be present.</p>"
                "<p>This option will only delete the files in folders in which new "
                "G-code files are being written. If you change the name of a "
                "folder, for example, it will not be deleted.</p>")

            # check box to delete entire output folder
            input = inputs.addBoolValueInput("delFolder", 
                                             "Delete output folder", 
                                             True, 
                                             "", 
                                             docSettings["delFolder"] and docSettings["delFiles"])
            input.isEnabled = docSettings["delFiles"] # enable only if delete existing files
            input.tooltip = "Delete Entire Output Folder First"
            input.tooltipDescription = (
                "Delete the entire output folder before post processing. This "
                "deletes all files and subfolders regardless of whether or not "
                "new G-code files are written to a particular folder."
                "<p><b>WARNING!</b> Be absolutely sure the output folder is set "
                "correctly before selecting this option. Run the command once "
                "before setting this option and verify the results are in the "
                "correct folder. An incorrect setting of the output folder with "
                "this option selected could result in unintentionally wiping out "
                "a vast number of files.</p>")

            # check box to prepend sequence numbers
            input = inputs.addBoolValueInput("sequence", 
                                             "Prepend sequence number", 
                                             True, 
                                             "", 
                                             docSettings["sequence"])
            input.tooltip = "Add Sequence Numbers to Name"
            input.tooltipDescription = (
                "Begin each file name with a sequence number. The numbering "
                "represents the order that the setups appear in the browser tree. "
                "Each folder has its own sequence numbers starting with 1.")

            # check box to select 2-digit sequence numbers
            input = inputs.addBoolValueInput("twoDigits", 
                                             "Use 2-digit numbers", 
                                             True, 
                                             "", 
                                             docSettings["twoDigits"])
            input.isEnabled = docSettings["sequence"] # enable only if using sequence numbers
            input.tooltip = "Use 2-Digit Sequence Numbers"
            input.tooltipDescription = (
                "Sequence numbers 0 - 9 will have a leading zero added, becoming"
                '"01" to "09". This could be useful for formatting or sorting.')

            # select units
            input = inputs.addDropDownCommandInput("units", 
                                                   "Post output units",
                                                   adsk.core.DropDownStyles.TextListDropDownStyle)
            # Document Unit = 0
            # Inches = 1
            # Millimeters = 2
            # We'll use these convenient values so the item index will be the value
            input.listItems.add('Document units', True)
            input.listItems.add('Inches', False)
            input.listItems.add('Milimeters', False)
            input.listItems.item(docSettings["units"]).isSelected = True;
            input.isFullWidth = True
            input.tooltip = "Post Output Units"
            input.tooltipDescription = (
                "Select the units to be used in the output - inches or millimeters. "
                "This may be chosen explicitly, or the units used in the design "
                "can be used.")

            # "Personal Use" version
            # check box to split up setup into individual operations
            inputGroup = inputs.addGroupCommandInput("groupPersonal", "Personal Use")
            input = inputGroup.children.addBoolValueInput("splitSetup",
                                                          "Use individual operations",
                                                          True,
                                                          "",
                                                          docSettings["splitSetup"])
            input.tooltip = "Split Setup Into Individual Operations"
            input.tooltipDescription = (
                "Generate output for each operation individually. This is usually "
                "REQUIRED when using Fusion 360 for Personal Use, because tool "
                "changes are not allowed. The individual operations will be "
                "grouped back together into the same file, eliminating this "
                "limitation. You will get an error if there is a tool change "
                "in a setup and this options is not selected.")

            # text box as a label for tool change command
            input = inputGroup.children.addTextBoxCommandInput("toolLabel", 
                                                               "", 
                                                               "G-code for tool change:",
                                                               1,
                                                               True)
            input.isFullWidth = True
            label = input

            # enter G-code for tool change
            input = inputGroup.children.addStringValueInput("toolChange", "", docSettings["toolChange"])
            input.isEnabled = docSettings["splitSetup"] # enable only if using individual operations
            input.isFullWidth = True
            input.tooltip = "G-code to Precede Tool Change"
            input.tooltipDescription = (
                "Allows inserting a line of code before tool changes. For example, "
                "you might want M5 (spindle stop), M9 (coolant stop), and/or G28 or G30 "
                "(return to home). The code will be placed on the line before the "
                "tool change. You can get mulitple lines by separating them with "
                "a colon (:)."
                "<p>If you want a line number, just put a dummy line number in front. "
                "If you use the colon to get multiple lines, only put the dummy line "
                "number on the first line. For example, <b>N10 M9:G30</b> will give "
                "you two lines, both with properly sequenced line numbers.</p>"
            )
            label.tooltip = input.tooltip
            label.tooltipDescription = input.tooltipDescription
           
            # check box to enable restoring rapid moves
            input = inputGroup.children.addBoolValueInput("fastZ",
                                                          "Restore rapid moves",
                                                          True,
                                                          "",
                                                          docSettings["fastZ"])
            input.isEnabled = docSettings["splitSetup"] # enable only if using individual operations
            input.tooltip = "Restore Rapid Moves (Experimental)"
            input.tooltipDescription = (
                "Replace appropriate moves at feed rate with rapid (G0) moves. "
                "In Fusion 360 for Personal Use, moves that could be rapid are "
                "now limited to the current feed rate. When this option is selected, "
                "the G-code will be analyzed to find moves at or above the feed "
                "height and replace them with rapid moves."
                "<p><b>WARNING!<b> This option should be used with caution. "
                "Review the G-code to verify it is correct. Comments have been "
                "added to indicate the changes.")
           
            # G-code end marked by move home
            input = inputGroup.children.addBoolValueInput("homeEndsOp",
                                                          "Move home ends op",
                                                          True,
                                                          "",
                                                          docSettings["homeEndsOp"])
            input.isEnabled = docSettings["splitSetup"] # enable only if using individual operations
            input.tooltip = "Move to Home Marks End of Operation"
            input.tooltipDescription = (
                "To combine operations generated individually, the ending sequence "
                "(which should only appear once) must be found. This option includes "
                "a return to home (G28/G30) in the set that marks the ending sequence."
                "<p>You should select this option if you find that there is a G28 or "
                "G30 at the end of each operation of a G-code file.")
            inputGroup.isExpanded = docSettings["groupPersonal"]

            # Advanced -- retry settings
            inputGroup = inputs.addGroupCommandInput("groupAdvanced", "Advanced")
            # Time delay
            input = inputGroup.children.addFloatSpinnerCommandInput("initialDelay", 
                "Initial time allowance", "s", 0.1, 1.0, 0.1, docSettings["initialDelay"])
            input.tooltip = "Initial Time to Post Process an Operation"
            input.tooltipDescription = (
                "Initial delay to wait for post processor. Doubled for each retry.")
            # Retry count
            input = inputGroup.children.addIntegerSpinnerCommandInput("postRetries", 
                "Number of retries", 1, 9, 1, docSettings["postRetries"])
            input.tooltip = "Number of Retries"
            input.tooltipDescription = (
                "Retries if post processing failed. Time delay is doubled each retry.")
            inputGroup.isExpanded = docSettings["groupAdvanced"]
            
            # post processor
            inputGroup = inputs.addGroupCommandInput("groupPost", "Post Processor")
            input = inputGroup.children.addStringValueInput("post", "", docSettings["post"])
            input.isFullWidth = True
            input.tooltip = "Post Processor"
            input.tooltipDescription = (
                "Full path name of the post processor (.cps file).")
            
            # Browse for post processor
            input = inputGroup.children.addBoolValueInput("browsePost", "Browse", False)
            input.resourceFolder = "resources/Browse"
            input.tooltip = "Browse for Post Processor"
            inputGroup.isExpanded = docSettings["groupPost"]

            # Numeric name required?
            input = inputGroup.children.addBoolValueInput("numericName",
                                                          "Name must be numeric",
                                                          True,
                                                          "",
                                                          docSettings["numericName"])
            input.tooltip = "Output File Name Must Be Numeric"
            input.tooltipDescription = (
                "The name of the setup will not be used in the file name, "
                "only sequence numbers. The option to prepend sequence numbers "
                "will have no effect.")

            # File name extension
            input = inputGroup.children.addStringValueInput("fileExt", "Output file extension", docSettings["fileExt"])
            input.tooltip = "Output File Extension"
            
            # button to save default settings
            input = inputs.addBoolValueInput("save", "Save as default", False)
            input.resourceFolder = "resources/Save"
            input.tooltip = "Save These Settings as System Default"
            input.tooltipDescription = (
                "Save these settings to use as the default for each new design.")

            # text box for error messages
            input = inputs.addTextBoxCommandInput("error", "", "", 3, True)
            input.isFullWidth = True
            input.isVisible = False

            # Connect to the inputChanged event.
            onInputChanged = CommandInputChangedHandler(docSettings)
            cmd.inputChanged.add(onInputChanged)
            handlers.append(onInputChanged)

            # Connect to the validateInputs event.
            onValidateInputs = CommandValidateInputsHandler()
            cmd.validateInputs.add(onValidateInputs)
            handlers.append(onValidateInputs)

            # Connect to the execute event.
            onExecute = CommandExecuteHandler(docSettings)
            cmd.execute.add(onExecute)
            handlers.append(onExecute)
        except:
            ui = app.userInterface
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

# Event handler for the inputChanged event.
class CommandInputChangedHandler(adsk.core.InputChangedEventHandler):
    def __init__(self, docSettings):
        self.docSettings = docSettings
        super().__init__()

    def notify(self, args):
        app = adsk.core.Application.get()
        ui  = app.userInterface
        try:
            eventArgs = adsk.core.InputChangedEventArgs.cast(args)
            inputs = eventArgs.inputs

            doc = app.activeDocument
            product = doc.products.itemByProductType(constCAMProductId)

            # See if button clicked
            input = eventArgs.input
            if input.id == "save":
                settingsMgr.SaveDefault(self.docSettings)
                
            elif input.id == "browsePost":
                dialog = ui.createFileDialog()
                post = self.docSettings["post"]
                if len(post) != 0:
                    dialog.initialFilename = post
                else:
                    dialog.initialDirectory = product.genericPostFolder

                dialog.filter = "post processors (*.cps);;All files (*.*)"
                dialog.title = "Select post processor"
                if dialog.showOpen() == adsk.core.DialogResults.DialogOK:
                    self.docSettings["post"] = dialog.filename
                    inputs.itemById("post").value = dialog.filename

            elif input.id == "browseOutput":
                dialog = ui.createFolderDialog()
                dialog.initialDirectory = self.docSettings["output"]
                dialog.title = "Select output folder"
                if dialog.showDialog() == adsk.core.DialogResults.DialogOK:
                    self.docSettings["output"] = dialog.folder
                    inputs.itemById("output").value = dialog.folder

            elif input.id == "units":
                self.docSettings[input.id] = input.selectedItem.index

            elif input.id in self.docSettings:
                if input.objectType == adsk.core.GroupCommandInput.classType():
                    self.docSettings[input.id] = input.isExpanded
                else:
                    self.docSettings[input.id] = input.value

            # Enable twoDigits only if sequence is true
            if input.id == "sequence":
                inputs.itemById("twoDigits").isEnabled = input.value

            # Enable delFolder only if delFiles is true
            if input.id == "delFiles":
                item = inputs.itemById("delFolder")
                item.value = input.value and item.value
                item.isEnabled = input.value

            # Options for splitSetup
            if input.id == "splitSetup":
                inputs.itemById("toolChange").isEnabled = input.value
                inputs.itemById("toolLabel").isEnabled = input.value
                inputs.itemById("fastZ").isEnabled = input.value

        except:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


# Event handler for the validateInputs event.
class CommandValidateInputsHandler(adsk.core.ValidateInputsEventHandler):
    def __init__(self):
        super().__init__()

    def notify(self, args):
        app = adsk.core.Application.get()
        ui  = app.userInterface
        try:
            eventArgs = adsk.core.ValidateInputsEventArgs.cast(args)
            inputs = eventArgs.firingEvent.sender.commandInputs

            fIsOutputValid = len(inputs.itemById("output").value) != 0
            post = inputs.itemById("post").value
            fIsPostValid = post.endswith(".cps") and os.path.isfile(post)
            eventArgs.areInputsValid = fIsOutputValid and fIsPostValid
            error = inputs.itemById("error")
            error.isVisible = not eventArgs.areInputsValid
            if not eventArgs.areInputsValid:
                # Build a message explaining what's missing
                err1 = err2 = combine = ""
                if not fIsOutputValid:
                    err1 = "the output folder"
                    # ensure it's not collapsed
                    inputs.itemById("groupOutput").isExpanded = True
                if not fIsPostValid:
                    err2 = "a valid post processor"
                    # ensure it's not collapsed
                    inputs.itemById("groupPost").isExpanded = True
                if not fIsOutputValid and not fIsPostValid:
                    combine = " and "
                msg = "<b>Please select {}{}{}.</b>".format(err1, combine, err2)
                # Display message
                error.formattedText = msg
        except:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


# Event handler for the execute event.
class CommandExecuteHandler(adsk.core.CommandEventHandler):
    def __init__(self, docSettings):
        self.docSettings = docSettings
        super().__init__()

    def notify(self, args):
        eventArgs = adsk.core.CommandEventArgs.cast(args)

        # Code to react to the event.
        PerformPostProcess(self.docSettings)


def stop(context):
    ui = None
    try:
        app = adsk.core.Application.get()
        ui  = app.userInterface

        # Clean up the UI.
        cmdDef = ui.commandDefinitions.itemById(constCmdDefId)
        if cmdDef:
            cmdDef.deleteMe()
            
        addinsPanel = ui.allToolbarPanels.itemById(constCAMActionsPanelId)
        cmdControl = addinsPanel.controls.itemById(constCmdDefId)
        if cmdControl:
            cmdControl.deleteMe()
    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))	


def PerformPostProcess(docSettings):
    ui = None
    progress = None
    try:
        app = adsk.core.Application.get()
        ui  = app.userInterface
        doc = app.activeDocument

        # Save settings in document attributes
        settingsMgr.SaveSettings(doc.attributes, docSettings)

        cntFiles = 0
        cntSkipped = 0
        lstSkipped = ""
        product = doc.products.itemByProductType(constCAMProductId)

        if product != None:
            cam = adsk.cam.CAM.cast(product)
            setups = cam.setups
            progress = ui.createProgressDialog()
            progress.isCancelButtonShown = True
            progress.show("Generating toolpaths...", "Beginning toolpath generation", 0, 1)
            progress.progressValue = 1; # try to get it to display
            progress.progressValue = 0;

            if setups.count != 0 and cam.allOperations.count != 0:
                # make sure we're not going to delete too much
                outputFolder = docSettings["output"]
                if docSettings["delFolder"]:
                    strMsg = CountOutputFolderFiles(outputFolder, setups.count, docSettings["fileExt"])
                    if strMsg:
                        docSettings["delFolder"] = False
                        strMsg = (
                            "The output folder contains {}. "
                            "It will not be deleted. You may wish to make sure you selected "
                            "the correct folder. If you want the folder deleted, you must "
                            "do it manually."
                            ).format(strMsg)
                        res = ui.messageBox(strMsg, 
                                            constCmdName,
                                            adsk.core.MessageBoxButtonTypes.OKCancelButtonType,
                                            adsk.core.MessageBoxIconTypes.WarningIconType)
                        if res == adsk.core.DialogResults.DialogCancel:
                            return  # abort!

                # generate toolpaths with progess dialog
                genStat = cam.generateAllToolpaths(True)
                if not genStat.isGenerationCompleted:
                    progress.maximumValue = genStat.numberOfOperations;
                    progress.message = "Generating toolpath %v of %m";
                    while not genStat.isGenerationCompleted:
                        if progress.wasCancelled:
                            return  #abort!
                        progress.progressValue = genStat.numberOfCompleted
                        time.sleep(.1)

                if False and not cam.checkAllToolpaths():   # checkAllToolpaths() always throws exception
                    res = ui.messageBox("Some toolpaths are not valid. If you continue, affected Setups will be skipped", 
                                        constCmdName,
                                        adsk.core.MessageBoxButtonTypes.OKCancelButtonType,
                                        adsk.core.MessageBoxIconTypes.WarningIconType)
                    if res == adsk.core.DialogResults.DialogCancel:
                        return  # abort!
                
                if docSettings["delFolder"]:
                    try:
                        shutil.rmtree(outputFolder, True)
                    except:
                        pass #ignore errors

                progressMsg = "{} files written to " + outputFolder
                progress.show("Post Processing...", "", 0, setups.count)

                cntSetups = 0
                seqDict = dict()

                for setup in setups:
                    if progress.wasCancelled:
                        break
                    if not setup.isSuppressed and setup.allOperations.count != 0:
                        nameList = setup.name.split(':')    # folder separator
                        setupFolder = outputFolder
                        cnt = len(nameList) - 1
                        i = 0
                        while i < cnt:
                            setupFolder += "/" + nameList[i].strip()
                            i += 1
                    
                        # keep a separate sequence number for each folder
                        if setupFolder in seqDict:
                            seqDict[setupFolder] += 1
                        else:
                            seqDict[setupFolder] = 1
                            # first file for this folder
                            if (docSettings["delFiles"]):
                                # delete all the files in the folder
                                try:
                                    for entry in os.scandir(setupFolder):
                                        if entry.is_file():
                                            try:
                                                os.remove(entry.path)
                                            except:
                                                pass #ignore errors
                                except:
                                    pass #ignore errors

                        # prepend sequence number if enabled
                        fname = nameList[i].strip()
                        if docSettings["sequence"] or docSettings["numericName"]:
                            seq = seqDict[setupFolder]
                            seqStr = str(seq)
                            if docSettings["twoDigits"] and seq < 10:
                                seqStr = "0" + seqStr
                            if docSettings["numericName"]:
                                fname = seqStr
                            else:
                                fname = seqStr + ' ' + fname

                        # post the file
                        status = PostProcessSetup(fname, setup, setupFolder, docSettings)
                        if status == None:
                            cntFiles += 1
                        else:
                            cntSkipped += 1
                            lstSkipped += "\nFailed on setup " + setup.name + ": " + status
                         
                    cntSetups += 1
                    progress.message = progressMsg.format(cntFiles)
                    progress.progressValue = cntSetups

            progress.hide()

        # done with setups, report results
        if cntSkipped != 0:
            ui.messageBox("{} files were written. {} Setups were skipped due to error:{}".format(cntFiles, cntSkipped, lstSkipped), 
                constCmdName, 
                adsk.core.MessageBoxButtonTypes.OKButtonType,
                adsk.core.MessageBoxIconTypes.WarningIconType)
            
        elif cntFiles == 0:
            ui.messageBox('No CAM operations posted', 
                constCmdName, 
                adsk.core.MessageBoxButtonTypes.OKButtonType,
                adsk.core.MessageBoxIconTypes.WarningIconType)

    except:
        if progress:
            progress.hide()
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

def PostProcessSetup(fname, setup, setupFolder, docSettings):
    ui = None
    fileHead = None
    fileBody = None
    fileOp = None
    retVal = "Fusion 360 reported an exception"

    try:
        app = adsk.core.Application.get()
        ui  = app.userInterface
        doc = app.activeDocument
        product = doc.products.itemByProductType(constCAMProductId)
        cam = adsk.cam.CAM.cast(product)

        # Create PostProcessInput
        opName = fname
        opFolder = setupFolder
        if docSettings["splitSetup"]:
            opName = constOpTmpFile
            opFolder = tempfile.gettempdir()
        postInput = adsk.cam.PostProcessInput.create(opName, 
                                                    docSettings["post"], 
                                                    opFolder, 
                                                    docSettings["units"])
        postInput.isOpenInEditor = False

        # Do it all at once?
        if not docSettings["splitSetup"]:
            try:
                if not cam.postProcess(setup, postInput):
                    return "Fusion 360 reported an error."
                time.sleep(constPostLoopDelay) # files missing sometimes unless we slow down (??)
                return None
            except Exception as exc:
                retVal += ": " + str(exc)
                return retVal

        # Split setup into individual operations
        path = setupFolder + "/" + fname
        fileExt = docSettings["fileExt"]
        pathlib.Path(setupFolder).mkdir(parents=True, exist_ok=True)
        fileHead = open(path + fileExt, "w")
        fileBody = open(opFolder + "/" + constBodyTmpFile + fileExt, "w")
        fFirst = True
        lineNum = 10
        regToolComment = re.compile(r"\(T[0-9]+\s")
        fFastZenabled = docSettings["fastZ"]
        regBody = re.compile(r""
            "(?P<N>N[0-9]+ *)?" # line number
            "(?P<line>"         # line w/o number
            "(M(?P<M>[0-9]+) *)?" # M-code
            "(T(?P<T>[0-9]+))?" # Tool
            ".+)",              # to end of line
            re.IGNORECASE | re.DOTALL)
        toolChange = docSettings["toolChange"]
        fToolChangeNum = False
        if len(toolChange) != 0:
            toolChange = toolChange.replace(":", "\n")
            toolChange += "\n"
            match = regBody.match(toolChange).groupdict()
            if match["N"] != None:
                fToolChangeNum = True
                toolChange = match["line"]
                # split into individual lines to add line numbers
                toolChange = toolChange.splitlines(True)
        if fFastZenabled:
            regGcode = re.compile(r""
                "(G(?P<G>[0-9]+(\.[0-9]*)?)[^XYZF]*)?"
                "(?P<XY>((X-?[0-9]+(\.[0-9]*)?)[^XYZF]*)?"
                "((Y-?[0-9]+(\.[0-9]*)?)[^XYZF]*)?)"
                "(Z(?P<Z>-?[0-9]+(\.[0-9]*)?)[^XYZF]*)?"
                "(F(?P<F>-?[0-9]+(\.[0-9]*)?)[^XYZF]*)?",
                re.IGNORECASE)

        i = 0;
        ops = setup.allOperations;
        while i < ops.count:
            op = ops[i]
            i += 1
            if op.isSuppressed:
                continue

            # Look ahead for operations without a toolpath. This can happen
            # with a manual operation. Group it with current operation.
            # Or if first, group it with subsequent ones.
            opHasTool = None
            hasTool = op.hasToolpath
            if hasTool:
                opHasTool = op
            opList = adsk.core.ObjectCollection.create()
            opList.add(op)
            while i < ops.count:
                op = ops[i]
                if op.isSuppressed:
                    i += 1
                    continue
                if op.hasToolpath:
                    if not hasTool:
                        opList.add(op)
                        opHasTool = op
                        i += 1
                    break;
                opList.add(op)
                i += 1

            opPath = opFolder + "/" + opName + fileExt
            retries = docSettings["postRetries"]
            delay = docSettings["initialDelay"]
            while True:
                try:
                    if not cam.postProcess(opList, postInput):
                        retVal = "Fusion 360 reported an error processing operation"
                        if (opHasTool != None):
                            retVal += ": " +  opHasTool.name
                        return retVal
                except Exception as exc:
                    retVal += ": " + str(exc)
                    return retVal

                time.sleep(delay) # wait for it to finish (??)
                try:
                    fileOp = open(opPath, encoding="utf8", errors='replace')
                    break
                except:
                    delay *= 2
                    retries -= 1
                    if retries > 0:
                        continue;
                    # Maybe the file name extension is wrong
                    for file in os.listdir(opFolder):
                        if file.startswith(opName):
                            ext = file[len(opName):]
                            if ext != fileExt:
                                ui.messageBox("Unable to open output file. "
                                    "Found the file with extension '{}' instead "
                                    "of '{}'. Make sure you have the correct file "
                                    "extension set in the Post Process All "
                                    "dialog.".format(ext, fileExt))
                            break;
                    return "Unable to open " + opPath
            
            # Parse the gcode. We expect a header like this:
            #
            # % <optional>
            # (<comments>) <0 or more lines>
            # (<Txx tool comment>) <optional>
            # <comments or G-code initialization, up to Txx>
            #
            # This header is stripped from all files after the first,
            # except the tool comment is put in a list at the top.
            # The header ends when we find the body, which starts with:
            #
            # Txx ...   (optionally preceded by line number Nxx)
            #
            # We copy all the body, looking for the tail. The start
            # of the tail is marked by one of these:
            # M30 - end program
            # M5 - stop spindle
            # M9 - stop coolant
            # The tail is stripped until the last operation is done.

            # Space between operations
            if not fFirst:
                fileBody.write("\n")

            # % at start only
            line = fileOp.readline()
            if line[0] == "%":
                if fFirst:
                    fileHead.write(line)
                line = fileOp.readline()

            # check for initial comments and tool
            # send it to header
            while line[0] == "(" or line[0] == "O" or line[0] == "\n":
                if regToolComment.match(line) != None:
                    fileHead.write(line)
                    line = fileOp.readline()
                    break;

                if fFirst:
                    pos = line.upper().find(opName.upper())
                    if pos != -1:
                        pos += len(opName)
                        if docSettings["numericName"]:
                            fill = "0" * (pos - len(fname) - 1)
                        else:
                            fill = ""
                        line = line[0] + fill + fname + line[pos:]    # correct file name
                    fileHead.write(line)
                line = fileOp.readline()

            # Body starts at tool code, T
            # Grab preceding comment if present
            fBody = False
            while True:
                match = regBody.match(line).groupdict();
                line = match["line"]        # filter off line number if present
                fNum = match["N"] != None
                if (fBody):
                    break
                toolCur = match["T"]
                if (toolCur != None):
                    toolCur = int(toolCur)
                    if not fFirst:
                        # Is this a tool change?
                        if toolCur != toolLast:
                            if len(toolChange) != 0:
                                if fToolChangeNum:
                                    # Add line number to tool change
                                    for code in toolChange:
                                        fileBody.write("N" + str(lineNum) + " " + code)
                                        lineNum += constLineNumInc
                                else:
                                    fileBody.write(toolChange)
                        else:
                            fBody = True
                            line = fileOp.readline()
                            continue    # don't output tool line
                    toolLast = toolCur
                    break

                if fFirst or line[0] == "(":
                    if (fNum):
                        fileBody.write("N" + str(lineNum) + " ")
                        lineNum += constLineNumInc
                    fileBody.write(line)
                line = fileOp.readline()
                if len(line) == 0:
                    return "Tool change G-code (Txx) not found; this post processor is not compatible with Post Process All."

            # We're done with the head, move on to the body
            # Initialize rapid move optimizations
            fFastZ = fFastZenabled
            Gcode = None
            Zcur = None
            Zfeed = None
            fZfeedNotSet = True
            feedCur = 0
            fFirstG1 = False
            fLockSpeed = False

            # Note that match, line, and fNum are already set
            while True:
                # End of program marker?
                Mtmp = match["M"]
                if Mtmp != None:
                    Mtmp = int(Mtmp)
                    if Mtmp in constEndMcodeSet:
                        break
                    # When M49/M48 is used to turn off speed changes, disable fast moves as well
                    if Mtmp == 49:
                        fLockSpeed = True
                    elif Mtmp == 48:
                        fLockSpeed = False

                if docSettings["homeEndsOp"] and ("G28" in line or "G30" in line):
                    break

                if fFastZ:
                    # Analyze code for chances to make rapid moves
                    match = regGcode.match(line)
                    if match.end() != 0:
                        try:
                            match = match.groupdict()
                            GcodeTmp = match["G"]
                            if GcodeTmp != None:
                                GcodeTmp = int(float(GcodeTmp))
                                if GcodeTmp in constMotionGcodeSet:
                                    Gcode = GcodeTmp
                                    if Gcode != 1:
                                        fFirstG1 = False

                            Ztmp = match["Z"]
                            if Ztmp != None:
                                Zlast = Zcur
                                Zcur = float(Ztmp)

                            feedTmp = match["F"]
                            if feedTmp != None:
                                feedCur = float(feedTmp)

                            XYcur = match["XY"].rstrip("\n ")

                            if (Zfeed == None or fZfeedNotSet) and (Gcode == 0 or Gcode == 1) and Ztmp != None and len(XYcur) == 0:
                                # Figure out Z feed
                                if (Zfeed != None):
                                    fZfeedNotSet = False
                                Zfeed = Zcur
                                if Gcode != 0:
                                    # Replace line with rapid move
                                    line = constRapidZgcode.format(Zcur, line[:-1])
                                    fFirstG1 = True
                                    Gcode = 0

                            if Gcode == 1 and not fLockSpeed:
                                if Ztmp != None:
                                    if len(XYcur) == 0 and (Zcur >= Zlast or Zcur >= Zfeed or feedCur == 0):
                                        # Upward move, above feed height, or anomalous feed rate.
                                        # Replace with rapid move
                                        line = constRapidZgcode.format(Zcur, line[:-1])
                                        fFirstG1 = True
                                        Gcode = 0

                                elif Zcur >= Zfeed:
                                    # No Z move, at/above feed height
                                    line = constRapidXYgcode.format(XYcur, line[:-1])
                                    fFirstG1 = True
                                    Gcode = 0

                            elif fFirstG1 and GcodeTmp == None:
                                # No G-code present, changing to G1
                                if Ztmp != None:
                                    if len(XYcur) != 0:
                                        # Not Z move only - back to G1
                                        line = constFeedXYZgcode.format(XYcur, Zcur, feedCur, line[:-1])
                                        fFirstG1 = False
                                        Gcode = 1
                                    elif Zcur < Zfeed and Zcur <= Zlast:
                                        # Not up nor above feed height - back to G1
                                        line = constFeedZgcode.format(Zcur, feedCur, line[:-1])
                                        fFirstG1 = False
                                        Gcode = 1
                                        
                                elif len(XYcur) != 0 and Zcur < Zfeed:
                                    # No Z move, below feed height - back to G1
                                    line = constFeedXYgcode.format(XYcur, feedCur, line[:-1])
                                    fFirstG1 = False
                                    Gcode = 1

                            if (Gcode == 1 and fFirstG1):
                                if (feedTmp == None):
                                    # Feed rate not present, add it
                                    line = line[:-1] + constAddFeedGcode.format(feedCur)
                                fFirstG1 = False

                            if Zcur != None and Zfeed != None and Zcur > Zfeed and Gcode != None and \
                                Gcode != 0 and len(XYcur) != 0 and (Ztmp != None or Gcode != 1):
                                # We're above the feed height, but made a cutting move.
                                # Feed height is wrong, bring it up
                                Zfeed = Zcur + 0.001
                        except:
                            fFastZ = False # Just skip changes

                # copy line to output
                if (fNum):
                    fileBody.write("N" + str(lineNum) + " ")
                    lineNum += constLineNumInc
                fileBody.write(line)
                lineFull = fileOp.readline()
                if len(lineFull) == 0:
                    break
                match = regBody.match(lineFull).groupdict()
                line = match["line"]        # filter off line number if present
                fNum = match["N"] != None

            # Found tail of program
            if fFirst:
                tailGcode = lineFull + fileOp.read()
            fFirst = False
            fileOp.close()
            os.remove(fileOp.name)
            fileOp = None

        # Completed all operations, add tail
        # Update line numbers if present
        if len(tailGcode) != 0:
            tailGcode = tailGcode.splitlines(True);
            for code in tailGcode:
                match = regBody.match(code).groupdict()
                if match["N"] != None:
                    fileBody.write("N" + str(lineNum) + " " + match["line"])
                    lineNum += constLineNumInc
                else:
                    fileBody.write(code)

        # Copy body to head
        fileBody.close()
        fileBody = open(fileBody.name)  # open for reading
        # copy in chunks
        while True:
            block = fileBody.read(10240)
            if len(block) == 0:
                break
            fileHead.write(block)
            block = None    # free memory
        fileBody.close()
        os.remove(fileBody.name)
        fileBody = None
        fileHead.close()
        fileHead = None

        return None

    except:
        if fileHead:
            try:
                fileHead.close()
                os.remove(fileHead.name)
            except:
                pass

        if fileBody:
            try:
                fileBody.close()
                os.remove(fileBody.name)
            except:
                pass

        if fileOp:
            try:
                fileOp.close()
                os.remove(fileOp.name)
            except:
                pass

        if ui:
            retVal += " " + traceback.format_exc()

        return retVal
