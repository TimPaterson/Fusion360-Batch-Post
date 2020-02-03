#Author-Tim Paterson
#Description-Post process all CAM setups, using the setup name as the output file name.

import adsk.core, adsk.fusion, adsk.cam, traceback, shutil, json, os, os.path, time

# Version number of settings as saved in documents and settings file
# update this whenever settings content changes
version = 2

# Initial default values of settings
defaultSettings = {
    "version" : version,
    "post" : "",
    "units" : adsk.cam.PostOutputUnitOptions.DocumentUnitsOutput,
    "output" : "",
    "sequence" : True,
    "twoDigits" : False,
    "delFiles" : False,
    "delFolder" : False
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
constGcodeFileExt = ".nc"

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


def CountOutputFolderFiles(folder, limit):
    cntFiles = 0
    cntNcFiles = 0
    for path, dirs, files in os.walk(folder):
        for file in files:
            if file.endswith(constGcodeFileExt):
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
            inputGroupPost = inputs.addGroupCommandInput("groupOutput", "Output Folder")
            input = inputGroupPost.children.addStringValueInput("output", "", docSettings["output"])
            input.isFullWidth = True
            input.tooltip = "Output Folder"
            input.tooltipDescription = (
                "Full path name of the output folder. Any subfolders, as denoted "
                "by colons in the setup name, will be relative this folder.")

            input = inputGroupPost.children.addBoolValueInput("browseOutput", "Browse", False)
            input.resourceFolder = "resources/Browse"
            input.tooltip = "Browse for Output Folder"

            # check box to delete existing files
            input = inputs.addBoolValueInput("delFiles", 
                                             "Delete existing files in each folder", 
                                             True, 
                                             "", 
                                             docSettings["delFiles"])
            input.tooltip = "Delete Existing Files"
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
                                             "Delete entire output folder first", 
                                             True, 
                                             "", 
                                             docSettings["delFolder"] and docSettings["delFiles"])
            input.isEnabled = docSettings["delFiles"] # enable only if delete existing files
            input.tooltip = "Delete Output Folder"
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
                                             "Prepend sequence number to name", 
                                             True, 
                                             "", 
                                             docSettings["sequence"])
            input.tooltip = "Add Sequence Numbers"
            input.tooltipDescription = (
                "Begin each file name with a sequence number. The numbering "
                "represents the order that the setups appear in the browser tree. "
                "Each folder has its own sequence numbers starting with 1.")

            # check box to select 2-digit sequence numbers
            input = inputs.addBoolValueInput("twoDigits", 
                                             "Use 2-digit sequence numbers", 
                                             True, 
                                             "", 
                                             docSettings["twoDigits"])
            input.isEnabled = docSettings["sequence"] # enable only if using sequence numbers
            input.tooltip = "Use 2-Digit Sequence Numbers"
            input.tooltipDescription = (
                "Sequence numbers 0 - 9 will have a leading zero added, becoming"
                '"01" to "09". This could be useful for formatting or sorting.')

            # post processor
            inputGroupPost = inputs.addGroupCommandInput("groupPost", "Post Processor")
            input = inputGroupPost.children.addStringValueInput("post", "", docSettings["post"])
            input.isFullWidth = True
            input.tooltip = "Post Processor"
            input.tooltipDescription = (
                "Full path name of the post processor (.cps file).")
            
            input = inputGroupPost.children.addBoolValueInput("browsePost", "Browse", False)
            input.resourceFolder = "resources/Browse"
            input.tooltip = "Browse for Post Processor"
            if (len(docSettings["post"]) != 0):
                inputGroupPost.isExpanded = False

            # button to save default settings
            input = inputs.addBoolValueInput("save", "Save these settings as system default", False)
            input.resourceFolder = "resources/Save"
            input.tooltip = "Save Default Settings"
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

            elif input.id in self.docSettings:
                self.docSettings[input.id] = input.value

            # Enable twoDigits only if sequence is true
            if input.id == "sequence":
                inputs.itemById("twoDigits").isEnabled = input.value

            # Enable delFolder only if delFiles is true
            if input.id == "delFiles":
                item = inputs.itemById("delFolder")
                item.value = input.value and item.value
                item.isEnabled = input.value

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
                    strMsg = CountOutputFolderFiles(outputFolder, setups.count)
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
                        if not cam.checkToolpath(setup):
                            cntSkipped += 1
                            lstSkipped += "\n"+ setup.name
                        else:
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
                            if docSettings["sequence"]:
                                seq = seqDict[setupFolder]
                                seqStr = str(seq)
                                if docSettings["twoDigits"] and seq < 10:
                                    seqStr = "0" + seqStr
                                fname = seqStr + ' ' + fname

                            # post the file
                            postInput = adsk.cam.PostProcessInput.create(fname, 
                                                                        docSettings["post"], 
                                                                        setupFolder, 
                                                                        docSettings["units"])
                            postInput.isOpenInEditor = False
                            try:
                                if not cam.postProcess(setup, postInput):
                                    cntSkipped += 1
                                    lstSkipped += "\Failed: "+ setup.name
                                else:
                                    cntFiles += 1
                                time.sleep(.05) # files missing on big projects unless we slow down (??)
                            except:
                                cntSkipped += 1
                                lstSkipped += "\nException: "+ setup.name
                        
                    cntSetups += 1
                    progress.message = progressMsg.format(cntFiles)
                    progress.progressValue = cntSetups

            progress.hide()

        # done with setups, report results
        if cntSkipped != 0:
            ui.messageBox("{} files were written. {} Setups were skipped due to invalid toolpath:{}".format(cntFiles, cntSkipped, lstSkipped), 
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
