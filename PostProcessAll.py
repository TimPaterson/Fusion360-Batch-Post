#Author-Tim Paterson
#Description-Post process all CAM setups, using the setup name as the output file name.

import adsk.core, adsk.fusion, adsk.cam, traceback, shutil, json, os, os.path, time, re, pathlib, enum

# Version number of settings as saved in documents and settings file
# update this whenever settings content changes
version = 4

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
    # Groups are expanded or not
    "groupOutput" : True,
    "groupPersonal" : True,
    "groupPost" : True
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
constPostLoopDelay = 0.1
constBodyTmpFile = "$body"
constOpTmpFile = "$op"
constToolChangeGcode = "G30\n"
constEndProgramGcode = "G30\nM30\n%\n"
constRapidZgcode = 'G00 Z{} (Changed from: "{}")\n'
constRapidXYgcode = 'G00 {} (Changed from: "{}")\n'
constAddFeedGcode = " F{} (Feed rate added)\n"

# Errors in post processing
class PostError(enum.Enum):
    Success = 0
    Fail = 1
    Except = 2
    BadFormat = 3

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

            # "Personal Use" version
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

            input = inputGroup.children.addBoolValueInput("fastZ",
                                                              "Make rapid Z moves",
                                                              True,
                                                              "",
                                                              docSettings["fastZ"])
            input.isEnabled = docSettings["splitSetup"] # enable only if using individual operations
            input.tooltip = "Make Initial Z Moves Rapid"
            input.tooltipDescription = (
                "Replace the initial Z moves at feed rate with rapid (G0) moves. "
                "In Fusion 360 for Personal Use, moves that could be rapid are "
                "now limited to the current feed rate. When this optionis selected, "
                "the G-code will be analyzed to find the initial Z moves and "
                "replace them with rapid moves."
                "<p><b>WARNING!<b> This option should be used with caution. "
                "Review the G-code to verify it is correct. Comments have been "
                "added to indicate the changes.")
            inputGroup.isExpanded = docSettings["groupPersonal"]

            # post processor
            inputGroup = inputs.addGroupCommandInput("groupPost", "Post Processor")
            input = inputGroup.children.addStringValueInput("post", "", docSettings["post"])
            input.isFullWidth = True
            input.tooltip = "Post Processor"
            input.tooltipDescription = (
                "Full path name of the post processor (.cps file).")
            
            input = inputGroup.children.addBoolValueInput("browsePost", "Browse", False)
            input.resourceFolder = "resources/Browse"
            input.tooltip = "Browse for Post Processor"
            inputGroup.isExpanded = docSettings["groupPost"]

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

            # Enable fastZ only if splitSetup is true
            if input.id == "splitSetup":
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
                            lstSkipped += "\n" + setup.name
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
                            status = PostProcessSetup(fname, setup, setupFolder, docSettings)
                            if status == PostError.Success:
                                cntFiles += 1
                            else:
                                cntSkipped += 1
                                if status == PostError.Fail:
                                    lstSkipped += "\nFailed: "
                                elif status == PostError.BadFormat:
                                    lstSkipped += "\nGcode file format not recognized: "
                                else:
                                    lstSkipped += "\nException: "
                                lstSkipped += setup.name
                         
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

class FileFormatError(Exception):
    pass

def PostProcessSetup(fname, setup, setupFolder, docSettings):
    ui = None
    fileHead = None
    fileBody = None
    fileOp = None
    retVal = PostError.Except

    try:
        app = adsk.core.Application.get()
        ui  = app.userInterface
        doc = app.activeDocument
        product = doc.products.itemByProductType(constCAMProductId)
        cam = adsk.cam.CAM.cast(product)

        # Create PostProcessInput
        opName = fname
        if docSettings["splitSetup"]:
            opName += constOpTmpFile
        postInput = adsk.cam.PostProcessInput.create(opName, 
                                                    docSettings["post"], 
                                                    setupFolder, 
                                                    docSettings["units"])
        postInput.isOpenInEditor = False

        # Do it all at once?
        if not docSettings["splitSetup"]:
            try:
                if not cam.postProcess(setup, postInput):
                    return PostError.Fail
                time.sleep(constPostLoopDelay) # files missing sometimes unless we slow down (??)
                return PostError.Success
            except:
                return PostError.Except

        # Split setup into individual operations
        path = setupFolder + "/" + fname
        pathlib.Path(setupFolder).mkdir(parents=True, exist_ok=True)
        fileHead = open(path + constGcodeFileExt, "w")
        fileBody = open(path + constBodyTmpFile + constGcodeFileExt, "w")
        fFirst = True
        lineNum = 10
        toolLast = -1
        fFastZenabled = docSettings["fastZ"]
        if fFastZenabled:
            strReg = (r"(G(?P<G>[0-9]+)[^XYZF]*)?"
                "(?P<XY>((X-?[0-9]+(\.[0-9]*)?)[^XYZF]*)?"
                "((Y-?[0-9]+(\.[0-9]*)?)[^XYZF]*)?)"
                "(Z(?P<Z>-?[0-9]+(\.[0-9]*)?)[^XYZF]*)?"
                "(F(?P<F>-?[0-9]+(\.[0-9]*)?)[^XYZF]*)?")
            regGcode = re.compile(strReg)

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
            hasTool = op.hasToolpath
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
                        i += 1
                    break;
                opList.add(op)
                i += 1

            retries = 5
            delay = constPostLoopDelay
            while True:
                try:
                    if not cam.postProcess(opList, postInput):
                        return PostError.Fail
                except:
                    return PostError.Except

                time.sleep(delay) # wait for it to finish (??)
                try:
                    fileOp = open(path + constOpTmpFile + constGcodeFileExt)
                    break
                except:
                    delay *= 2
                    retries -= 1
                    if retries == 0:
                        raise
                    pass
            
            # Parse the gcode. we want to move each tool comment to the top.
            # We expect this format:
            # %
            # (<file name>)
            # (<tool>)
            # <several lines of initialization or comments
            # Nxx(<op name>)
            # Txx ...
            # ...
            # Nxx(<op name>) <if pattern>
            # ...
            # G30
            # M30
            # %

            # % at start only
            line = fileOp.readline()
            if line[0] != "%":
                raise FileFormatError
            if fFirst:
                fileHead.write(line)

            # filaname comment at start only
            line = fileOp.readline()
            if line[0] != "(":
                raise FileFormatError
            if fFirst:
                fileHead.write("(" + fname + ")\n")

            # Now get the tool comment
            line = fileOp.readline()
            fileHead.write(line)
            if line[0:2] != "(T":
                raise FileFormatError

            # We're done with the head, move on to the body
            # Skip to the first line number: "N"
            fHavBody = False
            line2 = ""
            line1 = ""
            # Initialize rapid move optimizations
            fFastZ = fFastZenabled
            Zcur = None
            Zfeed = None
            feedCur = 0
            fFirstG1 = False

            while True:
                line = fileOp.readline()
                if len(line) == 0:
                    raise FileFormatError;  #unexpected EOF

                ch = line[0]

                # Have a line number?
                if ch == "N":
                    pos = line.find("(")
                    line = "N" + str(lineNum) + (line[pos:] if pos != -1 else "\n")
                    lineNum += 10
                    fHavBody = True

                # Have a tool change?
                elif ch == "T":
                    pos = line.find(" ")
                    if pos == -1:
                        raise FileFormatError
                    try:
                        toolCur = int(line[1:pos])
                    except:
                        raise FileFormatError
                    # Is this a tool change?
                    if toolCur != toolLast and toolLast != -1:
                        fileBody.write(constToolChangeGcode)
                    toolLast = toolCur
                    fHavBody = True

                # End of program marker?
                elif ch == "%":
                    break

                elif fFastZ:
                    # Looking in previous line for a Z move
                    match = regGcode.match(line1)
                    if match.end() != 0:
                        try:
                            match = match.groupdict()
                            Gcode = match["G"]
                            if Gcode != None:
                                Gcode = int(Gcode)

                            Ztmp = match["Z"]
                            if Ztmp != None:
                                Zlast = Zcur
                                Zcur = float(Ztmp)

                            feedTmp = match["F"]
                            if feedTmp != None:
                                feedCur = float(feedTmp)

                            if Gcode == 1:
                                if Ztmp != None:
                                    if len(match["XY"]) == 0:
                                        # Z move only
                                        if Zfeed == None:
                                            # Have first Z-only move, make it rapid
                                            Zfeed = Zcur
                                            # Replace line with rapid move
                                            line1 = constRapidZgcode.format(Zcur, line1[:-1])
                                            fFirstG1 = True
                                            Gcode = 0

                                            # First move was retract height, check for second move to feed height
                                            match = regGcode.match(line)
                                            if match.end() != 0:
                                                match = match.groupdict()
                                                Ztmp = match["Z"]                                                
                                                if Ztmp != None and len(match["XY"]) == 0:
                                                    # Assume this is feed height. This is wrong if threading/boring
                                                    # from the bottom of a hole
                                                    Zfeed = float(Ztmp)

                                        elif Zcur >= Zlast or Zcur >= Zfeed or feedCur == 0:
                                            # Upward move, above feed height, or anomalous feed rate.
                                            # Replace with rapid move
                                            line1 = constRapidZgcode.format(Zcur, line1[:-1])
                                            fFirstG1 = True
                                            Gcode = 0

                                elif Zcur >= Zfeed:
                                    # No Z move, at/above feed height
                                    line1 = constRapidXYgcode.format(match["XY"].rstrip("\n "), line1[:-1])
                                    fFirstG1 = True
                                    Gcode = 0

                            if (Gcode == 1 and fFirstG1):
                                if (feedTmp == None):
                                    # Feed rate not present, add it
                                    line1 = line1[:-1] + constAddFeedGcode.format(feedCur)
                                fFirstG1 = False

                            if Zcur != None and Zfeed != None and Zcur > Zfeed and Gcode != None and \
                                Gcode != 0 and len(match["XY"]) != 0 and (Ztmp != None or Gcode != 1):
                                # We're above the feed height, but made a cutting move.
                                # Feed height is wrong, bring it up
                                Zfeed = Zcur + 0.001
                        except:
                            fFastZ = False # Just skip changes

                # copy line to output
                if fFirst or fHavBody:
                    fileBody.write(line2)
                    line2 = line1
                    line1 = line

            fFirst = False
            fileOp.close()
            os.remove(fileOp.name)
            fileOp = None

        # Completed all operations
        # Copy body to head
        fileBody.close()
        fileBody = open(fileBody.name)  # open for reading
        fileHead.write(fileBody.read())
        fileBody.close()
        os.remove(fileBody.name)
        fileBody = None

        # Append the ending line
        fileHead.write(constEndProgramGcode)
        fileHead.close()
        fileHead = None

        return PostError.Success

    except FileFormatError:
        retVal = PostError.BadFormat
        # Fall into all other errors
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

        if ui and retVal != PostError.BadFormat:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

        return retVal
