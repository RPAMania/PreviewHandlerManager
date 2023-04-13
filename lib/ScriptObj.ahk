/**
 * ============================================================================ *
 * @Author           : RaptorX <graptorx@gmail.com>, modified by TJay🐦
 * @Script Name      : Script Object
 * @Script Version   : 0.20.2
 * @Homepage         :
 *
 * @Creation Date    : November 09, 2020
 * @Modification Date: July 02, 2021
 *
 * @Description      :
 * -------------------
 * This is an object used to have a few common functions between scripts
 * Those are functions and variables related to basic script information,
 * upgrade and configuration.
 *
 * ============================================================================ *
 */

; global script := {base         : script
;                  ,name          : regexreplace(A_ScriptName, "\.\w+")
;                  ,version      : "0.1.0"
;                  ,author       : ""
;                  ,email        : ""
;                  ,crtdate      : ""
;                  ,moddate      : ""
;                  ,homepagetext : ""
;                  ,homepagelink : ""
;                  ,donateLink   : "https://www.paypal.com/donate?hosted_button_id=MBT5HSD9G94N6"
;                  ,resfolder    : A_ScriptDir "\res"
;                  ,iconfile     : A_ScriptDir "\res\sct.ico"
;                  ,configfile   : A_ScriptDir "\settings.ini"
;                  ,configfolder : A_ScriptDir ""}
#Requires AutoHotkey v2.0-
class script
{
  static DBG_NONE     := 0
        ,DBG_ERRORS   := 1
        ,DBG_WARNINGS := 2
        ,DBG_VERBOSE  := 3

  static name         := ""
        ,version      := ""
        ,author       := ""
        ,email        := ""
        ,crtdate      := ""
        ,moddate      := ""
        ,homepagetext := ""
        ,homepagelink := ""
        ,resfolder    := ""
        ,icon         := ""
        ,config       := ""
        ,systemID     := ""
        ,dbgFile      := ""
        ,dbgLevel     := this.DBG_NONE


  /**
    Function: Update
    Checks for the current script version
    Downloads the remote version information
    Compares and automatically downloads the new script file and reloads the script.

    Parameters:
    vfile - Version File
            Remote version file to be validated against.
    rfile - Remote File
            Script file to be downloaded and installed if a new version is found.
            Should be a zip file that will be unzipped by the function

    Notes:
    The versioning file should only contain a version string and nothing else.
    The matching will be performed against a SemVer format and only the three
    major components will be taken into account.

    e.g. '1.0.0'

    For more information about SemVer and its specs click here: <https://semver.org/>
  */
  static Update(vfile, rfile := "")
  {
    ; Error Codes
    static ERR_INVALIDVFILE := 1
    ,ERR_INVALIDRFILE       := 2
    ,ERR_NOCONNECT          := 3
    ,ERR_NORESPONSE         := 4
    ,ERR_INVALIDVER         := 5
    ,ERR_CURRENTVER         := 6
    ,ERR_MSGTIMEOUT         := 7
    ,ERR_USRCANCEL          := 8

    ; IWinHttpRequest COM object's WinHttpRequestOption for retrieving the URL from a response
    static WinHttpRequestOption_URL := 1
    
    if (rfile == "")
    {
      rfile := vfile
    }

    ; A URL is expected in this parameter, we just perform a basic check
    if (!vfile ~= "i)^((?:http(?:s)?|ftp):\/\/)?((?:[a-z0-9_\-]+\.)+.*$)")
      throw {code: ERR_INVALIDVFILE, msg: "Invalid URL`n`n"
          . "The version file parameter must point to a valid URL."}

    ; This function expects a ZIP file or a GitHub URL ending "/latest" or "/tag/v[versionNumber]"
    if (!rfile ~= "i)\.zip" && 
        !rfile ~= "i)^(?:http(?:s)?:\/\/)?github.com/.*/releases/(latest|tag/v(\d\.){2}\d)$")
      throw {code: ERR_INVALIDRFILE, msg: "Invalid Zip or URL to a GitHub release page`n`n"
          . "The remote file parameter must either point to "
          . "a zip file or be a URL to a page of a specific GitHub release version." }

    ; Check if we are connected to the internet
    http := ComObject("WinHttp.WinHttpRequest.5.1")
    http.Open("GET", "https://www.google.com", true)
    http.Send()
    try
      http.WaitForResponse(1)
    catch Any as e
      throw {code: ERR_NOCONNECT, msg: e.message}

    progressGui := Gui("-Caption +Border", "Updating")
    progressGui.SetFont("bold")
    progressGui.Add("Text", "w250", "Checking for updates")
    progressGui.SetFont("norm")
    progressGui.Add("Progress", "wp h20 vprogressPercentage", 50)
    progressGui.Add("Text", "wp Center vprogressText", "50/100")
    progressGui.Show()

    ; Download remote version file
    http.Open("GET", vfile, true)
    http.Send(), http.WaitForResponse()

    if (!http.responseText)
    {
      progressGui.Destroy()
      throw {code: ERR_NORESPONSE
          , msg: "There was an error trying to download the version info.`n"
              . "The server did not respond."}
    }

    if (http.Status == 404)
    {
      progressGui.Destroy()
      throw {code: ERR_INVALIDVFILE
          , msg: "Version file was not found via the given url." }
    }

    regexmatch(this.version, "\d+\.\d+\.\d+", &loVersion)
    regexmatch(http.responseText, "\d+\.\d+\.\d+", &remVersion)

    progressGui["progressPercentage"].Value := 100
    progressGui["progressText"].Text := "100/100"
    sleep 500 	; allow progress to update

    progressGui.Destroy()

    ; Make sure SemVer is used
    if (!loVersion || !remVersion)
      throw {code: ERR_INVALIDVER, msg: "Invalid version.`nThis function works with SemVer. "
                      . "For more information refer to the documentation in the function"}

    ; Compare against current stated version
    ver1 := strsplit(loVersion.0, ".")
    ver2 := strsplit(remVersion.0, ".")

    for i1,num1 in ver1
    {
      for i2,num2 in ver2
      {
        if (i1 == i2)
          if (num2 > num1)
          {
            newVersion := true
            break 2
          }
          else if (num2 < num1)
          {
            newVersion := false
            break 2
          }
          else
          {
            newVersion := false
          }
      }
    }

    if (!newVersion)
      throw {code: ERR_CURRENTVER, msg: "You are using the latest version"}
    else
    {
      ; If new version ask user what to do
      ; Yes/No | Icon Question | System Modal
      msgboxResult := msgbox("There is a new update available for this application.`n"
          "Do you wish to upgrade to v" remVersion.0 "?",
          "New Update Available",
          (0x4 + 0x20 + 0x1000) " T10")

      if (msgboxResult == "Timeout")
        throw {code: ERR_MSGTIMEOUT, msg: "The Message Box timed out."}
      if (msgboxResult == "No")
        throw {code: ERR_USRCANCEL, msg: "The user pressed the cancel button."}

      ; Create temporal dirs
      zipFileContainerDirCandidates := []
      if (InStr(rfile, "github"))
      {
        zipFileContainerDirCandidates.Push(
            regexreplace(a_scriptname, "\..*$") "-" remVersion.0)
      }

      DirCreate(tmpDir := a_temp "\" regexreplace(a_scriptname, "\..*$"))
      try
      {
        DirDelete(zipDir := tmpDir "\uzip", true)
      }
      DirCreate(zipDir := tmpDir "\uzip")

      ; Create lock file
      fileappend(a_now, lockFile := tmpDir "\lock")

      ; Download zip file
      try
      {
        try
        {
          FileDelete(tmpDir "\temp.zip")
        }
        
        if (rfile ~= "i)\.zip$")
        {
          ; .zip location given directly
          
          zipFileDownloadUrl := rfile
        }
        else
        {
          ; GitHub release page URL given

          gitReleasePageUrl := rfile ~= "i)latest$" 
              ; From  https://github.com/[Account]/[project]/releases/latest
              ; to    https://github.com/[Account]/[project]/releases/tag/v[latestVersionNumber]
              ? http.Option(WinHttpRequestOption_URL)
              : rFile
          
          ; https://github.com/[Account]/[project]/releases/expanded_assets/v[versionNumber]
          gitDownloadPageUrl := strreplace(gitReleasePageUrl, "/tag/", "/expanded_assets/")
          http.Open("GET", gitDownloadPageUrl, true)
          http.Send()
          http.WaitForResponse()
          if (http.Status != 200)
          {
            throw {code: ERR_NORESPONSE, msg: Format("There was an error trying to retrieve a "
                . ".zip file download link from a GitHub page.`n`n"
                . "Status code: {1}`nServer response: {2}", http.Status, http.responseText)}
          }

          if (a_iscompiled)
          {
            ; Pinpoint download URL of the .zip file containing the executable in GitHub
            ; /[Account]/[Project]/releases/download/v[versionNumber]/[filenameWithoutExtension]_[versionNumber].zip
            regexmatch(http.responseText, "(?<=<a href=`").*?" 
                PreviewHandlerManager.Prototype.__Class "(?:_" remVersion.0 ")?\.zip(?=`" )", 
                &urlWithoutDomain)
          }
          else
          {
            ; Pinpoint download URL of the .zip file containing script source code in GitHub
            ; /[Account]/[Project]/archive/refs/tags/v[versionNumber].zip
            regexmatch(http.responseText, "(?<=<a href=`").*?" remVersion.0 "\.zip(?=`" )", 
                &urlWithoutDomain)
          }

          ; https://github.com/[Account]/[Project]/releases/download/v[versionNumber]/[filenameWithoutExtension]_[versionNumber].zip
          ; OR
          ; https://github.com/[Account]/[Project]/archive/refs/tags/v[versionNumber].zip
          zipFileDownloadUrl := Format("{1}{2}", 
              regexreplace(rfile, "^(.*?)(?<![:\/])(?=\/).*", "$1"), urlWithoutDomain.0)
        }

        Download(zipFileDownloadUrl, tmpDir "\temp.zip")
      }
      catch Error as e
      {
        if (!FileExist(tmpDir "\temp.zip"))
        {
          throw {code: ERR_NORESPONSE
              , msg: "There was an error trying to download the .zip file.`n"
                  . "The server did not respond."}
        }
      }
      

      ; Extract zip file to temporal folder
      oShell := ComObject("Shell.Application")
      oDir := oShell.NameSpace(zipDir), oZip := oShell.NameSpace(tmpDir "\temp.zip")
      oDir.CopyHere(oZip.Items), oShell := oDir := oZip := ""

      filedelete(tmpDir "\temp.zip")

      /*
      ******************************************************
      * Wait for lock file to be released
      * Copy all files to current script directory
      * Cleanup temporal files
      * Run main script
      * EOF
      *******************************************************
      */

      if (!a_iscompiled && zipFileContainerDirCandidates.Length)
      {
        ; Validate container subdirectory inside the .zip file that contains sources

        containerCandidateDirsString := ""
        for dirCandidate in zipFileContainerDirCandidates
        {
          ; Build comma-separated string for logging
          containerCandidateDirsString .= "> " dirCandidate ", "
          
          if (instr(fileexist(zipDir "\" dirCandidate), "D"))
          {
            zipFileContainerDir := dirCandidate "\"
            break
          }
        }

        
        if (!isset(zipFileContainerDir))
        {
          ; Prepare and throw the error

          containerCandidateDirsString := SubStr(containerCandidateDirsString, 1, -2)

          ; Build comma-separated string for logging
          actualContainerDirsString := "> [root], "
          loop files, zipDir "\*.*", "D"
          {
            actualContainerDirsString .= "> " a_loopfilename ", "
          }

          actualContainerDirsString := SubStr(actualContainerDirsString, 1, -2)

          throw UnsetError(Format("A container subdirectory inside the .zip file downloaded "
              "from GitHub was not`nrecognized among preset subdirectories.`n`nFollowing "
              "presets were searched in '{1}':`n{2}`n`nFollowing subdirectories were "
              "found:`n{3}",
              zipDir, containerCandidateDirsString, actualContainerDirsString))
        }
      }
      else
      {
        ; .zip file contains a naked executable, no container subdirectory expected
        zipFileContainerDir := ""
      }

      if (a_iscompiled)
      {
        tmpBatch :=
        (Ltrim
          ":lock
          if not exist `"" lockFile "`" goto continue
          timeout /t 10
          goto lock
          :continue

          xcopy `"" zipDir "\" zipFileContainerDir "*.*`" `"" a_scriptdir "\`" /E /C /I /Q /R /K /Y
          if exist `"" a_scriptfullpath "`" cmd /C `"" a_scriptfullpath "`"

          cmd /C `"rmdir `"" tmpDir "`" /S /Q`"
          exit"
        )
        
        try
        {
          filedelete(tmpDir "\update.bat")
        }
        fileappend(tmpBatch, tmpDir "\update.bat")
        run(a_comspec " /c `"" tmpDir "\update.bat`"",, "hide")
      }
      else
      {
        tmpScript :=
        (Ltrim
          "while (fileExist(`"" lockFile "`"))
            sleep 10

          DirCopy `"" zipDir "\" zipFileContainerDir "`", `"" a_scriptdir "`", true
          DirDelete `"" tmpDir "`", true

          if (fileExist(`"" a_scriptfullpath "`"))
            run `"" a_scriptfullpath "`" 
          else
            msgbox(`"There was an error while running the updated version.``n`"
                      `"Try to run the program manually.`",
                  `"Update Error`",
                  (0x10 + 0x1000) `" T10`")
            exitapp"
        )
        try
        {
          filedelete(tmpDir "\update.ahk")
        }
        fileappend(tmpScript, tmpDir "\update.ahk")
        run(a_ahkpath " " tmpDir "\update.ahk")
      }
      filedelete(lockFile)
      exitapp
    }
  }

  /**
    Function: Autostart
    This Adds the current script to the autorun section for the current
    user.

    Parameters:
    status - Autostart status
             It can be either true or false.
             Setting it to true would add the registry value.
             Setting it to false would delete an existing registry value.
  */
  /*  
  static Autostart(status)
  {
    if (status)
    {
      RegWrite(a_scriptfullpath, 
          "REG_SZ",
          "HKCU\SOFTWARE\microsoft\windows\currentversion\run",
          a_scriptname)
    }
    else
      regdelete("HKCU\SOFTWARE\microsoft\windows\currentversion\run", a_scriptname)
  }
  */
  /**
    Function: Splash
    Shows a custom image as a splash screen with a simple fading animation

    Parameters:
    img   (opt) - file to be displayed
    speed (opt) - fast the fading animation will be. Higher value is faster.
    pause (opt) - long in seconds the image will be paused after fully displayed.
  */
  /*
  static Splash(img:="", speed:=10, pause:=2)
  {
    ; static picImage

    splashGui := Gui("-caption +lastfound +border +alwaysontop +owner")
    ; gui, splash: -caption +lastfound +border +alwaysontop +owner
    alpha := 0
    WinSetTransparent(0)

    splashGui.Add("Picture", "x0 y0 vpicImage", img)
    splashGui["picImage"].GetPos(, , &picImageW, &picImageH)
    ; guicontrolget, picimage, splash:pos
    splashGui.Show("w" picImageW " h" picImageH)
    ; gui, splash: show, w%picimagew% h%picimageh%

    ; setbatchlines 3
    loop 255
    {
      if (alpha >= 255)
        break
      alpha += speed
      WinSetTransparent(alpha)
      Sleep 10
    }

    ; pause duration in seconds
    sleep pause * 1000

    loop 255
    {
      if (alpha <= 0)
        break
      alpha -= speed
      WinSetTransparent(alpha)
      Sleep 10
    }
    

    splashGui.Destroy()
  }
  */
  /**
    Funtion: Debug
    Allows sending conditional debug messages to the debugger and a log file filtered
    by the current debug level set on the object.

    Parameters:
    level - Debug Level, which can be:
            * this.DBG_NONE
            * this.DBG_ERRORS
            * this.DBG_WARNINGS
            * this.DBG_VERBOSE

    If you set the level for a particular message to *this.DBG_VERBOSE* this message
    wont be shown when the class debug level is set to lower than that (e.g. *this.DBG_WARNINGS*).

    label - Message label, mainly used to show the name of the function or label that triggered the message
    msg   - Arbitrary message that will be displayed on the debugger or logged to the log file
    vars* - Aditional parameters that whill be shown as passed. Useful to show variable contents to the debugger.

    Notes:
    The point of this function is to have all your debug messages added to your script and filter them out
    by just setting the object's dbgLevel variable once, which in turn would disable some types of messages.
  */
  /*
  static Debug(level:=1, label:=">", msg:="", vars*)
  {
    if !this.dbglevel
      return

    for var in vars
      varline .= "|" var

    dbgMessage := label ">" msg "`n" varline

    if (level <= this.dbglevel)
      outputdebug dbgMessage
    if (this.dbgFile)
      FileAppend(dbgMessage, this.dbgFile)
  }
  */
  /**
    Function: About
    Shows a quick HTML Window based on the object's variable information

    Parameters:
    scriptName   (opt) - Name of the script which will be
                         shown as the title of the window and the main header
    version      (opt) - Script Version in SimVer format, a "v"
                         will be added automatically to this value
    author       (opt) - Name of the author of the script
    homepagetext (opt) - Display text for the script website
    homepagelink (opt) - Href link to that points to the scripts
                         website (for pretty links and utm campaing codes)
    donateLink   (opt) - Link to a donation site
    email        (opt) - Developer email

    Notes:
    The function will try to infer the paramters if they are blank by checking
    the class variables if provided. This allows you to set all information once
    when instatiating the class, and the about GUI will be filled out automatically.
  */
  /*
  static About(scriptName:="", version:="", author:="", homepagetext:="", homepagelink:="", donateLink:="", email:="")
  {
    ; static doc

    scriptName := scriptName ? scriptName : this.name
    version := version ? version : this.version
    author := author ? author : this.author
    homepagetext := homepagetext ? homepagetext : RegExReplace(this.homepagetext, "http(s)?:\/\/")
    homepagelink := homepagelink ? homepagelink : RegExReplace(this.homepagelink, "http(s)?:\/\/")
    donateLink := donateLink ? donateLink : RegExReplace(this.donateLink, "http(s)?:\/\/")
    email := email ? email : this.email

    if (donateLink)
    {
      donateSection :=
      (
        "<div class=`"donate`">
          <p>If you like this tool please consider <a href=`"https://" donateLink "`">donating</a>.</p>
        </div>
        <hr>"
      )
    }

    html :=
    (
      "<!DOCTYPE html>
      <html lang=`"en`" dir=`"ltr`">
        <head>
          <meta charset=`"utf-8`">
          <meta http-equiv=`"X-UA-Compatible`" content=`"IE=edge`">
          <style media=`"screen`">
            .top {
              text-align:center;
            }
            .top h2 {
              color:#2274A5;
              margin-bottom: 5px;
            }
            .donate {
              color:#E83F6F;
              text-align:center;
              font-weight:bold;
              font-size:small;
              margin: 20px;
            }
            p {
              margin: 0px;
            }
          </style>
        </head>
        <body>
          <div class=`"top`">
            <h2>%scriptName%</h2>
            <p>v%version%</p>
            <hr>
            <p>%author%</p>
            <p><a href=`"https://" homepagelink "`" target=`"_blank`">" homepagetext "</a></p>
          </div>
          donateSection
        </body>
      </html>"
    )

    btnxPos := 300/2 - 75/2
    axHeight := donateLink ? 16 : 12

    aboutGui := Gui("+alwaysontop +toolwindow", "About " this.name)
    aboutGui.MarginX(0)
    aboutGui.BackColor("white")
    aboutGui.Add("ActiveX", "w300 r" axHeight " vdoc", "htmlFile")
    aboutGui.Add("Button", "w75 x" btnxPos, "Close").OnEvent("Click", (*) => aboutGui.Destroy())
    aboutGui["doc"].write(html)
    aboutGui.Show()
    ; gui aboutScript:new, +alwaysontop +toolwindow, % "About " this.name
    ; gui margin, 0
    ; gui color, white
    ; gui add, activex, w300 r%axHeight% vdoc, htmlFile
    ; gui add, button, w75 x%btnxPos% gaboutClose, % "Close"
    ; doc.write(html)
    ; gui show
    ; return
    
    ;aboutClose:
    ;  gui aboutScript:destroy
    ;return
  }
  */
  /*
    Function: GetLicense
    Parameters:
    Notes:
  */
  /*
  static GetLicense()
  {
    static licenseGui

    this.systemID := this.GetSystemID()
    cleanName := RegexReplace(A_ScriptName, "\..*$")
    LicenseType := License := ""

    for value in ["Type", "License"]
    {
      varNameRef := (value == "Type" ? "License" : "") value
      %varNameRef% := RegRead("HKCU\SOFTWARE\" cleanName, value)
      ; RegRead, %value%, % "HKCU\SOFTWARE\" cleanName, % value
    }

    if (!License)
    {
      msgboxResult := MsgBox("Seems like there is no license activated on this computer.`n"
                          "Do you have a license that you want to activate now?",
                      "No license",
                      0x4 + 0x20)

      if (msgboxResult == "Yes")
      {
        licenseGui := Gui("")
        licenseGui.Add("Text", "w160", "Paste the License Code here")
        licenseGui.Add("Edit", "w160 vlicenseNumber")
        licenseGui.Add("Button", "w75", "Save").OnEvent("Click", licenseButtonSave)
        licenseGui.Add("Button", "w75 x+10", "Cancel").OnEvent("Click", licenseButtonCancel)
        licenseGui.Show()

        ; Gui, license:new
        ; Gui, add, Text, w160, % "Paste the License Code here"
        ; Gui, add, Edit, w160 vLicenseNumber
        ; Gui, add, Button, w75 vTest, % "Save"
        ; Gui, add, Button, w75 x+10, % "Cancel"
        ; Gui, show

        ; saveFunction := licenseButtonSave.bind(this)
        ; GuiControl, +g, test, % saveFunction
        Exit()
      }

      licenseButtonCancel()
    }

    return { type: LicenseType, number: License }
    
    licenseButtonSave(this, licenseGui, *)
    {
      licenseNumber := licenseGui["licenseNumber"].Value

      if this.IsLicenceValid(this.eddID, licenseNumber, "https://www.the-automator.com")
      {
        this.SaveLicense(this.eddID, licenseNumber)
        MsgBox("The license was applied correctly!`n"
                    "The program will start now.",
                "License Saved",
                0x30)
        
        Reload
      }
      else
      {
        MsgBox("The license you entered is invalid and cannot be activated.",
                "Invalid License",
                0x10)

        ExitApp(1)
      }
    }

    licenseButtonCancel(*)
    {
      MsgBox("This program cannot run without a license.",
              "Unable to Run",
              0x30)

      ExitApp(1)
    }
  }
  */
  /*
    Function: SaveLicense
    Parameters:
    Notes:
  */
  /*
  static SaveLicense(licenseType, licenseNumber)
  {
    cleanName := RegexReplace(A_ScriptName, "\..*$")

    Try
    {
      RegWrite(licenseType, 
              "REG_SZ",
              "HKCU\SOFTWARE\" cleanName, 
              "Type")

      RegWrite(licenseNumber, 
              "REG_SZ",
              "HKCU\SOFTWARE\" cleanName,
              "License")

      return true
    }
    catch
      return false
  }
  */
  /*
    Function: IsLicenceValid
    Parameters:
    Notes:
  */
  /*
  static IsLicenceValid(licenseType, licenseNumber, URL)
  {
    res := this.EDDRequest(URL, "check_license", licenseType ,licenseNumber)

    if InStr(res, "`"license`":`"inactive`"")
      res := this.EDDRequest(URL, "activate_license", licenseType ,licenseNumber)

    if InStr(res, "`"license`":`"valid`"")
      return true
    else
      return false
  }
  
  static GetSystemID()
  {
    wmi := ComObjGet("winmgmts:{impersonationLevel=impersonate}!\\" A_ComputerName "\root\cimv2")
    (wmi.ExecQuery("Select * from Win32_BaseBoard")._newEnum)(&Computer)
    return Computer.SerialNumber
  }
  */
  /*
    Function: EDDRequest
    Parameters:
    Notes:
  */
  /*
  static EDDRequest(URL, Action, licenseType, licenseNumber)
  {
    strQuery := url "?edd_action=" Action
            .  "&item_id=" licenseType
            .  "&license=" licenseNumber
            .  (this.systemID ? "&url=" this.systemID : "")

    try
    {
      http := ComObject("WinHttp.WinHttpRequest.5.1")
      http.Open("GET", strQuery)
      http.SetRequestHeader("Pragma", "no-cache")
      http.SetRequestHeader("Cache-Control", "no-cache, no-store")
      http.SetRequestHeader("User-Agent", "Mozilla/4.0 (compatible; Win32)")

      http.Send()
      http.WaitForResponse()

      return http.responseText
    }
    catch Any as err
      return err.what ":`n" err.message
  }
  */
  ; Activate()
  ; 	{
  ; 	strQuery := this.strEddRootUrl . "?edd_action=activate_license&item_id=" . this.strRequestedProductId . "&license=" . this.strEddLicense . "&url=" . this.strUniqueSystemId
  ; 	strJSON := Url2Var(strQuery)
  ; 	Diag(A_ThisFunc . " strQuery", strQuery, "")
  ; 	Diag(A_ThisFunc . " strJSON", strJSON, "")
  ; 	return JSON.parse(strJSON)
  ; 	}
  ; Deactivate()
  ; 	{
  ; 	Loop, Parse, % "/|", |
  ; 	{
  ; 	strQuery := this.strEddRootUrl . "?edd_action=deactivate_license&item_id=" . this.strRequestedProductId . "&license=" . this.strEddLicense . "&url=" . this.strUniqueSystemId . A_LoopField
  ; 	strJSON := Url2Var(strQuery)
  ; 	Diag(A_ThisFunc . " strQuery", strQuery, "")
  ; 	Diag(A_ThisFunc . " strJSON", strJSON, "")
  ; 	this.oLicense := JSON.parse(strJSON)
  ; 	if (this.oLicense.success)
  ; 	break
  ; 	}
  ; 	}
  ; GetVersion()
  ; 	{
  ; 	strQuery := this.strEddRootUrl . "?edd_action=get_version&item_id=" . this.oLicense.item_id . "&license=" . this.strEddLicense . "&url=" . this.strUniqueSystemId
  ; 	strJSON := Url2Var(strQuery)
  ; 	Diag(A_ThisFunc . " strQuery", strQuery, "")
  ; 	Diag(A_ThisFunc . " strJSON", strJSON, "")
  ; 	return JSON.parse(strJSON)
  ; 	}
  ; RenewLink()
  ; 	{
  ; 	strUrl := this.strEddRootUrl . "checkout/?edd_license_key=" . this.strEddLicense . "&download_id=" . this.oLicense.item_id
  ; 	Diag(A_ThisFunc . " strUrl", strUrl, "")
  ; 	return strUrl
  ; 	}
}