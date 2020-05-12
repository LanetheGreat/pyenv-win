Option Explicit

Dim objfs
Dim objws
Dim objweb

Set objfs = CreateObject("Scripting.FileSystemObject")
Set objws = WScript.CreateObject("WScript.Shell")
Set objweb = CreateObject("WinHttp.WinHttpRequest.5.1")

' Set proxy settings, called on library import for objweb.
Sub SetProxy()
    Dim httpProxy
    Dim proxyArr

    httpProxy = objws.Environment("Process")("http_proxy")
    If httpProxy <> "" Then
        If InStr(1, httpProxy, "@") > 0 Then
            ' The http_proxy environment variable is set with basic authentication
            ' WinHttp seems to work fine without the credentials, so we should be
            ' okay with just the hostname/port part
            proxyArr = Split(httpProxy, "@")
            objweb.setProxy 2, proxyArr(1)
        Else
            objweb.setProxy 2, httpProxy
        End If
    End If
End Sub
SetProxy

' Regular expressions for extracting python version numbers.
Dim regexVer
Dim regexFile
Set regexVer = New RegExp
Set regexFile = New RegExp
With regexVer
    .Pattern = "^(\d+)(?:\.(\d+))?(?:\.(\d+))?(?:([a-z]+)(\d*))?$"
    .Global = True
    .IgnoreCase = True
End With
With regexFile
    .Pattern = "^python-(\d+)(?:\.(\d+))?(?:\.(\d+))?(?:([a-z]+)(\d*))?([\.-]amd64)?(-webinstall)?\.(exe|msi)$"
    .Global = True
    .IgnoreCase = True
End With

Dim strCurrent
Dim strPyenvHome
Dim strPyenvParent
Dim strDirCache
Dim strDirVers
Dim strDirLibs
Dim strDirShims
Dim strDBFile
Dim strVerFile
strCurrent   = objfs.GetAbsolutePathName(".")
strPyenvHome = objfs.getParentFolderName(objfs.getParentFolderName(WScript.ScriptFullName))
strPyenvParent = objfs.getParentFolderName(strPyenvHome)
strDirCache  = strPyenvHome & "\install_cache"
strDirVers   = strPyenvHome & "\versions"
strDirLibs   = strPyenvHome & "\libexec"
strDirShims  = strPyenvHome & "\shims"
strDBFile    = strPyenvHome & "\.versions_cache.xml"
strVerFile   = "\.python-version"

Function GetCurrentVersionsGlobal()
    GetCurrentVersionsGlobal = Null

    Dim fname
    Dim objFile
    fname = strPyenvHome & "\version"
    If objfs.FileExists(fname) Then
        Set objFile = objfs.OpenTextFile(fname)
        If objFile.AtEndOfStream <> True Then
           GetCurrentVersionsGlobal = Array(Split(objFile.ReadLine, ";"), fname)
        End If
        objFile.Close
    End If
End Function

Function GetCurrentVersionsLocal(path)
    GetCurrentVersionsLocal = Null

    Dim fname
    Dim objFile
    Do While path <> ""
        fname = path & strVerFile
        If objfs.FileExists(fname) Then
            Set objFile = objfs.OpenTextFile(fname)
            If objFile.AtEndOfStream <> True Then
               GetCurrentVersionsLocal = Array(Split(objFile.ReadLine, ";"), fname)
            End If
            objFile.Close
            Exit Function
        End If
        path = objfs.GetParentFolderName(path)
    Loop
End Function

Function GetCurrentVersionsShell()
    GetCurrentVersionsShell = Null

    Dim str
    str = objws.Environment("Process")("PYENV_VERSION")
    If str <> "" Then _
        GetCurrentVersionsShell = Array(Split(str, ";"), "%PYENV_VERSION%")
End Function

Function GetCurrentVersions()
    Dim str
    str = GetCurrentVersionsShell
    If IsNull(str) Then str = GetCurrentVersionsLocal(strCurrent)
    If IsNull(str) Then str = GetCurrentVersionsGlobal
    If IsNull(str) Then
		WScript.echo "No global python version has been set yet. Please set the global version by typing:"
		WScript.echo "pyenv global 3.8.2"
		WScript.quit
	End If
	GetCurrentVersions = str
End Function

Function GetCurrentVersionsNoError()
    Dim str
    str = GetCurrentVersionsShell
    If IsNull(str) Then str = GetCurrentVersionsLocal(strCurrent)
    If IsNull(str) Then str = GetCurrentVersionsGlobal
    GetCurrentVersionsNoError = str
End Function

Function IsVersion(version)
    Dim re
    Set re = new regexp
    re.Pattern = "^[a-zA-Z_0-9-.]+$"
    IsVersion = re.Test(version)
End Function

Function GetBinDir(ver)
    Dim str
    str = strDirVers &"\"& ver
    If Not(IsVersion(ver) And objfs.FolderExists(str)) Then
		WScript.Echo "pyenv specific python requisite didn't meet. Project is using different version of python."
		WScript.Echo "Install python '"& ver &"' by typing: 'pyenv install "& ver &"'"
		WScript.Quit
	End If
    GetBinDir = str
End Function

Sub SetGlobalVersions(versions)
    Dim ver
    For Each ver In versions
        GetBinDir(ver)
    Next

    With objfs.CreateTextFile(strPyenvHome &"\version", True)
        .WriteLine(Join(versions, ";"))
        .Close
    End With
End Sub

Sub SetLocalVersions(versions, localPath)
    Dim ofile
    Dim ver

    For Each ver In versions
        GetBinDir(ver)
    Next

    If objfs.FileExists(localPath) Then
        Set ofile = objfs.OpenTextFile(localPath, 2)
    Else
        Set ofile = objfs.CreateTextFile(localPath, True)
    End If

    With ofile
        .WriteLine(Join(versions, ";"))
        .Close
    End With
End Sub

Function GetExtensions(addPy)
    Dim exts
    exts = ";"& objws.Environment("Process")("PATHEXT") &";"
    Set GetExtensions = CreateObject("Scripting.Dictionary")

    If addPy Then
        If InStr(1, exts, ";.PY;", 1) = 0 Then exts = exts &".PY;"
        If InStr(1, exts, ";.PYW;", 1) = 0 Then exts = exts &".PYW;"
    End If
    exts = Mid(exts, 2, Len(exts)-2)

    Do While InStr(1, exts, ";;", 1) <> 0
        exts = Replace(exts, ";;", ";")
    Loop

    Dim ext
    For Each ext In Split(exts, ";")
        GetExtensions.Item(ext) = Empty
    Next
End Function

Function GetExtensionsNoPeriod(addPy)
    Dim key
    Set GetExtensionsNoPeriod = GetExtensions(addPy)
    For Each key In GetExtensionsNoPeriod.Keys
        If Left(key, 1) = "." Then
            GetExtensionsNoPeriod.Key(key) = LCase(Mid(key, 2))
        Else
            GetExtensionsNoPeriod.Key(key) = LCase(key)
        End If
    Next
End Function

Sub WriteWinScript(strDirBin, shimName, execName)
    With objfs.CreateTextFile(strDirShims &"\"& shimName &".bat")
        .WriteLine("@echo off")
        .WriteLine("setlocal")
        .WriteLine("chcp 1250 > NUL")
        .WriteLine("set ""PATH="& strDirBin &"\Scripts;"& strDirBin &";%PATH%""")
        .WriteLine(execName &" %*")
        .Close
    End With
End Sub

Sub WriteLinuxScript(strDirBin, shimName, execName)
    With objfs.CreateTextFile(strDirShims &"\"& shimName)
        .WriteLine("#!/bin/sh")
        .WriteLine("export PATH="& strDirBin &"/Scripts:"& strDirBin &":$PATH")
        .WriteLine(execName &" $*")
        .Close
    End With
End Sub

Function ReverseArray(arr)
    Dim reverse(), i
    ReDim reverse(UBound(arr))

    For i = LBound(arr) To UBound(arr)
        reverse(i) = arr(UBound(arr) - i)
    Next

    ReverseArray = reverse
End Function

Sub Rehash()
    Dim file

    If Not objfs.FolderExists(strDirShims) Then objfs.CreateFolder(strDirShims)
    For Each file In objfs.GetFolder(strDirShims).Files
        file.Delete True
    Next

    Dim versions, ver, verMatch
    Dim verMajor, verMinor
    Dim winBinDir, nixBinDir
    Dim exts
    Dim baseName
    Dim appendVer

    versions = GetCurrentVersionsNoError()
    If IsNull(versions) Then Exit Sub

    Set exts = GetExtensionsNoPeriod(True)

    For Each ver In ReverseArray(versions(0))
        Set verMatch = regexVer.Execute(ver)
        appendVer = (verMatch.Count > 0)
        If appendVer Then
            verMajor = verMatch(0).SubMatches(0)
            verMinor = verMatch(0).SubMatches(1)
        End If

        winBinDir = strDirVers &"\"& ver
        If Not objfs.FolderExists(winBinDir) Then Exit Sub

        nixBinDir = "/"& Replace(Replace(winBinDir, ":", ""), "\", "/")

        For Each file In objfs.GetFolder(winBinDir).Files
            If exts.Exists(LCase(objfs.GetExtensionName(file))) Then
                baseName = objfs.GetBaseName(file)
                WriteWinScript winBinDir, baseName, baseName
                WriteLinuxScript nixBinDir, baseName, baseName

                If appendVer And (LCase(basename) = "python" Or LCase(basename) = "pythonw") Then
                    WriteWinScript winBinDir, baseName & verMajor &"."& verMinor, baseName
                    WriteLinuxScript nixBinDir, baseName & verMajor &"."& verMinor, baseName

                    WriteWinScript winBinDir, baseName & verMajor, baseName
                    WriteLinuxScript nixBinDir, baseName & verMajor, baseName
                End If    
            End If
        Next

        If objfs.FolderExists(winBinDir & "\Scripts") Then
            For Each file In objfs.GetFolder(winBinDir & "\Scripts").Files
                If exts.Exists(LCase(objfs.GetExtensionName(file))) Then
                    baseName = objfs.GetBaseName(file)
                    WriteWinScript winBinDir, baseName, baseName
                    WriteLinuxScript nixBinDir, baseName, baseName
                End If
            Next
        End If
    Next
End Sub

' SYSTEM:PROCESSOR_ARCHITECTURE = AMD64 on 64-bit computers. (even when using 32-bit cmd.exe)
Function Is32Bit()
    Dim arch
    arch = objws.Environment("Process")("PYENV_FORCE_ARCH")
    If arch = "" Then arch = objws.Environment("System")("PROCESSOR_ARCHITECTURE")
    Is32Bit = (UCase(arch) = "X86")
End Function

' If on a 32bit computer, default to -win32 versions.
Function Check32Bit(version)
    If Is32Bit And Right(LCase(version), 6) <> "-win32" Then _
        version = version & "-win32"
    Check32Bit = version
End Function
