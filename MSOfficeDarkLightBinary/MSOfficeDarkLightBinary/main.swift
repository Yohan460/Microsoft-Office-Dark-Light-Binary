//
//  main.swift
//  MSOfficeDarkLightBinary
//
//  Created by Johan McGwire on 12/12/18.
//  Copyright © 2018 Johan McGwire. All rights reserved.
//

import Foundation

func runCommand(cmd : String, args : String...) -> (output: [String], error: [String], exitCode: Int32) {
    
    var output : [String] = []
    var error : [String] = []
    
    let task = Process()
    task.launchPath = cmd
    task.arguments = args
    
    let outpipe = Pipe()
    task.standardOutput = outpipe
    let errpipe = Pipe()
    task.standardError = errpipe
    
    task.launch()
    
    let outdata = outpipe.fileHandleForReading.readDataToEndOfFile()
    if var string = String(UnsafePointer(outdata.bytes)) {
        string = string.stringByTrimmingCharactersInSet(NSCharacterSet.newlineCharacterSet())
        output = string.componentsSeparatedByString("\n")
    }
    
    let errdata = errpipe.fileHandleForReading.readDataToEndOfFile()
    if var string = String(UnsafePointer(errdata.bytes)) {
        string = string.stringByTrimmingCharactersInSet(NSCharacterSet.newlineCharacterSet())
        error = string.componentsSeparatedByString("\n")
    }
    
    task.waitUntilExit()
    let status = task.terminationStatus
    
    return (output, error, status)
}
func ForceLightMode(appName: String, rootAuthority: Bool){
    if rootAuthority{
        //shell("defaults", "write", "/Library/Preferences/com.microsoft.\(appName).plist", "NSRequiresAquaSystemAppearance", "-bool", "yes")
    } else {
        runCommand(cmd: "/usr/bin/defaults", args: "write", "~/Library/Preferences/com.microsoft.\(appName).plist", "NSRequiresAquaSystemAppearance", "-bool", "yes")
    }
}

func ResetApp(appName: String, rootAuthority: Bool){
    if rootAuthority{
        //shell("defaults delete /Library/Preferences/com.microsoft.\(appName).plist NSRequiresAquaSystemAppearance")
    } else {
        //shell("defaults delete ~/Library/Preferences/com.microsoft.\(appName).plist NSRequiresAquaSystemAppearance")
    }
}

let version = "0.0.1"

// full arguments list as single string
let argString = CommandLine.arguments.joined(separator: " ").uppercased()

// print help and quit if asked
if argString.contains("-H") || argString.contains("-HELP") {
    print("""
MSOfficeDarkLightBinary is a utility to help you manage the preferences of Microsoft office products to enable, disable, or toggle the dark and light modes.

version: \(version)

Note: This utilty manages the preferences at the highest level of authority it is run (eg, user level = user preferences, root level = system preferences)

Some basic options:

-Version                 : prints the version number
-Help | -h               : prints this help statement
-LightMode <app names>   : sets light mode as the default mode for those applications
-ResetApps <app names>   : sets the applications back to their default states
        
Usage:

    User priv:
        
    MSOfficeDarkLightBinary -LightMode Outlook Powerpoint
        
        Will set outlook and powerpoint to use light mode for that specific user.
        
    MSOfficeDarkLightBinary -ResetApps Outlook Powerpoint
        
        Will set outlook and powerpoint back to their default preferences
        
    Root priv:
        
    MSOfficeDarkLightBinary -LightMode Outlook Powerpoint
        
        Will set outlook and powerpoint to use light mode in the /Library/Preferences folder which applies to all users
        
    MSOfficeDarkLightBinary -ResetApps Outlook Powerpoint
        
        Will set outlook and powerpoint back to their default preferences in the /Library/Preferences folder which applies to all users
""")
    exit(0)
}

// print version and quit if asked
if argString.contains("-VERSION") {
    print(version)
    exit(0)
}

// Determining if the binary is being run as root or not
let rootAuthority: Bool = (NSUserName() == "root")

// Determining the apps to change to light mode

var lightModeApps:[String] = [], resetApps:[String] = []

if argString.contains("-LIGHTMODE") || argString.contains("-RESETAPPS") {
    let argArrayCap = (CommandLine.arguments).map{$0.uppercased()}
    var i = 1
    while i < argArrayCap.count {
        if argArrayCap[i] == "-LIGHTMODE" {
            i += 1
            if i >= argArrayCap.count{break}
            while !(argArrayCap[i]).hasPrefix("-"){
                lightModeApps.append((CommandLine.arguments)[i])
                i += 1
                if i >= argArrayCap.count{break}
            }
            i -= 1
        }
        if argArrayCap[i] == "-RESETAPPS" {
            i += 1
            if i >= argArrayCap.count{break}
            while !(argArrayCap[i]).hasPrefix("-"){
                resetApps.append((CommandLine.arguments)[i])
                i += 1
                if i >= argArrayCap.count{break}
            }
            i -= 1
        }
        i += 1
    }
}

if lightModeApps.count > 0 {
    for lightApp in lightModeApps {
        ForceLightMode(appName: lightApp, rootAuthority: rootAuthority)
    }
}

if resetApps.count > 0 {
    for resetApp in resetApps {
        ResetApp(appName: resetApp, rootAuthority: rootAuthority)
    }
}
//outlook
//excel
//powerpoint
//onenote
//word
//visual-studio
//vscode
//autoupdate

