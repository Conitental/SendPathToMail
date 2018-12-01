# SendPathToMail
This is a VB Script to be included in windows "Send To" menu to open a mail with the filepaths of the selected files.

It resolves the UNC paths of network drives and includes the actual paths.
So `H:\Somefile.txt` will result in `\\Server01\Somefile.txt`.

## Requirements
SendPathToMail will only work on machines with installed Microsoft Outlook application.

## Installation
To use this extension simply place the script itself or a link to it in the following windows users profile path:
`%APPDATA%\Roaming\Microsoft\Windows\SendTo`

## Usage
To use the script after installation simply select at least one file and select `Send to` in the context menu and then select `SendPathToMail`. This will open outlook with an empty mail and the actual filepaths.