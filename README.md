# SendPathToMail
This is a VB Script to be included in windows "Send To" menu to open a mail with the filepaths of the selected files.

It resolves the UNC paths of network drives and includes the actual paths.
So `H:\Somefile.txt` will result in `\\Server01\Somefile.txt`.

## Requirements
SendPathToMail will only work on machines with installed Microsoft Outlook application.

## Installation
The installation can be done using the releases binary. Currently the installation requires administrator rights.

## Usage
Use SPTM by selecting up to 15 (this is the windows threshold up to which context menu items are displayed) folders or files on a network drive or share and select `Send path as mail` in the context menu.
