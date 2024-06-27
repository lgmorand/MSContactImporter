# MSContactImporter

![CI Build](https://github.com/lgmorand/MSContactImporter/workflows/CI%20Build/badge.svg)

## Presentation

Is an unofficial tool designed to import a large set of users from an ActiveDirectory into the Contacts list of Outlook based on the hierarchical structure of the directory.
The tool has been tested for internal usage only 

## Download

Download the latest version from the releases page: [https://github.com/lgmorand/MSContactImporter/releases](https://github.com/lgmorand/MSContactImporter/releases)

## Usage

Run it, specify your email and your domain's password and then follow the wizard. You'll have to add the alias(es) of the manager(s) from whom you want to retrieve the subordinates. Since version 1.1 you can also add the alias(es) of distribution list(s) or security group(s) to import all its (their) members (with their subordinates as well if you want to). 
The program can:

- add contacts : add contacts who are subordinates of the defined alias(es) / add members of the defined distribution list or security group
- update contacts : it will update the information for the contacts previously imported
- delete orphans : delete previously imported contacts which are not in the directory anymore

> /!\ All managed contacts are marked in the category "Ms Staff v2" and others contacts will remain untouched

## Compatibility

The program has been tested against Office 2016 *(32 bits and 64 bits)* and requires the minimum version of .Net framework 4.8

## Known issues and solution

### Testing connection is failing

Solution: open the config file and ensure that RootMSFTees contains a valid alias name

### The property XXXX/MsStaffId is unknown or cannot be found

Solution: It means that you have corrupted contacts inside the category "Ms Staff v2". Remove old contacts and re-import them

### Un ou plusieurs éléments dans le dossiers que vous avez synchronisé ne correspondent pas

Solution: There are conflicted elements in your Outlook data file. Please clean it before using the tool: [https://www.fieldstonsoftware.com/support/support_gsyncit_olsyncfail.shtml](https://www.fieldstonsoftware.com/support/support_gsyncit_olsyncfail.shtml)

### My contacts have a prefix "MS:" on my iphone

Indeed, there is a field (file as) which contains a prefix and was never used. with iOS 15, Apple decided to change the behavior of the Contacts app and now uses this field. There is not known workaround in settings to change the display format.
