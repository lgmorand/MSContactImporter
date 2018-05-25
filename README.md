# MSContactImporter

## Presentation

Is an unofficial tool designed to import a large set of users from an ActiveDirectory into the Contacts list of Outlook based on the hierarchical structure of the directory.
The tool has been tested for internal usage only

## Usage

Run it, specify your email and your domain's password and then follow the wizard. You'll have to add the alias(es) of the manager(s) from whom you want to retrieve the subordinates.
The program can:

- add contacts : add contacts who are subordinates of the defined alias(es)
- update contacts : it will update the information for the contacts previously imported
- delete duplicate : delete previously imported contacts which are not in the directory anymore

> /!\ All managed contacts are marked in the category "Ms Staff v2" and others contacts will remain untouched

## Compatibility

The program has been tested against Office 2016 *(32 bits)* and requires the minimum version of .Net framework 4.5.2

## Known issues and solution

### The property XXXX/MsStaffId is unknown or cannot be found

Solution: It means that you have corrupted contacts inside the category "Ms Staff v2". Remove old contacts and re-import them

### Un ou plusieurs éléments dans le dossiers que vous avez synchronisez ne correspondent pas

Solution: There are conflicted elements in your Outlook data file. Please clean it before using the tool: https://www.fieldstonsoftware.com/support/support_gsyncit_olsyncfail.shtml