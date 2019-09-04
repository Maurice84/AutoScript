# Maurice AutoScript

This is a PowerShell framework consisting of operational tasks for Microsoft environments. It contains a menu to easily access and create, modify, move, export Active Directory, Exchange and other objects. To do this it will detect the presence of an Active Directory and/or Exchange in the domain to connect to. It's also possible to connect to Azure and/or Office 365. It's a collection being worked on for several years with passion and dedication during my career as an IT Pro and can be used on all versions of Microsoft Windows containing PowerShell.

## Prerequisites

For the best results you need to have the following updated to run the code

```
Microsoft .NET Framework 4.5 or higher
PowerShell 5.1 or higher
```

## Deployment

There are 2 ways to run this collection:

* Non-merged mode by running Start.ps1: This will load all the separate scripts into memory. This is handy to debug the code if running into errors and therefore no need to scroll through thousands lines of code.
* Merged mode by supplying 'Merge' as an argument when executing Start.ps1: This will create an all-in-one PowerShell file containing all functions for easy transfer and flexibility.

## To do

1. Rework before publishing this to Git:
   * Translate everything to English (as well as comments in the code)
   * Restructuring the menu (it's must follow the folder structure)
   * Rework the customer function (SetCustomer) and all references in the code
   * Remove the email settings from the SetEmail function and add this to a config JSON
2. Function fixes:
   * ActiveDirectory > Create > Account > UsingCSV
     * Not working and fully commented, must be completed
3. Add new features:
   * Exchange > Overview > Mailbox:
     * Add mobile sync details to the export

## Known issues

1. AutoUpdate and FTP has been excluded, this needs to be set to Git

## Built With

* PowerShell IDE (started in 2015)
* [Visual Studio Code](https://code.visualstudio.com/) (enhanced since 2019)


## Authors

* **Maurice Heikens** - *Cloud Infra Specialist* - [Maurice84](https://github.com/Maurice84)
