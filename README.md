# Contacts Updater Google Apps Script Library

This Library is used to allow syncing of contacts when you don't have access/permission to third party software, just google drive and contacts.

It works by creating a trigger to export contacts from a specified email into a google spreadsheet. Every morning this will run again to get more up to date contacts

All the users that want to sync their emails with these contacts will then run the import function. 

This will create a trigger that will run every morning to put those contacts on their email account.
It checks for outdated contacts, no longer exisiting ones, missing ones, etc and will update the contacts to reflect the changes.

All of this is done on specified groups, so you can still have external contacts tied to those emails that won't be affected by this project.

The scriptID to use the library is

`1DYBGuLF7xlNTZFTykP-nTznIUgKNM5hh0jFY-uRdkvo3DIZvjOVw0Adc`

An example of client-side set up is 
```
function ExportContacts() {
  // Runs our exporter script.
  // First creates the exporter
  // This first array is group names on your office email to export
  // The second is a mission/company group name, that all contacts made will go under
  // The third is group names we are pulling from, but don't want to create on the other accounts
  ContactUpdater.Export({ source_groups: ['starred', 'ICE', 'Static', 'Church', 'IMOS Roster'], mission_group: 'NDBM Contacts',
                          exclude_groups: ['IMOS Roster', 'Static'], remove_duplicate_numbers: true })
}

function ImportContacts() {
  // Note, the trigger has to be a name of a function that you create here.
  // For example Importer is a function down below
  ContactUpdater.Import({black_list_emails: ['northdakota.bismarck@missionary.org']}, {trigger_handler: 'Importer', trigger_frequency: 1})
}

function Importer(){
  // If you leave out the trigger name it won't delete the old trigger and create a new one. Which is what we want
  ContactUpdater.Import({black_list_emails: ['northdakota.bismarck@missionary.org']})
}

// Function used to clear someones daily trigger to update their contacts
function ClearUsersTrigger(){
  ContactUpdater.clearTriggers()
}
```

All you would need to do is run the export function once on the office email, and then create a macro to run the importContacts. Whenever you want to add another email to sync to this all you have to do is run the macro. 
The first time creating all the contacts will take 5 min, due to api restrictions on photo uploads. After that it should only take 10 seconds to update them.
