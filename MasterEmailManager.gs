var ARCHIVE_OLDER_THAN_HOURS = 24; //Specify the number of hours the script should archive emails.  In this example, the script archives emails older than 24 hours hold.
var IMPORTANT_TO_TRASH_AFTER_DAYS = 1461; //Specify the number of days your retention policy or statute of limitations requires.  If your retention policy or statute of limitations is expressed in years, multiply by 365.25.  Here, we will assume the user deletes all emails older than four years old.
var UNIMPORTANT_TO_TRASH_AFTER_DAYS = 3; //Specify the number of days you would like this script to move unimportant emails to trash.  In this example, the script moves unimportant emails older than three days to trash.
//Maximum number of message threads to process per run.
var TRASH_PURGE_AFTER_DAYS = 5; //Specify the number of days you would like the script to delete emails in trash.  In this example, the script deletes emails after five days.
var PAGE_SIZE = 200; //Batch size for GmailApp.search
var ssID = "1adBAqaA5qVB2awvPCf9rc-O6a1S8fzZetUuM6uE5Osc"; //Input spreadsheet ID here.  Use spreadsheet ID from URL.  Here, I input the template spreadsheet.  Do not ask me for editing permissions.  Instead, create a copy of the template sheet this spreadsheet ID points to and replace this spreadsheet ID with the spreadsheet ID in your copy.
//List offensive words in the languages you know in the offensiveWords list.  Be sure to wrap each word with quotation marks and separate each word with a comma.
var offensiveWords = [
  
];
var PROFANITY_TRASH_PURGE = 1; //Specify the number of days you would like this script to delete emails containing profanity.  In this example, the script deletes emails with profanity after one day.
var DELETE_SPAM_AFTER_DAYS = 1; //Specify the number of days you would like this script to delete spam.  In this example, the script deletes spam emails after one day.

//Trigger functions
function createTimeTriggerEveryNHours() {
  ScriptApp.newTrigger("MasterEmailManager")
    .timeBased()
    .atHour(12)
    .everyHours(12) 
    .inTimezone('America/New_York') //Specify time zone.
    .create();
}

function removeSpecificTriggers(handlerFunctionName){
  const triggers = ScriptApp.getProjectTriggers();
  for (let i=0; i < triggers.length; i++){
    const trigger = triggers[i];
    if(trigger.getHandlerFunction() === handlerFunctionName) {
      ScriptApp.deleteTrigger(trigger);
      console.log(`Removed existing trigger for ${handlerFunctionName}.`);
    }
  }
}

//Run below function manually if Gmail quotas are reached.
function stopTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  console.log("All project triggers removed.");
}

//Email maintenance functions
function autoArchiveEmails() {
  console.log(`Archiving emails older than ${ARCHIVE_OLDER_THAN_HOURS} hours.`);
  const searchString = `in:inbox older_than:${ARCHIVE_OLDER_THAN_HOURS}h`
  const threads = GmailApp.search(searchString,0,PAGE_SIZE);
  let archivedCount = 0;
  for (const thread of threads) {
    thread.moveToArchive();
    archivedCount++;
  };
  console.log(`Archived ${archivedCount} threads.`);
  return archivedCount;
}

function handleSpecificDeletions() {
  console.log("Handling specific deletions (named terms and profanity).");
  let movedToTrashCount = 0;
  //Step 1: Process profanity
  if (offensiveWords.length === 0) {
    console.log("No profanity terms specified for deleting emails with profanity.  Skipping step.");
  } else {
    const profanitySearchTerms = offensiveWords.map(word => `"${word}"`).join(' OR ');
    const profanitySearchString = `in:all (${profanitySearchTerms})`;
    const profanityThreads = GmailApp.search(profanitySearchString, 0, PAGE_SIZE);
    for (const thread of profanityThreads) {
      thread.moveToTrash();
      movedToTrashCount++;
    }
    console.log(`Profanity filter moved ${profanityThreads.length} emails to Trash.`);
  }
  //Step 2: Process terms from spreadsheet
  try {
    const ss = SpreadsheetApp.openById(ssID);
    const sheet = ss.getSheets()[0];
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      console.log("Spreadsheet is blank or has only headers.");
    } else {
      const terms = sheet.getRange("A2:A" + lastRow).getValues()
        .flat()
        .filter(String);
      if (terms.length > 0) {
        const combinedSearchTerms = terms.map(term => `"${term}"`).join(' OR ');
        const namedSearchString = `in:all (${combinedSearchTerms})`;
        const namedThreads = GmailApp.search(namedSearchString, 0, PAGE_SIZE);
        for (const thread of namedThreads) {
          thread.moveToTrash();
          movedToTrashCount++;
        }
        console.log(`Spreadsheet terms moved ${namedThreads.length} emails to Trash.`);
      } else {
        console.log("No deletion terms found in spreadsheet.");
      }
    }
  } catch(e) {
    console.error(`Error accessing spreadsheet for deletion terms. ${e.message}`);
  }
  //Step 3: Process password resets
  const resetSearchTerms = "password reset";
  const resetSearchString = `in:all ${resetSearchTerms}`;
  const resetThreads = GmailApp.search(resetSearchString, 0, PAGE_SIZE);
  for (const thread of resetThreads) {
    thread.moveToTrash();
    movedToTrashCount++;
  }
  console.log(`Moved ${resetThreads.length} password reset emails to Trash.`);
  return movedToTrashCount;
}

function purgeOldEmails(){
  console.log(`Moving emails older than ${IMPORTANT_TO_TRASH_AFTER_DAYS/365.25} years to Trash.`);
  const searchString = `in:all older_than:${IMPORTANT_TO_TRASH_AFTER_DAYS}d`;
  const threads = GmailApp.search(searchString, 0, PAGE_SIZE);
  let movedCount = 0;
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - IMPORTANT_TO_TRASH_AFTER_DAYS);
  for (const thread of threads) {
    if (thread.getLastMessageDate() < cutoff) {
      thread.moveToTrash();
      movedCount++;
    }
  }
  console.log(`Moved ${movedCount} emails older than ${IMPORTANT_TO_TRASH_AFTER_DAYS/365.25} years to Trash.`);
  return movedCount;
}

function purgeUnimportantEmails(){
  console.log(`Moving unimportant emails to Trash older than ${UNIMPORTANT_TO_TRASH_AFTER_DAYS} days.`);
  const searchString = `in:all -label:important -label:starred older_than:${UNIMPORTANT_TO_TRASH_AFTER_DAYS}d`;
  const threads = GmailApp.search(searchString, 0, PAGE_SIZE);
  let movedCount = 0;
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - UNIMPORTANT_TO_TRASH_AFTER_DAYS);
  for (const thread of threads) {
    if (thread.getLastMessageDate() < cutoff) {
      thread.moveToTrash();
      movedCount++;
    }
  }
  console.log(`Moved ${movedCount} unimportant emails to Trash.`);
  return movedCount;
}

function purgeProfanityFromTrash() {
  let deletedCount = 0;
  if (offensiveWords.length === 0) {
    console.log("No profanity terms are specified.  Skipping permanent deletion from Trash.");
    return deletedCount;
  }
  const profanitySearchTerms = offensiveWords.map(word => `${word}`).join(' OR ');
  const profanitySearchString = `in:trash (${profanitySearchTerms})`;
  const profanityThreads = GmailApp.search(profanitySearchString, 0, PAGE_SIZE);
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - PROFANITY_TRASH_PURGE);
  for (const thread of profanityThreads) {
    if (thread.getLastMessageDate() < cutoff) {
      Gmail.Users.Threads.remove('me',thread.getId());
      deletedCount++;
    }
  }
  console.log(`Permanently deleted ${deletedCount} emails containing profanity from Trash.`);
  return deletedCount;
}

function purgePasswordResets() {
  let deletedCount = 0;
  const resetSearchTerms = "password reset";
  const resetSearchString = `in:trash ${resetSearchTerms}`;
  const resetThreads = GmailApp.search(resetSearchString, 0, PAGE_SIZE);
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - PROFANITY_TRASH_PURGE);
  for (const thread of resetThreads) {
    if (thread.getLastMessageDate() < cutoff) {
      Gmail.Users.Threads.remove("me", thread.getId());
      deletedCount++;
    }
  }
  console.log(`Permanently deleted ${deletedCount} password reset emails from Trash.`);
  return deletedCount;
}

function purgeTrash() {
  console.log(`Purging trash older than ${TRASH_PURGE_AFTER_DAYS} days.`);
  const searchString = `in:trash older_than:${TRASH_PURGE_AFTER_DAYS}d`;
  const threads = GmailApp.search(searchString, 0, PAGE_SIZE);
  let deletedCount = 0;
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - TRASH_PURGE_AFTER_DAYS);
  for (const thread of threads) {
    if (thread.getLastMessageDate() < cutoff) {
      Gmail.Users.Threads.remove('me',thread.getId());
      deletedCount++;
    }
  }
  console.log(`Permanently deleted ${deletedCount} emails from Trash.`);
  return deletedCount;
}

function purgeSpam() {
 console.log(`Purging spam older than ${DELETE_SPAM_AFTER_DAYS} days.`);
 const searchString = `in:spam older_than:${DELETE_SPAM_AFTER_DAYS}d`;
 const threads = GmailApp.search(searchString, 0, PAGE_SIZE);
 let deletedCount = 0;
 const cutoff = new Date();
 cutoff.setDate(cutoff.getDate() - DELETE_SPAM_AFTER_DAYS);
 for (const thread of threads) {
   if (thread.getLastMessageDate() < cutoff) {
     Gmail.Users.Threads.remove('me',thread.getId());
     deletedCount++;
    }
  }
 console.log(`Permanently deleted ${deletedCount} emails from Spam.`);
 return deletedCount;
}

function MasterEmailManager() {
  let totalArchived = 0;
  let totalMovedToTrash = 0;
  let totalDeletedPermanently = 0;
  try {
    totalArchived += autoArchiveEmails();
  } catch (e) {
    console.error(`Error in Master Email Manager (Archiving): ${e.message}`);
  }
  //If you do not have a spreadsheet for specific deletions, remove this step from your code.
  try {
    totalMovedToTrash += handleSpecificDeletions();
  } catch (e) {
    console.error(`Error in Master Email Manager (Deleting specific emails): ${e.message}`);
  }
  try {
    totalMovedToTrash += purgeOldEmails();
  } catch (e) {
    console.error(`Error in Master Email Manager (Deleting old emails): ${e.message}`);
  }
  try {
    totalMovedToTrash += purgeUnimportantEmails();
  } catch (e) {
    console.error(`Error in Master Email Manager (Deleting unimportant emails): ${e.message}`);
  }
  try {
    totalDeletedPermanently += purgeProfanityFromTrash();
  } catch (e) {
    console.error(`Error in Master Email Manager (Permanently deleting trash with profanity): ${e.message}`);
  }
  try {
    totalDeletedPermanently += purgePasswordResets();
  } catch (e) {
    console.error(`Error in Master Email Manager (Deleting password reset emails): ${e.message}`);
  }
  try {
    totalDeletedPermanently += purgeTrash();
  } catch (e) {
    console.error(`Error in Master Email Manager (Permanently deleting trash): ${e.message}`);
  }
  try {
    totalDeletedPermanently += purgeSpam();
  } catch (e) {
    console.error(`Error in Master Email Manager (Permanently deleting spam): ${e.message}`);
  }
  console.log("Master email manager run finished.");
  console.log(`Summary: Archived ${totalArchived}, Moved to Trash ${totalMovedToTrash}, Permanently Deleted ${totalDeletedPermanently}.`);
}
