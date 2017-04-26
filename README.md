# outlook_attachments
Automatically save outlook attachments using this script.

The script itself only handles saving the files. To tell it what to save, create a new rule in your inbox. 

For example, you can make a rule that messages from a specific person with certain words in the subject 
    line will be moved to a folder and all attachments will be saved.

Your rule description may look something like this:

Apply this rule after the message arrives

from "SENDER"
and with "SUBJECT LINE" in the subject
move it to the "FOLDER" folder
and run Project1.saveAttachtoDisk
