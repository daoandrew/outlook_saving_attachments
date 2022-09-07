from pathlib import Path
import win32com.client

# Create output folder
output_dir = Path.cwd() / "Output"
output_dir.mkdir(parents=True, exist_ok=True)

# Connect to Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNameSpace("MAPI")

# Connect to Inbox
inbox = outlook.GetDefaultFolder(6)

messages = inbox.Folders.Item("Safety").Items

for message in messages:
    attachments = message.Attachments
    subject = message.Subject
    body = message.body

    try: 
        target_folder = output_dir / str(subject)
    except NotADirectoryError:
        pass
    else:
        target_folder = output_dir

    target_folder.mkdir(parents=True, exist_ok=True)

    # Write body to text file
    Path(target_folder / "EMAIL_BODY.txt").write_text(str(body))

    # Save attachements
    for attachment in attachments:
        attachment.SaveAsFile(target_folder / str(attachment))

