import win32com.client


def extract(count):
    """Get emails from outlook."""
    items = []
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # "6" refers to the inbox
    messages = inbox.Items
    message = messages.GetFirst()
    i = 0
    while message:
        try:
            message = dict()
            message["Subject"] = getattr(message, "Subject", "<UNKNOWN>")
            message["SentOn"] = getattr(message, "SentOn", "<UNKNOWN>")
            message["EntryID"] = getattr(message, "EntryID", "<UNKNOWN>")
            message["Sender"] = getattr(message, "Sender", "<UNKNOWN>")
            message["Size"] = getattr(message, "Size", "<UNKNOWN>")
            message["Body"] = getattr(message, "Body", "<UNKNOWN>")
            items.append(message)
        except Exception as ex:
            print("Error processing mail", ex)
        i += 1
        if i < count:
            message = messages.GetNext()
        else:
            return items

    return items


def show_message(items):
    """Show the messages."""
    items.sort(key=lambda tup: tup["SentOn"])
    for i in items:
        print(i["SentOn"], i["Subject"])


def main():
    """Fetch and display top message."""
    items = extract(5)
    show_message(items)


if __name__ == "__main__":
    main()