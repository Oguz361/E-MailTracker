import win32com.client as client

def main():
    emailCount = {}
    try:
        outlook = client.Dispatch('Outlook.Application')
        namespace = outlook.GetNameSpace('MAPI')
        inbox = namespace.GetDefaultFolder(6)
        
        for email in inbox.Items:

            absender = email.SenderName

            if absender in emailCount:
                emailCount[absender] += 1
            else:
                emailCount[absender] = 1

        for absender, anzahl in emailCount.items():
            print(f"{absender}: {anzahl}")


        
    except Exception as e:
        print(f"Fehler beim Zugriff auf Outlook: {e}")

if __name__ == '__main__':
    main()