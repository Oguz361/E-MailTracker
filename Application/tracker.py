import win32com.client as client

def main():
    emailCount = {}
    try:
        outlook = client.Dispatch('Outlook.Application')
        namespace = outlook.GetNameSpace('MAPI')
        
        FOLDER_IDS = {
            "Posteingang": 6,
            "Spam": 23  
        }
        
        # Alle definierten Ordner durchlaufen
        for folder_name, folder_id in FOLDER_IDS.items():
            try:
                folder = namespace.GetDefaultFolder(folder_id)
                print(f"\nVerarbeite Ordner: {folder_name}")
                
                # Emails im Ordner zählen
                for email in folder.Items:
                    try:
                        absender = email.SenderName
                        absender_mit_ordner = f"{absender} ({folder_name})"
                        
                        if absender_mit_ordner in emailCount:
                            emailCount[absender_mit_ordner] += 1
                        else:
                            emailCount[absender_mit_ordner] = 1
                    except Exception as e:
                        print(f"Fehler bei Email in {folder_name}: {e}")
                        continue
                        
            except Exception as e:
                print(f"Fehler beim Zugriff auf Ordner {folder_name}: {e}")
                continue
        
        print("\n=== Ergebnisse ===")
        
        for folder_name in FOLDER_IDS.keys():
            print(f"\n--- {folder_name} ---")
            folder_emails = {k: v for k, v in emailCount.items() if folder_name in k}
            for absender, anzahl in folder_emails.items():
                print(f"{absender}: {anzahl}")
        
        print("\n=== Top 10 Absender (alle Ordner) ===")
        sorted_absender = sorted(emailCount.items(), key=lambda x: x[1], reverse=True)
        for absender, anzahl in sorted_absender[:10]:
            print(f"{absender}: {anzahl}")
            
    except Exception as e:
        print(f"Fehler beim Zugriff auf Outlook: {e}")

if __name__ == '__main__':
    main()