import win32com.client
import pythoncom
import os

import re

def clean_email_body(body):
    # Remove URLs inside < >
    body = re.sub(r'<https?://[^>]+>', '', body)

    # Remove normal URLs
    body = re.sub(r'https?://\S+', '', body)

    # Remove extra spaces before/after lines
    lines = [line.strip() for line in body.splitlines()]

    # Remove empty lines at beginning/end
    lines = [line for line in lines if line]

    return "\n".join(lines)


class OutlookManager:

    def __init__(self):
        self.outlook = None
        self.namespace = None

    def connect(self):

        pythoncom.CoInitialize()

        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")

            print("Connected to Outlook")
            print("Namespace:", self.namespace)

            return True

        except Exception as e:
            print("Outlook connection error:", e)
            return False

    def get_inbox(self, limit=20):

        pythoncom.CoInitialize()

        try:
        
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")

            # Get mailbox
            mailbox = namespace.Folders.Item(1)

            print("Mailbox:", mailbox.Name)

            # Get Inbox
            
            inbox = mailbox.Folders.Item("Inbox")


            messages = inbox.Items

            try:
                messages.Sort("[ReceivedTime]", True)
            except Exception:
                pass

            emails = []

            for message in messages:

                if len(emails) >= limit:
                    break

                try:

                    # emails.append({
                    #     "subject": str(message.Subject or ""),
                    #     "sender": str(message.SenderName or ""),
                    #     "sender_email": str(message.SenderEmailAddress or ""),
                    #     "received": str(message.ReceivedTime),
                    #     "body": str(message.Body or ""),
                    #     "unread": bool(message.UnRead),
                    #     "importance": int(message.Importance),
                    #     "arrival_time": str(message.ReceivedTime),
                    #      
                    # })
                    received_time = message.ReceivedTime.strftime("%d %b - %H:%M")

                    body = clean_email_body(str(message.Body))

                    attachments = []

                    for attachment_index, attachment in enumerate(message.Attachments):

                        attachments.append({
                            "index": attachment_index,
                            "name": str(attachment.FileName),
                            "size": int(attachment.Size)
                        })


                    emails.append({
                        "sender": str(message.SenderName or ""),
                        "sender_email": str(message.SenderEmailAddress or ""),
                        "timestamp": received_time,
                        "body": str(body or ""),
                        "subject": str(message.Subject or ""),
                        "html_body": str(message.HTMLBody or " "),
                        "attachments": attachments
                    })


                except Exception as e:
                    print("Email error:", e)

            return {
                "success": True,
                "emails": emails
            }

        except Exception as e:

            print("Inbox error:", e)

            return {
                "success": False,
                "message": str(e),
                "emails": []
            }

        finally:
            pythoncom.CoUninitialize()




    def get_important(self, limit=20):

        pythoncom.CoInitialize()

        try:

            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")

            mailbox = namespace.Folders.Item(1)

            gmail_folder = mailbox.Folders.Item("[Gmail]")

            important = gmail_folder.Folders.Item("Important")

            messages = important.Items

            try:
                messages.Sort("[ReceivedTime]", True)
            except Exception:
                pass

            emails = []

            for message in messages:

                if len(emails) >= limit:
                    break

                try:

                    received_time = message.ReceivedTime.strftime("%d %b - %H:%M")
                    
                    body = clean_email_body(str(message.Body))
                    
                    attachments = []
                    
                    for attachment_index, attachment in enumerate(message.Attachments):
                    
                        attachments.append({
                            "index": attachment_index,
                            "name": str(attachment.FileName),
                            "size": int(attachment.Size)
                        })
                    
                    
                    emails.append({
                    "sender": str(message.SenderName or ""),
                    "sender_email": str(message.SenderEmailAddress or ""),
                    "timestamp": received_time,
                    "body": str(body or ""),
                    "subject": str(message.Subject or ""),
                    "html_body": str(message.HTMLBody or " "),
                    "attachments": attachments
                    })

                except Exception as e:
                    print("Important error:", e)

            return {
                "success": True,
                "emails": emails
            }

        except Exception as e:

            print("Important folder error:", e)

            return {
                "success": False,
                "message": str(e),
                "emails": []
            }

        finally:
            pythoncom.CoUninitialize()


    def get_sent_emails(self, limit=20):

        pythoncom.CoInitialize()

        try:

            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")

            mailbox = namespace.Folders.Item(1)

            gmail_folder = mailbox.Folders.Item("[Gmail]")

            sent_folder = gmail_folder.Folders.Item("Sent Mail")

            messages = sent_folder.Items

            try:
                messages.Sort("[SentOn]", True)
            except Exception:
                pass

            emails = []

            

            for message in messages:

                if len(emails) >= limit:
                    break

                try:
                    attachments = []
                                
                    for attachment_index, attachment in enumerate(message.Attachments):
                                
                        attachments.append({
                            "index": attachment_index,
                            "name": str(attachment.FileName),
                            "size": int(attachment.Size)
                        })

                    emails.append({
                        "subject": str(message.Subject or ""),
                        "to": str(message.To or ""),
                        "sender": str(message.SenderName or ""),
                        "sent": str(message.SentOn.strftime("%d %b - %H:%M")),
                        "body": str(message.Body or ""),
                        "attachments": attachments
                    })

                except Exception as e:
                    print("Sent email error:", e)

            return {
                "success": True,
                "emails": emails
            }

        except Exception as e:

            print("Sent folder error:", e)

            return {
                "success": False,
                "message": str(e),
                "emails": []
            }

        finally:
            pythoncom.CoUninitialize()


    def send_email(self, to, subject, body, attachments=None):

        pythoncom.CoInitialize() 
        try: 
            outlook = win32com.client.Dispatch("Outlook.Application") 
            mail = outlook.CreateItem(0) 
            mail.To = to 
            mail.Subject = subject 
            mail.Body = body 
            # Add attachments 
            if attachments: 
                for attachment in attachments: 
                    if attachment: 
                        mail.Attachments.Add(attachment) 

            mail.Send() 
            return { "success": True, "message": "Email sent successfully" } 
        except Exception as e: 
            return { "success": False, "message": str(e) } 
        
        finally: 
            pythoncom.CoUninitialize()

    def download_attachment(self, email_index, attachment_index):

        pythoncom.CoInitialize()

        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")

            mailbox = namespace.Folders.Item(1)
            inbox = mailbox.Folders.Item("Inbox")

            messages = inbox.Items
            messages.Sort("[ReceivedTime]", True)

            # Get the email
            message = messages.Item(email_index + 1)

            # Get attachment
            attachment = message.Attachments.Item(attachment_index + 1)

            # Download folder
            download_folder = os.path.join(
                os.path.expanduser("~"),
                "Downloads"
            )

            os.makedirs(download_folder, exist_ok=True)

            file_name = str(attachment.FileName)
            file_path = os.path.join(download_folder, file_name)

            attachment.SaveAsFile(file_path)

            return {
                "success": True,
                "message": "Attachment downloaded",
                "path": file_path,
                "name": file_name
            }

        except Exception as e:

            return {
                "success": False,
                "message": str(e)
            }

        finally:
            pythoncom.CoUninitialize()

    def refresh_outlook(self):

        pythoncom.CoInitialize()

        try:
            outlook = win32com.client.Dispatch("Outlook.Application")

            # Get the active Outlook Explorer
            explorer = outlook.ActiveExplorer()

            if explorer is not None:
                explorer.CommandBars.ExecuteMso("SendReceiveAll")

            return {
                "success": True,
                "message": "Send/Receive started"
            }

        except Exception as e:
            return {
                "success": False,
                "message": str(e)
            }

        finally:
            pythoncom.CoUninitialize()