import webview
from outlooker_manager import OutlookManager
from llm_embedding import mail_reply_generator


class API:

    def __init__(self):
        self.outlook = OutlookManager()

    def get_inbox(self):
        return self.outlook.get_inbox()

    def get_important(self):
        return self.outlook.get_important()

    def get_sent(self):
        return self.outlook.get_sent_emails()

    def send_email(self, to, subject, body,attachment):
        return self.outlook.send_email(to, subject, body,attachment)

    def download_attachment(self,email_idx, attachment_idx):
        return self.outlook.download_attachment(email_idx,attachment_idx)

    def reply_generate(self,user_role, message_subject, message_content):
        return mail_reply_generator(user_role, message_subject, message_content)

    def sync_outlook(self):
        return self.outlook.refresh_outlook()

    def select_files(self): 
        files = webview.windows[0].create_file_dialog( webview.OPEN_DIALOG, allow_multiple=True ) 
        if not files: 
            return [] 
        return list(files)


api = API()

window = webview.create_window(
    "Email Automator",
    "ui/index.html",
    js_api=api,
    width=1200,
    height=700
)

if __name__=="__main__":
    webview.start(debug=True)

    # print(api.get_important()["emails"])