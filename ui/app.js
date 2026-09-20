const right_panel_title = document.getElementById("right_panel_title");
const list_mail = document.getElementById("mail_list");

const show_mails = document.getElementById("show_mails");
const show_details = document.getElementById("show_details");




const loadEmails = async (load_mail) => {

    document.querySelector(".email-compose-panel").style.display = "none";
    let result;

    if (load_mail === "INBOX"){
        result = await window.pywebview.api.get_inbox();
    }
    else if(load_mail === "SEND MAIL"){
        result = await window.pywebview.api.get_sent();
    }
    else if(load_mail === "IMPORTANT"){
        result = await window.pywebview.api.get_important();
    }
    else {
        console.error("Unknown mail type:", load_mail);
        return;
    }
    console.log(result);

    if (!result.success) {
        console.error(result.message);
        return;
    }

    const list = result.emails;

    list_mail.innerHTML = "";

    list.forEach((items,mail_idx) => {

        const div = document.createElement("div");
        div.className = "mail_card";
        
        const mail_sender = document.createElement("h2");
        const mail_content = document.createElement("p");
        const mail_time = document.createElement("span");

        mail_sender.textContent = (load_mail==="SEND MAIL")? items.to:items.sender;

        mail_content.innerHTML = `<b>${items.subject}</b> - ${items.body.substring(0, 150)}`;
        mail_content.style.display = "inline";

        mail_time.textContent = items.timestamp || items.sent;

        div.appendChild(mail_sender);
        div.appendChild(mail_content);
        div.appendChild(mail_time);
        div.onclick = ()=>{
            show_mail_details(items,load_mail, mail_idx);
        }
        list_mail.appendChild(div);
    });
};

const set_right_panel_title=(title_name)=>{
    
    right_panel_title.innerHTML = title_name;
}

const left_panel_selection = (index_name)=>{
    set_right_panel_title(index_name)
    if(show_mails.style.display == "none"){
        show_mails.style.display = "block";
        show_details.style.display = "none"
    }
    
    loadEmails(index_name);
    
}


// details panel
const show_mail_details = (mail,load_mail, index)=>{
        show_mails.style.display = "none";
        show_details.style.display = "block"

        if(load_mail==="SEND MAIL"){
            document.getElementById("reply_suggester").style.display ="none";
        }
        else{
            document.getElementById("reply_suggester").style.display ="block";
            display_replys(mail);
        }


        console.log(mail);

        const subject_card = document.getElementById("subject_panel");
        const sender_details = document.getElementById("sender_details");

        subject_card.textContent = "Subject: "+mail.subject;

        // const send_name = document.createElement("div");

        sender_details.innerHTML = `<b>${(!mail.to)?"FROM: ":"TO: " }</b> ${(mail.to)?mail.to:mail.sender} <p> &lt;${mail.sender_email || mail.to}&gt;  ${mail.timestamp || mail.sent}</p>`;

        // sender_details.appendChild(send_name);

        const message_body = document.getElementById("message_body");

        message_body.innerHTML = mail.html_body || mail.body;

        const attachment_panel = document.getElementById("attachment_panel");
        attachment_panel.innerHTML = `<h3>Attachments:</h3>`;

        mail.attachments.forEach((attachment) => {
            
            const attachment_div = document.createElement("div");
            attachment_div.className = "attachment";

            const file_name = document.createElement("span");
            file_name.textContent = attachment.name;

            const download_btn = document.createElement("button");
            download_btn.textContent = "Download";

            download_btn.onclick = () => {
                downloadAttachment(index, attachment.index);
            };

            attachment_div.appendChild(file_name);
            attachment_div.appendChild(download_btn);

            attachment_panel.appendChild(attachment_div);
        });
}

const display_replys = async (message) =>{

    const suggestion_box = document.getElementById("suggest_box");
    suggestion_box.replaceChildren();

    const replys =  await window.pywebview.api.reply_generate("employee",message.subject,message.body);

    Object.values(replys).forEach((items) => {
        const suggest_items = document.createElement("div");

        suggest_items.className = "suggest_items";

        suggest_items.innerHTML = `<p>${items.body}</p>`;

        suggest_items.onclick = ()=>{
            openCompose(message.sender_email, items.subject,items.body);
        }

        suggestion_box.appendChild(suggest_items);


    });

    console.log(replys);


}

const downloadAttachment = async (current_email_index,attachmentIndex) => {

    const result = await pywebview.api.download_attachment(
        current_email_index,
        attachmentIndex
    );

    if (result.success) {
        showDownloadPopup(
            `"${result.name}" downloaded successfully`
        );
    } else {
        showDownloadPopup(
            "Download failed: " + result.message
        );
    }
};



let selectedFiles = [];


/*
    Call this function when opening
    the email sending panel.
*/
function openCompose(to, subject, body) {

    document.querySelector(".email-compose-panel").style.display = "block";

    document.getElementById("mail-to").value = to;

    document.getElementById("mail-subject").value = subject;

    document.getElementById("mail-body").value = body;

    selectedFiles = [];

    document.getElementById("attachment-list").innerHTML = "";
}


/*
    Open native file picker
*/
async function chooseAttachment() {

    const files = await window.pywebview.api.select_files();

    if (!files) {
        return;
    }

    selectedFiles = files;

    displayAttachments();
}


/*
    Display selected attachments
*/
function displayAttachments() {

    const list = document.getElementById("attachment-list");

    list.innerHTML = "";

    selectedFiles.forEach((file, index) => {

        const div = document.createElement("div");

        div.className = "attachment-item";

        const fileName = file.split("\\").pop();

        div.innerHTML = `
            <span>📎 ${fileName}</span>

            <button
                class="remove-attachment"
                onclick="removeAttachment(${index})">
                ✕
            </button>
        `;

        list.appendChild(div);
    });
}


/*
    Remove attachment
*/
function removeAttachment(index) {

    selectedFiles.splice(index, 1);

    displayAttachments();
}


/*
    Send email
*/
async function sendEmail() {

    const to = document.getElementById("mail-to").value;
    const subject = document.getElementById("mail-subject").value;
    const body = document.getElementById("mail-body").value;

    if (!to) {
        alert("Recipient is required");
        return;
    }

    if (!subject) {
        alert("Subject is required");
        return;
    }

    const result = await window.pywebview.api.send_email(
        to,
        subject,
        body,
        selectedFiles
    );

    if (result.success) {
        alert(result.message);
    } else {
        alert("Failed: " + result.message);
    }
}







// popup function

const showDownloadPopup = (message) => {

    const popup = document.getElementById("download_popup");
    const message_box = document.getElementById("download_message");

    message_box.textContent = message;

    popup.style.display = "block";

    setTimeout(() => {
        popup.style.display = "none";
    }, 3000);
};

// sync outlook
const refreshInbox = async () => {

    const sync = await pywebview.api.sync_outlook();

    console.log(sync);


};





window.addEventListener("pywebviewready", function() {
    console.log("python connected");
    show_details.style.display = "none";
    refreshInbox();
    loadEmails("INBOX");
});

// window.addEventListener("DOMContentLoaded", function() {
//     console.log("python connected")
//     loadEmails();
// });



