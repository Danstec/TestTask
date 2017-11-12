function setSendEmailNotification() {
    var isdocumentsubmitted = Xrm.Page.getAttribute("stec_email_send_date").getValue();
    if (isdocumentsubmitted != null)
        Xrm.Page.ui.setFormNotification("Письмо успешно отправлено", "INFO", "2001");
}