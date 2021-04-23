function emailsUpdate(sourceUrl,sourceSheet,colUsed,option,targetUrl,targetSheet){
  var CA_socialmedia_gmailThread = GmailApp.search("from:noreply@mailer3.zendesk.com subject:\"CA Daily SVL Tracker Social Media Tickets Report\"", 0, 1)[0];
  var CA_escalcustotix_gmailThread = GmailApp.search("from:noreply@mailer3.zendesk.com subject:\"CA Daily SVL Tracker Escalated Customer Tickets Report\"", 0, 1)[0];
  var CA_custoemail_gmailThread = GmailApp.search("from:noreply@mailer3.zendesk.com subject:\"CA Daily SVL Tracker Customer Email Tickets Report\"", 0, 1)[0];
  var CA_facebooktwitter_gmailThread = GmailApp.search("from:noreply@mailer3.zendesk.com subject:\"CA Daily SVL Tracker 1hr + 30m Facebook Tickets Report\"", 0, 1)[0];
  var CA_facebooktwitter2_gmailThread = GmailApp.search("from:noreply@mailer3.zendesk.com subject:\"CA Daily SVL Tracker 30m Twitter+Facebook Tickets Report\"", 0, 1)[0];

  // Get the attachments in the thread -- since only 1 message is expected in the thread, we will grab the attachments in the first message
  var CA_socialmedia_attachments = CA_socialmedia_gmailThread.getMessages()[0].getAttachments();
  var CA_escalcustotix_attachments = CA_escalcustotix_gmailThread.getMessages()[0].getAttachments();
  var CA_custoemail_attachments = CA_custoemail_gmailThread.getMessages()[0].getAttachments();
  var CA_facebooktwitter_attachments = CA_facebooktwitter_gmailThread.getMessages()[0].getAttachments();
  var CA_facebooktwitter2_attachments = CA_facebooktwitter2_gmailThread.getMessages()[0].getAttachments();
  
  // Get and and parse the CSV from the attachment -- attachment position is still required despite having only one attachment in the thread
  var CA_socialmedia_csv = Utilities.parseCsv(CA_socialmedia_attachments[0].getDataAsString());
  var CA_escalcustotix_csv = Utilities.parseCsv(CA_escalcustotix_attachments[0].getDataAsString());
  var CA_custoemail_csv = Utilities.parseCsv(CA_custoemail_attachments[0].getDataAsString());
  var CA_facebooktwitter_csv = Utilities.parseCsv(CA_facebooktwitter_attachments[0].getDataAsString());
  var CA_facebooktwitter2_csv = Utilities.parseCsv(CA_facebooktwitter2_attachments[0].getDataAsString());
  
  // Send CSV data to their corresponding tabs
  tab_zen_CA_custochat.getRange(1,1,CA_custochat_csv.length,CA_custochat_csv[0].length).setValues(CA_custochat_csv);
  tab_zen_CA_socialmedia.getRange(1,1,CA_socialmedia_csv.length,CA_socialmedia_csv[0].length).setValues(CA_socialmedia_csv);
  tab_zen_CA_escalcustotix.getRange(1,1,CA_escalcustotix_csv.length,CA_escalcustotix_csv[0].length).setValues(CA_escalcustotix_csv);
  tab_zen_CA_custoemail.getRange(1,1,CA_custoemail_csv.length,CA_custoemail_csv[0].length).setValues(CA_custoemail_csv);
  tab_zen_CA_cursucemail.getRange(1,1,CA_cursucemail_csv.length,CA_cursucemail_csv[0].length).setValues(CA_cursucemail_csv);
  tab_zen_CA_cursucescal.getRange(1,1,CA_cursucescal_csv.length,CA_cursucescal_csv[0].length).setValues(CA_cursucescal_csv);
  tab_zen_CA_cursucescal.getRange(1,8,CA_cursucescalTL_csv.length,CA_cursucescalTL_csv[0].length).setValues(CA_cursucescalTL_csv);
  tab_zen_CA_curonbemail.getRange(1,1,CA_curonbemail_csv.length,CA_curonbemail_csv[0].length).setValues(CA_curonbemail_csv);
  tab_zen_CA_facebooktwitter.getRange(2,4,CA_facebooktwitter_csv.length,CA_facebooktwitter_csv[0].length).setValues(CA_facebooktwitter_csv);
  tab_zen_CA_facebooktwitter.getRange(2,1,CA_facebooktwitter2_csv.length,CA_facebooktwitter2_csv[0].length).setValues(CA_facebooktwitter2_csv);
  tab_zen_CA_cursucemail.getRange(1,9,CA_ghostTic_csv.length,CA_ghostTic_csv[0].length).setValues(CA_ghostTic_csv);
  
}