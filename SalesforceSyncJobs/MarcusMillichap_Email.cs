using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net.Mail;
using System.Text;

namespace SalesforceSyncJobs
{
    public class MarcusMillichap_Email
    {
        private string _to, _from, _subject, _message;
        private bool _isHtml, _isMultipleTo;
        private List<string> _multipleTo;

        public string To { get { return _to; } set { _to = value; } }
        public string From { get { return _from; } set { _from = value; } }
        public string Subject { get { return _subject; } set { _subject = value; } }
        public string Message { get { return _message; } set { _message = value; } }
        public bool IsHtml { get { return _isHtml; } set { _isHtml = value; } }
        public bool IsMultipleTo { get { return _isMultipleTo; } set { _isMultipleTo = value; } }
        public List<string> MultipleTo { get { return _multipleTo; } set { _multipleTo = value; } }

        public MarcusMillichap_Email()
        {
            _to = String.Empty;
            _from = String.Empty;
            _subject = String.Empty;
            _message = String.Empty;
            _isHtml = true;
            _isMultipleTo = false;
            _multipleTo = null;
        }

        public MarcusMillichap_Email(string to, string from, string subject, string message, bool isHtml)
        {
            _to = to;
            _from = from;
            _subject = subject;
            _message = message;
            _isHtml = isHtml;
            _multipleTo = null;
            _isMultipleTo = false;
        }

        public MarcusMillichap_Email(List<string> to, string from, string subject, string message, bool isHtml)
        {
            _from = from;
            _subject = subject;
            _message = message;
            _isHtml = isHtml;
            _multipleTo = to;
            _isMultipleTo = true;
            _to = String.Empty;
        }

        public bool SendEmail()
        {
            if (!String.IsNullOrEmpty(_to))
            {
                bool saved = false;
                try
                {
                    //initialize object
                    MailMessage mail = new MailMessage();
                    //to
                    if (!_isMultipleTo) { mail.To.Add(_to); }
                    else
                    {
                        if (_multipleTo != null)
                        {
                            foreach (string s in _multipleTo)
                            {
                                mail.To.Add(s);
                            }
                        }
                        else return false;
                    }
                    //from
                    mail.From = new MailAddress(_from);
                    //subject
                    mail.Subject = _subject;
                    //is html
                    mail.IsBodyHtml = _isHtml;
                    //body message
                    mail.Body = _message;
                    //create smpt object
                    SmtpClient smtp = new SmtpClient(ConfigurationManager.AppSettings["MailServer"].ToString());
                    smtp.UseDefaultCredentials = true;
                    //send email
                    smtp.Send(mail);
                }
                catch { return false; }
                finally
                {
                }
                return saved;
            }
            else return false;
        }

        public static string GenerateErrorBulkMessage(List<MarcusMillichap_Errors> errors, string eMessage = "")
        {
            string message = string.Empty;
            if (errors != null && errors.Count > 0)
            {
                foreach (MarcusMillichap_Errors e in errors)
                {
                    message += "<br /><br />Salesforce ID: <br/><br/>" + e.SalesforceId;
                    message += "<br /><br />External ID: <br/><br/>" + e.ExternalId;
                    if (!string.IsNullOrEmpty(eMessage)) message += "<br /><br />Extra Error Message: <br /><br />" + eMessage;
                    message += "<br /><br />Error Message: <br /><br />" + e.ErrorMessage;
                    if (!string.IsNullOrEmpty(e.EmailAddress)) message += "<br /><br />Email Address: <br /><br />" + e.EmailAddress;
                    if (!string.IsNullOrEmpty(e.FirstName)) message += "<br /><br />First Name: <br /><br />" + e.FirstName;
                    if (!string.IsNullOrEmpty(e.LastName)) message += "<br /><br />Last Name: <br /><br />" + e.LastName;
                }
            }
            return message;
        }
    }
}
