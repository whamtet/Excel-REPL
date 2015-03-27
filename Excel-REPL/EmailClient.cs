using Limilabs.Client.IMAP;
using Limilabs.Client.SMTP;
using Limilabs.Mail;
using Limilabs.Mail.Headers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Mail;
using System.Text;

public static class EmailClient
{
    public static List<Dictionary<Object, Object>> GetInbox(String user, String password)
    {
        List<Dictionary<Object, Object>> output = new List<Dictionary<Object, Object>>();
        using (Imap imap = new Imap())
        {
            imap.ConnectSSL("imap.gmail.com");
            imap.Login(user, password);

            imap.SelectInbox();
            
            List<long> uids = imap.Search(Flag.All);
            int numSelected = 0;
            foreach (long uid in uids)
            {

                var eml = imap.GetMessageByUID(uid);
                IMail mail = new MailBuilder().CreateFromEml(eml);

                Dictionary<Object, Object> message = new Dictionary<object, object>();
                List<Object> fromObjects = new List<Object>();
                foreach (var fromItem in mail.From)
                {
                    Dictionary<Object, Object> fromObject = new Dictionary<object, object>();
                    fromObject.Add("address", fromItem.Address);
                    fromObject.Add("name", fromItem.Name);
                    fromObjects.Add(fromObject);
                }
                message.Add("from", fromObjects);

                message.Add("subject", mail.Subject);
                message.Add("text", mail.GetBodyAsText());
                output.Add(message);

                numSelected++;
                if (numSelected == 100)
                {
                    break;
                }
            }
            imap.Close();
        }
        
        return output;
    }
    public static void SendMessage(String username, String password, String[] to, String subject, String body)
    {
        MailBuilder builder = new MailBuilder();
        builder.From.Add(new MailBox(username));
        foreach (String t in to)
        {
            builder.To.Add(new MailBox(t));
        }
        
        builder.Subject = subject;
        builder.Text = body;

        IMail email = builder.Create();

        // C#

        using (Smtp smtp = new Smtp())
        {
            smtp.Connect("smtp.gmail.com");    // or ConnectSSL for SSL
            smtp.UseBestLogin(username, password); // remove if authentication is not needed

            ISendMessageResult result = smtp.SendMessage(email);
            if (result.Status == SendMessageStatus.Success)
            {
                Console.WriteLine("Sent");
            }

            smtp.Close();
        }

    }
    public static void Main(String[] args)
    {
        GetInbox("whamtet@gmail.com", "NeedaEmail_1");
    }
}
    

