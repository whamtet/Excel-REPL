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
    public static List<Dictionary<String, String>> GetInbox(String user, String password)
    {
        List<Dictionary<String, String>> output = new List<Dictionary<String, String>>();
        using (Imap imap = new Imap())
        {
            imap.ConnectSSL("imap.gmail.com");
            imap.Login(user, password);

            imap.SelectInbox();

            List<long> uids = imap.Search(Flag.All);
            
            foreach (long uid in uids)
            {
                var eml = imap.GetMessageByUID(uid);
                IMail mail = new MailBuilder().CreateFromEml(eml);
                Dictionary<String, String> message = new Dictionary<string, string>();

                StringBuilder fromBuilder = new StringBuilder();
                foreach (var fromItem in mail.From)
                {
                    fromBuilder.Append(fromItem.ToString() + " ");
                }
                message.Add("from", fromBuilder.ToString());

                message.Add("subject", mail.Subject);
                message.Add("text", mail.GetBodyAsText());
                output.Add(message);

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
        //SendMessage(String username, String password, String from, String[] to, String subject, String body)
        SendMessage("whamtet.test@gmail.com", "NeedEmail", new String[] { "whamtet@gmail.com" }, "Hi there", "How are you?");
        Console.ReadLine();
    }
}
    

