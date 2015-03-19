using HtmlAgilityPack;
using S22.Imap;
using System;
using System.IO;
using System.Net.Mail;

public static class EmailClient
{
    public static ImapClient GetClient(String un, String pw)
    {
        return new ImapClient("imap.gmail.com", 993,
             un, pw, AuthMethod.Login, true);
        
    }
    
}
