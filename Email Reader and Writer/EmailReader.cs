using S22.Imap;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;

namespace ReportManager
{
    class EmailReader
    {

        public void Run()
        {
            using (ImapClient Client = new ImapClient("imap.gmail.com", 993, "biweeklytester@gmail.com", "aq1sw2de3fr4", AuthMethod.Login, true))
            {
                // checks all emails and sends to handlers to check format
                IEnumerable<MailMessage> uids = Client.Search(SearchCondition.All()).Select(uid => Client.GetMessage(uid));

                bool handled = false;

                foreach (MailMessage msg in uids)
                {

                    if (msg.Subject.Equals("Hello"))
                    {
                        Console.WriteLine("Hello message found...");
                        //Console.WriteLine(msg.Body);
                        msg.Dispose();

                    }
                    //else if (msg.Body.ToString().Trim().Equals("hello"))  // \r\n\r\n
                    else if (msg.Subject.Equals("subject"))
                    {
                        Console.WriteLine("subject message found...");
                        msg.Dispose();
                    }
                    else if (msg.Subject.Equals("hi"))
                    {
                        Console.WriteLine("hi message found...");
                        msg.Dispose();
                    }
                    else
                    {
                        Console.WriteLine("msg not found");
                    }


                }

                //EmailWriter write = new EmailWriter();
                //write.GenerateEmail("hello", "subject", "biweeklytester@gmail.com", false);


            }
        }

        public void EmptyInbox()
        {
            // delete specific emails
            using (ImapClient Client = new ImapClient("imap.gmail.com", 993, "biweeklytester@gmail.com", "aq1sw2de3fr4", AuthMethod.Login, true))
            {
                IEnumerable<uint> uids = Client.Search(SearchCondition.All());

                foreach (uint uid in uids)
                {
                    MailMessage msg = Client.GetMessage(uid, FetchOptions.Normal);
                    bool delete = false;

                    // process the message here
                    if (msg.Subject.Equals("subject"))
                    {
                        Console.WriteLine("subject message deleting...");
                        delete = true;

                    }

                    if (delete)
                        Client.DeleteMessage(uid);
                }
            }
        }

        public void mailboxList()
        {
            // move specific emails
            using (ImapClient Client = new ImapClient("imap.gmail.com", 993, "biweeklytester@gmail.com", "aq1sw2de3fr4", AuthMethod.Login, true))
            {
                IEnumerable<uint> uids = Client.Search(SearchCondition.All());

                IEnumerable<string> mailboxes = Client.ListMailboxes();

                foreach (string i in mailboxes)
                {
                    Console.WriteLine(i.ToString());
                }
            }
        }

        public void move()
        {
            // move specific emails
            using (ImapClient Client = new ImapClient("imap.gmail.com", 993, "biweeklytester@gmail.com", "aq1sw2de3fr4", AuthMethod.Login, true))
            {
                IEnumerable<uint> uids = Client.Search(SearchCondition.All());

                foreach (uint uid in uids)
                {
                    MailMessage msg = Client.GetMessage(uid, FetchOptions.Normal);
                    bool move = false;

                    // process the message here
                    if (msg.Subject.Equals("Hello"))
                    {
                        Console.WriteLine("Hello message moving to trash...");
                        move = true;

                    }

                    if (move)
                        Client.MoveMessage(uid, "[Gmail]/Trash");
                }
            }
        }
    }
}
