using System;
using System.Net.Mail;

namespace ReportManager
{
    class EmailWriter
    {
        public void GenerateEmail(String body, String subject, String recipient, bool attachFlag)
        {

            SmtpClient client = new SmtpClient("smtp.swri.edu");

            // Specify the e-mail sender.
            // Create a mailing address that includes a UTF8 character
            // in the display name.
            MailAddress from = new MailAddress("biweeklytester@gmail.com", "Report Manager");
            // Set destinations for the e-mail message.
            MailAddress to = new MailAddress(recipient);
            // Specify the message content.
            MailMessage message = new MailMessage(from, to);
            //set body & subject
            message.Body = body;
            message.Subject = subject;

            //check if attachment is needed
            if (attachFlag == true)
            {
                Attach(message);
            }

            // Set the method that is called back when the send operation ends.
            client.SendCompleted += (o, a) => Console.WriteLine("");

            client.SendAsync(message, null);
            Console.WriteLine("Sending message... press c to cancel mail. Press any other key to continue.");
            string answer = Console.ReadLine();
            // Clean up.
            message.Dispose();
            Console.WriteLine("Goodbye.");
        }

        void Attach(MailMessage memo)
        {
            //attach whatever's needed
            //magic sauce
            //memo.Attachments.Add(attachment);
        }
    }
}
