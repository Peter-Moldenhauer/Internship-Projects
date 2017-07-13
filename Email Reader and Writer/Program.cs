using ReportManager;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            //EmailReader read = new EmailReader();
            //read.Run();
            //read.EmptyInbox();
            //read.mailboxList();
            //read.move(); 

            // create an EmailWriter object
            EmailWriter write = new EmailWriter();  

            // loop to generate 10 emails with different subject lines 
            for(var i = 1; i <= 10; i++)
            {
                // create unique subject line in each email 
                var emailNumber = i.ToString();
                var emailSubject = "email" + emailNumber;

                // generate email 
                write.GenerateEmail("bodyText bodyText", emailSubject, "biweeklytester@gmail.com", false);
            }
            
        }
    }
}
