using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReadingApp
{
    class DataLogging
    {
       public List<string> DataToExport = new List<string>();

        public void FileOpener(string Ticket,string TicketType ,string Log)
        {
            DateTime Month = DateTime.Today;
            string month = Month.ToString("MMMM");
            string Year = Month.ToString("yyyy");
            string Date = Month.ToString("MM/dd/yyyy");
            if(File.Exists(@"\\netserver3\Data\Log_Tickets_all\TicketLog"+month+Year+".txt"))
            {
                //File.OpenWrite(@"\\netserver3\Data\Log_Tickets_all\TicketLog" + month + Year + ".txt");
                File.AppendAllText(@"\\netserver3\Data\Log_Tickets_all\TicketLog" + month + Year + ".txt",
                    "\r\n<TicketType> " + TicketType + " </TicketType> "+
                    "\r\n<Ticket> " + Ticket + " </Ticket> "+
                    "\r\n<Date> " + Date + " </Date> " +
                    "\r\n<Log> " + Log + " </Log> ");
            }
            else
            {
                File.Create(@"\\netserver3\Data\Log_Tickets_all\TicketLog" + month + Year + ".txt");
                if (File.Exists(@"\\netserver3\Data\Log_Tickets_all\TicketLog" + month + Year + ".txt"))
                {
                    //File.OpenWrite(@"\\netserver3\Data\Log_Tickets_all\TicketLog" + month + Year + ".txt");
                    File.AppendAllText(@"\\netserver3\Data\Log_Tickets_all\TicketLog" + month + Year + ".txt",
                        "\r\n<Date>" + Date + "</Date>" +
                        "\r\n<TicketType> " + TicketType + " </TicketType> " +
                        "\r\n<Ticket> " + Ticket + " </Ticket> " +
                        "\r\n<Log> " + Log + " </Log> ");
                }
            }
        }
    }
}
