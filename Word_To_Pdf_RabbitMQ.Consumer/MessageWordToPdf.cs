using System;
using System.Collections.Generic;
using System.Text;

namespace Word_To_Pdf_RabbitMQ.Consumer
{
    class MessageWordToPdf
    {
        public byte[] WordByte { get; set; }
        public string Email { get; set; }
        public string FileName { get; set; }
    }
}
