﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Word_To_Pdf_RabbitMQ.Producer.Models
{
    public class MessageWordToPdf
    {
        public byte[] WordByte { get; set; }
        public string Email { get; set; }
        public string FileName { get; set; }
    }
}
