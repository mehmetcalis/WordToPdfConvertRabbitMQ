using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Word_To_Pdf_RabbitMQ.Producer.Models
{
    public class WordtoPdf
    {
        public string Email { get; set; }
        public IFormFile WordFile { get; set; }
    }
}
