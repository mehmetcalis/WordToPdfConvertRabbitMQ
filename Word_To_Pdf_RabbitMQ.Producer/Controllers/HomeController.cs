using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using RabbitMQ.Client;
using Word_To_Pdf_RabbitMQ.Producer.Models;

namespace Word_To_Pdf_RabbitMQ.Producer.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IConfiguration _configuration;


        public HomeController(ILogger<HomeController> logger, IConfiguration configuration)
        {
            _logger = logger;
            _configuration = configuration;
        }

        public IActionResult WordtoPdf()
        {
            return View();
        }

        [HttpPost]
        public IActionResult WordtoPdf(WordtoPdf wordtopdf)
        {
            try
            {
                var factory = new ConnectionFactory();
                factory.Uri = new Uri(_configuration.GetConnectionString("RabbitMQCloudConnectionString"));

                using (var connection = factory.CreateConnection())
                {
                    using (var channel = connection.CreateModel())
                    {
                        channel.ExchangeDeclare("convert-exchange", type: ExchangeType.Direct, true, false, null);

                        channel.QueueDeclare(queue: "File", true, false, false, null);
                        channel.QueueBind(queue: "File", exchange: "convert-exchange", routingKey: "WordToPdf", arguments: null);

                        MessageWordToPdf messagewordtopdf = new MessageWordToPdf();

                        using (MemoryStream ms = new MemoryStream())
                        {
                            wordtopdf.WordFile.CopyTo(ms);
                            messagewordtopdf.WordByte = ms.ToArray();
                        }
                        messagewordtopdf.Email = wordtopdf.Email;
                        messagewordtopdf.FileName = Path.GetFileNameWithoutExtension(wordtopdf.WordFile.FileName.ToString());

                        string serializeMessage = JsonConvert.SerializeObject(messagewordtopdf);

                        byte[] ByteMessage = Encoding.UTF8.GetBytes(serializeMessage);


                        var properties = channel.CreateBasicProperties();

                        properties.Persistent = true;

                        channel.BasicPublish("convert-exchange", "WordToPdf", properties, ByteMessage);

                        ViewBag.result = "Word dosyanızın PDF dönüşümü yapılmasından sonra email ahesabınıza gönderilecektir.";

                        return View();

                    }
                }
            }
            catch (Exception ex)
            {

                return View("Error", ex);
            }

        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
