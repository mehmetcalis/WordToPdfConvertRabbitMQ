using Newtonsoft.Json;
using RabbitMQ.Client;
using RabbitMQ.Client.Events;
using Spire.Doc;
using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Text;

namespace Word_To_Pdf_RabbitMQ.Consumer
{
    class Program
    {

        public static bool EmailSend(string email, MemoryStream memoryStream, string fileName)
        {
            try
            {
                memoryStream.Position = 0;
                System.Net.Mime.ContentType ct = new System.Net.Mime.ContentType(System.Net.Mime.MediaTypeNames.Application.Pdf);
                Attachment attachment = new Attachment(memoryStream, ct);
                attachment.ContentDisposition.FileName = $"{fileName}.pdf";

                SmtpClient sc = new SmtpClient();
                sc.Port = 587;
                sc.Host = "smtp.gmail.com";
                sc.EnableSsl = true;
                sc.UseDefaultCredentials = true;

                sc.Credentials = new NetworkCredential("mc@cmc.com", "1234");

                MailMessage mail = new MailMessage();

                mail.From = new MailAddress("mc@mc.com", "MC");

                mail.To.Add("mc@mc.com");
                

                mail.CC.Add("mc@mc.com");
             

                mail.Subject = "Dönüştürülmüş word dosyası";
                mail.IsBodyHtml = true;
                mail.Body = "Word to Pdf dönüşümü";

                mail.Attachments.Add(attachment);

               

                sc.Send(mail);


                Console.WriteLine($"Sonuç: {email} adresine gönderilmiştir.");
                memoryStream.Close();
                memoryStream.Dispose();
                return true;
            }
            catch (Exception ex)
            {

                Console.WriteLine($"gönderim başarısız olumuştur. Hata: {ex.ToString()}");
                memoryStream.Close();
                memoryStream.Dispose();
                return false;
            }

        }
        static void Main(string[] args)
        {
            bool result = false;
            var factory = new ConnectionFactory();
            factory.Uri = new Uri("amqp://......cloudamqp.com/.......");//Add your rabbitMQ cloud connection string.

            using (var connection = factory.CreateConnection())
            {
                using (var channel = connection.CreateModel())
                {
                    channel.ExchangeDeclare("convert-exchange", type: ExchangeType.Direct, durable: true, autoDelete: false, null);
                    channel.QueueBind("File", "convert-exchange", "WordToPdf", null);

                    channel.BasicQos(0, 1, false);

                    var consumer = new EventingBasicConsumer(channel);

                    channel.BasicConsume("File", false, consumer);

                    consumer.Received += (model, ea) =>
                    {

                        try
                        {
                            Console.WriteLine("Kuyruktan bir mesaj alındı ve işleniyor...");

                            Document doc = new Document();

                            string DeserializeString = Encoding.UTF8.GetString(ea.Body.ToArray());

                            MessageWordToPdf messageWordToPdf = JsonConvert.DeserializeObject<MessageWordToPdf>(DeserializeString);

                            doc.LoadFromStream(new MemoryStream(messageWordToPdf.WordByte), FileFormat.Docx2013);

                            using (MemoryStream ms = new MemoryStream())
                            {
                                doc.SaveToStream(ms, FileFormat.PDF);
                                result = EmailSend(messageWordToPdf.Email, ms, messageWordToPdf.FileName);
                            }

                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Dönüşüm sırasında hata oluştu."+ex.Message);
                            
                        }


                        if (result)
                        {
                            Console.WriteLine("Kuyrukten mesaj başarı ile işlemip, mail gönderimi yapıldı.");
                            channel.BasicAck(ea.DeliveryTag, false);
                        }


                    };

                    Console.WriteLine("Çıkmak için tıklayınız.");
                    Console.ReadLine();

                }
            }

        }
    }
}
