using System;
using System.Collections.Generic;
using System.IO;

namespace OpenSDKTools.Examples
{
    class Program
    {
        static void Main(string[] args)
        {
            var template = new FileInfo(@"Templates\sample_template.docx");
            var doc = template.CopyTo($@"Output\sample_template_{Guid.NewGuid()}.docx");
            var service = new Word.Service(doc);
            service.AddVariables(new List<Word.Variable>
            {
                new Word.Variable
                {
                    Key = "FullName",
                    Value = "John Doe"
                },
                new Word.Variable
                {
                    Key = "Appointment.Date",
                    Value = DateTime.Now.AddDays(5).ToShortDateString()
                },
                new Word.Variable
                {
                    Key = "Appointment.StartTime",
                    Value = DateTime.Now.AddDays(5).ToShortTimeString()
                },
                new Word.Variable
                {
                    Key = "Appointment.Rooms",
                    Value = "Room A, Room B"
                }
            });


            var phoneNumbers = new Container
            {
                Key = "Appointment.PhoneNumbers"
            };

            var p = new Paragraph();
            p.AddText(new Text() { Value = "040 260 1200" });
            p.AddBreak();
            p.AddText(new Text() { Value = "123" });

            phoneNumbers.AddChild(p);

            service.AddContainer(phoneNumbers);

            service.Generate();
        }
    }
}
