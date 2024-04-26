using static System.Net.Mime.MediaTypeNames;
using System;
using Microsoft.Office.Interop.Outlook;
using System.Text;
using static System.Net.WebRequestMethods;
using Newtonsoft.Json;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Dynamic;
using Spectre.Console;
using static System.Runtime.InteropServices.JavaScript.JSType;
using System.Security.Cryptography;

namespace OutlookFormFinder
{

    internal class Program
    {
        public static readonly List<string> MessageClasses = new List<string>
        {
            "IPM.Activity",
            "IPM.Appointment",
            "IPM.Contact",
            "IPM.DistList",
            "IPM.Document",
            //"IPM.OLE.Class",
            //"IPM",
            "IPM.Note",
            "IPM.Note.IMC.Notification",
            "IPM.Note.Rules.OofTemplate.Microsoft",
            "IPM.Post",
            "IPM.StickyNote",
            "IPM.Recall.Report",
            "IPM.Outlook.Recall",
            "IPM.Remote",
            "IPM.Note.Rules.ReplyTemplate.Microsoft",
            "IPM.Report",
            "IPM.Resend",
            "IPM.Schedule.Meeting.Canceled",
            "IPM.Schedule.Meeting.Request",
            "IPM.Schedule.Meeting.Resp.Neg",
            "IPM.Schedule.Meeting.Resp.Pos",
            "IPM.Schedule.Meeting.Resp.Tent",
            "IPM.Note.Secure",
            "IPM.Note.Secure.Sign",
            "IPM.Task",
            "IPM.TaskRequest.Accept",
            "IPM.TaskRequest.Decline",
            "IPM.TaskRequest",
            "IPM.TaskRequest.Update",
            "IPM.Note.SMIME.MultipartSigned",
            "IPM.Note.StorageQuotaWarning.Warning"
        };

        public static string ByteArrayToHexString(byte [] bytes)
        {
            StringBuilder sb = new StringBuilder("0x"); // Начинаем строку с "0x"
            foreach (byte b in bytes)
            {
                sb.AppendFormat("{0:X2}", b); // Добавляем каждый байт в виде двух шестнадцатеричных цифр
            }
            return sb.ToString();
        }


        private static void WriteLogMessage(string message)
        {
            AnsiConsole.MarkupLine($"[bold grey]LOG:[/] {message}");
        }

        private static void WriteErrorMessage(string message)
        {
            AnsiConsole.MarkupLine($"[bold red]ERROR:[/] {message}");
        }

        private static void WriteWarningMessage(string message)
        {
            AnsiConsole.MarkupLine($"[bold gold1]Warning:[/] {message}");
        }

        private static void WriteGoodMessage(string message)
        {
            AnsiConsole.MarkupLine($"[bold green]{message} [/]");
        }

        private static Spectre.Console.Table CreateTableForHiddenObjects(string MessageClass, bool HasAttach, string ENTRYID)
        {
            Action<TableColumn> column = new Action<TableColumn>(c =>
            {
                c.LeftAligned();
                c.Width(24);
                c.NoWrap();
                c.Padding(0, 1);
            });
            TableColumn tc = new TableColumn("Have Attachments");
            
            var t = new Spectre.Console.Table()
                .Border(TableBorder.HeavyEdge)
                .BorderColor(HasAttach ? Color.DarkRed : Color.DarkOliveGreen1)
                .Width(200)
                .Title(System.String.Format("[magenta2]Listing entry {0}[/]", ENTRYID))                
                .AddColumn("Message Class")
                .AddColumn(tc)
                .AddRow(
                    new Spectre.Console.Text(MessageClass).LeftJustified(), 
                    new Markup(System.String.Format(HasAttach ? "[red]{0}[/]" : "[green]{0}[/]", HasAttach))
                );
            column.Invoke(tc);
            return t;
        }


        static void Main(string [] args)
        {
            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
            Microsoft.Office.Interop.Outlook.Application outlookApp = new Microsoft.Office.Interop.Outlook.Application();
            NameSpace outlookNs = outlookApp.GetNamespace("MAPI");
            MAPIFolder inboxFolder = outlookNs.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
            inboxFolder.Items.IncludeRecurrences = true;
            AnsiConsole.MarkupLine($"[bold green]Hidden items in Inbox {inboxFolder.Items.Count} [/]");
            
            Store store = outlookNs.DefaultStore;
            // Получаем доступ к коллекции правил
            Rules rules = store.GetRules();

            // Получаем доступ к содержимому папки через объект Table
            // https://docs.microsoft.com/en-us/office/vba/api/outlook.table
            // filter - строка фильтрации сообщений, но в данном случае не используется - все скрытые сообщения в папке
            // и так не являются сообщениями типа IPM.Note
            // но фильтр можно использовать для других целей
            string filter = "[MessageClass] <> 'IPM.Note'";
            Microsoft.Office.Interop.Outlook.Table table = inboxFolder.GetTable(filter, OlTableContents.olHiddenItems);

            table.Columns.RemoveAll();
            //PR_MESSAGE_CLASS
            table.Columns.Add("http://schemas.microsoft.com/mapi/proptag/0x001A001E");
            //PR_HASATTACH
            table.Columns.Add("http://schemas.microsoft.com/mapi/proptag/0x0E1B000B");
            //PR_ENTRYID
            table.Columns.Add("http://schemas.microsoft.com/mapi/proptag/0x0FFF0102");


            AnsiConsole.Record();

            AnsiConsole.Status()
                .AutoRefresh(true)
                .Spinner(Spinner.Known.Default)
                .SpinnerStyle(Style.Parse("yellow"))
                .Start(System.String.Format("[yellow]Searching for hidden messages with attachments[/]"), ctx =>
                {
                    int count = 0;
                    while (!table.EndOfTable)
                    {
                        Row nextRow = table.GetNextRow();
                        // Получаем значения столбцов
                        // обращаеем внимание на dynamic - это позволяет работать с объектами без указания типа
                        // в данном случае это удобно, так как мы не знаем типы значений столбцов
                        dynamic d = nextRow.GetValues();

                        string messageClass = d [0];
                        bool hasAttachment = d [1];
                        string PR_ENTRYID = ByteArrayToHexString(d [2]);
                        // GetItemFromID - get item by PR_ENTRYID
                        dynamic item = outlookNs.GetItemFromID(PR_ENTRYID.Replace("0x", ""));

                        AnsiConsole.Write(CreateTableForHiddenObjects(messageClass, hasAttachment, PR_ENTRYID));

                        if (hasAttachment)
                        {
                            // GetAttachments
                            Attachments attachments = item.Attachments;
                            if (attachments.Count > 0)
                            {
                                // Create directory
                                string newDirectoryPath = Path.Combine(baseDirectory, messageClass, PR_ENTRYID);
                                Directory.CreateDirectory(newDirectoryPath);
                                AnsiConsole.MarkupLine($"[bold red]Created dir for attach {newDirectoryPath} [/]");
                                foreach (Attachment attachment in attachments)
                                {
                                    AnsiConsole.MarkupLine($"[bold red]Saved attach to: {attachment.FileName}[/]");
                                    attachment.SaveAsFile(newDirectoryPath + "\\" + attachment.FileName);
                                }
                            }
                        }
                        count++;
                    }
                });
            System.IO.File.WriteAllText("output.html", AnsiConsole.ExportHtml());
        }
    }
}
   