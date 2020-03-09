using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;
using Task = System.Threading.Tasks.Task;
using RandomDataGenerator.Randomizers;
using RandomDataGenerator.FieldOptions;

namespace createmail
{
    public class Exchange
    {
        private int maxThreads { get; set; }
        private int maxItemsCreatedPerThread { get; set; }
        private string domain { get; set; }
        private string pwd { get; set; }
        private string smtp { get; set; }
        private string configUrl { get; set; }
        public FolderId folderid { get; set; }
        private ExchangeService service { get; set; }

        public List<ServiceResponseCollection<ServiceResponse>> results = new List<ServiceResponseCollection<ServiceResponse>>();

        public Exchange(int maxThreads, int maxItemsPerThread, string domain, string smtp, string pwd, string configUrl)
        {
            this.maxThreads = maxThreads;
            this.maxItemsCreatedPerThread = maxItemsPerThread;
            this.domain = domain;
            this.smtp = smtp;
            this.pwd = pwd;
            this.configUrl = configUrl;
        }

        private bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            bool result = false;

            Uri redirectionUri = new Uri(redirectionUrl);
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }

            return result;
        }

        public void Connect()
        {
            this.service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
            this.service.Credentials = new WebCredentials(smtp, pwd, domain);
            this.service.AutodiscoverUrl(smtp, RedirectionUrlValidationCallback);
        }

        public void MultiThreadedCreateEmails()
        {
            var tasks = new Task[maxThreads];
            for (int i = 0; i < maxThreads; i++)
            {
                tasks[i] = new Task((object param) =>
                {
                    var j = (int)param;
                    Console.WriteLine($"Task {j} started ...");
                    results.Add(this.service.CreateItems(GenerateListOfEmailItems(j), GetFolderId("inbox"), MessageDisposition.SaveOnly, null));
                    Console.WriteLine($"Task {j} exiting ...");
                }, i);
            }

            Parallel.ForEach<Task>(tasks, (t) => { t.Start(); });
            Task.WaitAll(tasks);
        }

        private List<Item> GenerateListOfEmailItems(int seed)
        {
            List<Item> emails = new List<Item>();

            string GenEmailBody()
            {
                return new RandomizerText(new FieldOptionsText
                {
                    Min = 100,
                    Max = 1000,
                    UseSpecial = true,
                    UseLetter = true,
                    UseNumber = true,
                    UseLowercase = true,
                    UseSpace = true
                }).Generate();
            }

            string GenEmailAddress()
            {
                return new RandomizerEmailAddress(new FieldOptionsEmailAddress
                {
                    Male = true,
                    Female = true

                }).Generate();
            }

            for (int x = 0; x < this.maxItemsCreatedPerThread; x++)
            {
                EmailMessage emailMsg = new EmailMessage(this.service);
                emailMsg.From = new EmailAddress(GenEmailAddress());
                emailMsg.ToRecipients.Add(GenEmailAddress());
                //emailMsg.Subject = $"Created-By-Thread{seed}-{new Random((x+1)*10).Next(11111111,99999999)}-{DateTime.Now.Ticks.ToString()+(x+1)*10}";
                emailMsg.Subject = new RandomDataGenerator.Randomizers.RandomizerText(new FieldOptionsText { Min = 32, Max = 64, UseLetter = true, UseLowercase = true,
                    UseNumber = true, UseUppercase = true, UseNullValues = false, UseSpace = false, UseSpecial = false }).Generate();
                emailMsg.Body = GenEmailBody();
                emailMsg.IsRead = false;
                emails.Add(emailMsg);
            }

            return emails;
        }

        public FolderId GetFolderId(string folderName)
        {
            FolderId id = null;

            this.service.Url = new Uri(configUrl);
            FolderView view = new FolderView(1000);
            view.PropertySet = new PropertySet(BasePropertySet.IdOnly);
            view.PropertySet.Add(FolderSchema.DisplayName);
            view.Traversal = FolderTraversal.Deep;
            FindFoldersResults findFolderResults = this.service.FindFolders(WellKnownFolderName.Root, view);

            foreach (Folder f in findFolderResults)
            {
                if (f.DisplayName.ToLower().Equals(folderName.ToLower()))
                {
                    id = f.Id;
                    break;
                }
            }
            return id;
        }
    }
}