using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Threading;


namespace createmail
{
    class Program
    {
        private static string domain;
        private static string pwd;
        private static string smtp;
        private static string configUrl;
        private static int iterations;
        private static int maxItemsPerThread;
        private static int maxThreads;

        static int items = 0;
        static int successful = 0;
        static int warning = 0;
        static int error = 0;
        static int start = 0;
        static int resume = 0;

        static Stopwatch m = new Stopwatch();
        static Exchange exchange;

        static void ParseResults(List<ServiceResponseCollection<ServiceResponse>> results)
        {
            foreach (var coll in results)
            {
                foreach (var res in coll)
                {
                    items++;
                    if (res.Result.ToString().ToLower().Contains("success")) successful++;
                    else
                    if (res.Result.ToString().Contains("warning")) warning++;
                    else
                    if (res.Result.ToString().Contains("error")) error++;
                }
            }
        }

        static void DoWork()
        {
            try
            {
                Console.Clear();
                Console.WriteLine();

                ReadAppSettings();

                Console.WriteLine($"Establishing connection to Exchange server");
                exchange.Connect();
                exchange.folderid = exchange.GetFolderId("INBOX");

                m.Start();

                Console.WriteLine($"Generating {iterations * maxItemsPerThread * maxThreads} emails for user {smtp}");
                Console.WriteLine();

                for (int i = start; i < iterations; i++)
                {
                    Console.WriteLine($"Pass {i + 1} : Saving {maxThreads * maxItemsPerThread} emails under folder Inbox");
                    exchange.MultiThreadedCreateEmails();

                    if (i % 20 == 0 && i != 0)
                    {
                        Console.WriteLine("Sleeping for 60 seconds");
                        Thread.Sleep(60000);
                    }

                    if (i != iterations - 1 && i % 20 != 0)
                    {
                        Console.WriteLine("Sleeping for 10 seconds");
                        Thread.Sleep(10000);
                    }

                    Console.WriteLine();
                    resume = i;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exiting due to {ex.Message}");
                Console.WriteLine($"Sleeping for 5 minutes before restarting ...");
                Thread.Sleep(TimeSpan.FromMinutes(5));
                start = resume;
                DoWork();
            }
        }

        public static void ReadAppSettings()
        {
            iterations = Int32.Parse(ConfigurationManager.AppSettings["Iterations"]);
            maxItemsPerThread = Int32.Parse(ConfigurationManager.AppSettings["MaxItemsPerThread"]);
            maxThreads = Int32.Parse(ConfigurationManager.AppSettings["MaxThreads"]);
            smtp = ConfigurationManager.AppSettings["UserSMTP"];
            pwd = ConfigurationManager.AppSettings["UserPassword"];
            domain = ConfigurationManager.AppSettings["Domain"];
            configUrl = @ConfigurationManager.AppSettings["ConfigUrl"];

            exchange = new Exchange(maxThreads, maxItemsPerThread, domain, smtp, pwd, configUrl);
        }

        public static void Main(string[] args)
        {
            try
            {
                DoWork();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exiting due to {ex.Message}");
            }
            finally
            {
                m.Stop();
                Console.WriteLine();

                //Parse Results
                ParseResults(exchange.results);

                Console.WriteLine($"No. of items written : {items}");
                Console.WriteLine($"No. of 'Success' ops : {successful}");
                Console.WriteLine($"No. of 'Warning' ops : {warning}");
                Console.WriteLine($"No. of 'Error' ops   : {error}");
                Console.WriteLine($"Time taken           : {m.Elapsed.TotalSeconds} seconds");
            }
        }
    }
}
