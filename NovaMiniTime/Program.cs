﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using System.Configuration;
using System.IO;

namespace NovaMiniTime
{
    class Program
    {
        static void Main(string[] args)
        {
            var siteUrl = ConfigurationManager.AppSettings["site"];
            using (var context = new ClientContext(siteUrl))
            {
            



                Web site = context.Web;
                var pwd = ConfigurationManager.AppSettings["password"];
                SecureString passWord = new SecureString();
                foreach (char c in pwd) passWord.AppendChar(c);
                context.Credentials = new SharePointOnlineCredentials(ConfigurationManager.AppSettings["username"], passWord);
                List targetList = site.Lists.GetByTitle(ConfigurationManager.AppSettings["list"]);
                CamlQuery query = new CamlQuery();
                query.ViewXml = @"<View><Query><Where><And><Geq><FieldRef Name='Date' /><Value IncludeTimeValue='TRUE' Type='DateTime'>" + DateTime.Now.AddDays(-100).ToString("yyyy-MM-ddTHH:mm:ssZ") + "</Value></Geq><Leq><FieldRef Name='Date' /><Value IncludeTimeValue='TRUE' Type='DateTime'>" + DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ssZ") + "</Value></Leq></And></Where></Query><RowLimit>10</RowLimit></View>";
                ListItemCollection collListItem = targetList.GetItems(query);

                context.Load(collListItem);
                context.ExecuteQuery();
                foreach (var targetListItem in collListItem.GroupBy(x => ((FieldLookupValue)x["Project_x0020__x002d__x0020_Cust0"]).LookupId))
                {
                    Console.WriteLine(((FieldLookupValue)targetListItem.FirstOrDefault()["Project_x0020__x002d__x0020_Cust0"]).LookupValue + " = " + ((FieldLookupValue)targetListItem.FirstOrDefault()["Project_x0020__x002d__x0020_Cust0"]).LookupId);
                }
                var ht = HoursToday();

                Console.WriteLine("How many hours have you worked today?");
                var hoursToday = ConsoleReadLineWithDefault(ht.Hours);
                if (hoursToday != ht.Hours)
                {
                    ht.EndTime = ht.StartTime.AddHours(double.Parse(hoursToday));
                }
                Console.WriteLine("Which project id have you worked on?");
                var projectId = ConsoleReadLineWithDefault(((FieldLookupValue)collListItem.FirstOrDefault()["Project_x0020__x002d__x0020_Cust0"]).LookupId.ToString());
                Console.WriteLine("Write a comment for the time registration?");
                var comment = Console.ReadLine();

                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem newItem = targetList.AddItem(itemCreateInfo);
                newItem["Date"] = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ssZ");
                newItem["Hours"] = hoursToday.Replace(",", ".");
                newItem["Project_x0020__x002d__x0020_Cust0"] = projectId;
                newItem["Comments"] = comment;
                //newItem["From"] = ht.StartTime.ToString("MM-dd-yyyy HH:mm"); Time Format is weird, maybe American
                //newItem["To"] = ht.EndTime.ToString("MM-dd-yyyy HH:mm");
                newItem["Person2"] = ((FieldLookupValue)collListItem.FirstOrDefault()["Person2"]).LookupId.ToString();
                newItem.Update();

                context.ExecuteQuery();
            }
        }

        public static string ConsoleReadLineWithDefault(string defaultValue)
        {
            System.Windows.Forms.SendKeys.SendWait(defaultValue);
            return Console.ReadLine();
        }

        public static WorkHours HoursToday()
        {
            WorkHours wh = new WorkHours();
            double hoursToday = 0;
            string logType = "System";

            //use this if your are are running the app on the server
            EventLog ev = new EventLog(logType, System.Environment.MachineName);

            //use this if you are running the app remotely
            // EventLog ev = new EventLog(logType, "[youservername]");

            if (ev.Entries.Count <= 0)
                Console.WriteLine("No Event Logs in the Log :" + logType);

            // Loop through the event log records. 
            for (int i = ev.Entries.Count - 1; i >= 0; i--)
            {
                EventLogEntry CurrentEntry = ev.Entries[i];

                if (CurrentEntry.InstanceId == 30 && CurrentEntry.TimeGenerated.Date == DateTime.Today)
                {
                    var whhours =
                        Math.Round((DateTime.Now - CurrentEntry.TimeGenerated).TotalHours*2,
                            MidpointRounding.AwayFromZero)/2;
                    wh.Hours = whhours.ToString();
                    wh.StartTime = RoundToNearest(CurrentEntry.TimeGenerated, TimeSpan.FromMinutes(15));
                    wh.EndTime = wh.StartTime.AddHours(whhours);
                }
            }
            ev.Close();
            return wh;
        }

        public static DateTime RoundToNearest(DateTime dt, TimeSpan d)
        {
            var delta = dt.Ticks % d.Ticks;
            bool roundUp = delta > d.Ticks / 2;
            var offset = roundUp ? d.Ticks : 0;

            return new DateTime(dt.Ticks + offset - delta, dt.Kind);
        }

        public class WorkHours
        {
            public string Hours { get; set; }
            public DateTime StartTime { get; set; }
            public DateTime EndTime { get; set; }
        }
    }
}
