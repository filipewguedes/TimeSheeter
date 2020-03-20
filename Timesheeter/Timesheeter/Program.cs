using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Reflection;
using System.Runtime.InteropServices;
using System.IO;

namespace Timesheeter
{
    class Program
    {
        const int MONTH = 3;
        const int YEAR = 2020;
        static List<string> activities = new List<string>() { "Daily", "Planning", "Demo", "Retrospective", "Development" };
        static string other = "Other Meetings";

        static Dictionary<int, List<Activity>> daysActivities = new Dictionary<int, List<Activity>>();

        static void Main(string[] args)
        {
            TakeAppointmentsInRange();
            OrganizeActivitiesAndDays();
        }


        private static void TakeAppointmentsInRange()
        {
            Outlook.Application application = GetApplicationObject();

            Outlook.Folder calFolder = application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar)
                                        as Outlook.Folder;

            for (int i = 1; i <= DateTime.DaysInMonth(YEAR, 8); i++)
            {
                daysActivities.Add(i, new List<Activity>());
            }

            for (int i = 1; i <= DateTime.DaysInMonth(YEAR, 8); i++)
            {
                DateTime start = new DateTime(YEAR, MONTH, i);
                DateTime end = start.AddDays(1);
                double hoursExpendInThisDay = 0f;

                if (start.DayOfWeek == DayOfWeek.Saturday || start.DayOfWeek == DayOfWeek.Sunday)
                    continue;


                Outlook.Items rangeAppts = GetAppointmentsInRange(calFolder, start, end);

                if (rangeAppts != null)
                {
                    foreach (Outlook.AppointmentItem appt in rangeAppts)
                    {
                        if (appt.ConversationTopic.Contains("Canceled"))
                            continue;

                        var hours = (appt.End - appt.Start).TotalHours;
                        hoursExpendInThisDay += hours;
                        daysActivities[appt.Start.Day].Add(new Activity(appt.ConversationTopic, hours));
                    }
                }
                
                daysActivities[i].Add(new Activity("Development", 8 - hoursExpendInThisDay));
            }
        }

        private static void OrganizeActivitiesAndDays()
        {            

            double[,] matrix = new double[activities.Count + 1, daysActivities.Count];

            for (int i = 1; i <= DateTime.DaysInMonth(YEAR, 8); i++)
            {
                for (int j = 0; j < activities.Count; j++)
                {
                    matrix[j, i - 1] = daysActivities[i].Where(d => d.Name.Contains(activities[j])).Sum(d => d.Hours);
                }
                
                matrix[activities.Count, i - 1] = daysActivities[i].Where(d => !activities.Any(d.Name.Contains)).Sum(d => d.Hours);
            }

            activities.Add(other);

            var csv = " ;";

            for (int i = 1; i <= DateTime.DaysInMonth(YEAR, 8); i++)
            {
                csv += i + ";";
            }

            csv += "\n";

            for (int i = 0; i < activities.Count ; i++)
            {
                csv += activities[i] + ";";

                for (int j = 0; j < daysActivities.Count; j++)
                {
                    csv += matrix[i, j] + ";";
                }

                csv += "\n";
            }

            csv = csv.Replace(".", ","); //ugly

            var writer = new StreamWriter("saida.csv");
            writer.Write(csv);
            writer.Close();

            }

        /// <summary>
        /// Get recurring appointments in date range.
        /// </summary>
        /// <param name="folder"></param>
        /// <param name="startTime"></param>
        /// <param name="endTime"></param>
        /// <returns>Outlook.Items</returns>
        private static Outlook.Items GetAppointmentsInRange(
            Outlook.Folder folder, DateTime startTime, DateTime endTime)
        {
            string filter = "[Start] >= '"
                + startTime.ToString("g")
                + "' AND [End] <= '"
                + endTime.ToString("g") + "'";
            Debug.WriteLine(filter);
            try
            {
                Outlook.Items calItems = folder.Items;
                calItems.IncludeRecurrences = true;
                calItems.Sort("[Start]", Type.Missing);
                Outlook.Items restrictItems = calItems.Restrict(filter);
                if (restrictItems.Count > 0)
                {
                    return restrictItems;
                }
                else
                {
                    return null;
                }
            }
            catch { return null; }
        }

        private static Outlook.Application GetApplicationObject()
        {

            Outlook.Application application = null;

            // Check if there is an Outlook process running. 
            if (Process.GetProcessesByName("OUTLOOK").Count() > 0)
            {

                // If so, use the GetActiveObject method to obtain the process and cast it to an Application object. 
                application = Marshal.GetActiveObject("Outlook.Application") as Outlook.Application;
            }
            else
            {

                // If not, create a new instance of Outlook and log on to the default profile. 
                application = new Outlook.Application();
                Outlook.NameSpace nameSpace = application.GetNamespace("MAPI");
                nameSpace.Logon("", "", Missing.Value, Missing.Value);
                nameSpace = null;
            }

            // Return the Outlook Application object. 
            return application;
        }
    }

    class Activity
    {
        public Activity(string name, double hours)
        {
            Name = name;
            Hours = hours;

        }
        public string Name { get; set; }
        public double Hours { get; set; }
    }


}
