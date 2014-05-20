using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Web;
using System.Text.RegularExpressions;
using System.Net.Mail;

namespace TPDailyRecapLight
{
    class Program
    {
        private const string PathToTp = "https://targetprocess.aemedia.ru/TargetProcess2/api/v1/";
        //private const string ReportPath = @"C:\Users\dmitry.mironov\Documents\Visual Studio 2013\Projects\dailyrecap\dailyrecap\";
        private const string ReportPath = @"C:\Users\dmitry.mironov\Documents\Visual Studio 2013\Projects\DailyRecapLight\";

        static void Main(string[] args)
        {
            //create web client to get response from webserver
            var client = new WebClient();
            client.UseDefaultCredentials = true;
            client.Encoding = Encoding.UTF8;

            //create Streamwriter to write html to report file
            //open report top part
            StreamWriter reportFile = new StreamWriter(@"c:\temp\report.html", false, Encoding.UTF8);

            //add report header
            StreamReader sr = new StreamReader(ReportPath + "Report_Header_Fixed.html");

            //update report date in the output report file
            string content = sr.ReadToEnd();
            content = content.Replace("##ReportDate##", reportDate().ToString("dddd dd MMMM", CultureInfo.CreateSpecificCulture("ru-RU")));
            reportFile.Write(content);

            StartBuidingReport(client, reportDate(), reportFile);

            //add report footer
            sr = new StreamReader(ReportPath + "Report_Footer_Fixed.html");
            content = sr.ReadToEnd();
            reportFile.Write(content);

            //save report file and dispose
            reportFile.Close();
            reportFile.Dispose();

            //SendEmail();
        }

        private static void StartBuidingReport(WebClient client, DateTime reportDate, StreamWriter sw)
        {
            //get collection of projects
            ProjectsCollection projectsCollection = JsonConvert.DeserializeObject<ProjectsCollection>(client.DownloadString(PathToTp + "projects?include=[name,id,owner]&where=IsActive%20eq%20%27true%27&take=1000&format=json"));

            //project acid property
            String acid = "";

            Console.WriteLine("Total projects: " + projectsCollection.Items.Count());
            //check that projec collection has at least one project
            if (projectsCollection.Items.Count > 0)
            {
                //for each project check if there are any  userstories or bugs modified on this report date
                foreach (Project project in projectsCollection.Items)
                {
                    StreamReader sr = new StreamReader(ReportPath + "Project_Header_Fixed.html");
                    String content = sr.ReadToEnd();


                    UserStoriesCollection userStoriesCollection = JsonConvert.DeserializeObject<UserStoriesCollection>(client.DownloadString(PathToTp + "UserStories?include=[Id,Name,Description,StartDate,EndDate,CreateDate,ModifyDate,Effort,EffortCompleted,EffortToDo,Owner[id,FirstName,LastName],EntityState[id,name],Feature[id,name],Assignments[Role,GeneralUser[id,FirstName,LastName]]]&where=(Project.Id eq " + project.Id + ") and (ModifyDate eq '" + reportDate.ToString("yyyy-MM-dd") + "')&take=1000&format=json"));
                    BugsCollection bugsCollection = JsonConvert.DeserializeObject<BugsCollection>(client.DownloadString(PathToTp + "Bugs?include=[Id,Name,Description,StartDate,EndDate,CreateDate,ModifyDate,Effort,EffortCompleted,EffortToDo,Owner[id,FirstName,LastName],EntityState[id,name],Assignments[Role,GeneralUser[id,FirstName,LastName]]]&where=(Project.Id eq " + project.Id + ") and (ModifyDate eq '" + reportDate.ToString("yyyy-MM-dd") + "')&take=1000&format=json"));

                    //check that we got at least one user story
                    if (userStoriesCollection.Items.Count > 0 || bugsCollection.Items.Count > 0)
                    {
                        //get project acid
                        ProjectContext projectContext = JsonConvert.DeserializeObject<ProjectContext>(client.DownloadString(PathToTp + "Context/?ids=" + project.Id + "&format=json"));  
                      
                        //Add report section to output file                        
                        content = content.Replace("##ProjectName##", project.Name);
                        //add project owner
                        content = content.Replace("##ProjectOwner##", project.Owner.ToString());
                        //TO-DO: Add project owner to the report
                        //content = content.Replace("##ProjectOwner##", project.Owner.ToString());
                        sw.Write(content);
                        sr.Close();

                        //process project and userstories
                        //get stories lists by status                        
                        try
                        {
                            List<UserStory> uc_Completed = userStoriesCollection.Items.Where(us => (us.EndDate.HasValue ? us.EndDate.Value.Date.Equals(reportDate.Date) : false)).ToList<UserStory>();
                            List<Bug> bugs_Completed = bugsCollection.Items.Where(bug => (bug.EndDate.HasValue ? bug.EndDate.Value.Date.Equals(reportDate.Date) : false)).ToList<Bug>();
                            if (uc_Completed.Count > 0 || bugs_Completed.Count > 0)
                            {
                                //add Section to report

                                sr = new StreamReader(ReportPath + "Section_Header_Fixed.html");
                                content = sr.ReadToEnd();
                                content = content.Replace("##SectionName##", "Выполнено");
                                sw.Write(content);
                                sr.Close();
                                if (uc_Completed.Count > 0)
                                {
                                    AddRecordsToReport(uc_Completed, sw, projectContext.Acid);
                                }
                                if (bugs_Completed.Count > 0)
                                {
                                    AddRecordsToReport(bugs_Completed, sw, projectContext.Acid);
                                }

                                //add status footer
                                sr = new StreamReader(ReportPath + "Section_Footer_Fixed.html");
                                content = sr.ReadToEnd();
                                sw.Write(content);
                                sr.Close();
                            }
                        }
                        catch (Exception e)
                        {

                        }

                        try
                        {
                            List<UserStory> uc_Modified = userStoriesCollection.Items.Where(us => (us.ModifyDate.HasValue ? us.ModifyDate.Value.Date.Equals(reportDate.Date) : false) && us.EntityState.Name != "Done" && us.CreateDate.Value.Date.CompareTo(reportDate.Date) != 0).ToList<UserStory>();
                            List<Bug> bugs_Modified = bugsCollection.Items.Where(bug => (bug.ModifyDate.HasValue ? bug.ModifyDate.Value.Date.Equals(reportDate.Date) : false) && bug.EntityState.Name != "Done" && bug.CreateDate.Value.Date.CompareTo(reportDate.Date) != 0).ToList<Bug>();
                            if (uc_Modified.Count > 0 || bugs_Modified.Count > 0)
                            {
                                //add Section to report
                                sr = new StreamReader(ReportPath + "Section_Header_Fixed.html");
                                content = sr.ReadToEnd();
                                content = content.Replace("##SectionName##", "Изменено");
                                sw.Write(content);
                                sr.Close();
                                if (uc_Modified.Count > 0)
                                {
                                    AddRecordsToReport(uc_Modified, sw, projectContext.Acid);
                                }
                                if (bugs_Modified.Count > 0)
                                {
                                    AddRecordsToReport(bugs_Modified, sw, projectContext.Acid);
                                }

                                //add status footer
                                sr = new StreamReader(ReportPath + "Section_Footer_Fixed.html");
                                content = sr.ReadToEnd();
                                sw.Write(content);
                                sr.Close();
                            }
                        }
                        catch (Exception e)
                        {

                        }

                        try
                        {
                            List<UserStory> uc_Added = userStoriesCollection.Items.Where(us => (us.EndDate.HasValue ? false : us.CreateDate.Value.Date.Equals(reportDate.Date))).ToList<UserStory>();
                            List<Bug> bugs_Added = bugsCollection.Items.Where(bug => (bug.EndDate.HasValue ? false : bug.CreateDate.Value.Date.Equals(reportDate.Date))).ToList<Bug>();

                            if (uc_Added.Count > 0 || bugs_Added.Count > 0)
                            {
                                //add Section to report
                                sr = new StreamReader(ReportPath + "Section_Header_Fixed.html");
                                content = sr.ReadToEnd();
                                content = content.Replace("##SectionName##", "Добавлено");
                                sw.Write(content);
                                sr.Close();

                                if (uc_Added.Count > 0)
                                {
                                    AddRecordsToReport(uc_Added, sw, projectContext.Acid);
                                }
                                if (bugs_Added.Count > 0)
                                {
                                    AddRecordsToReport(bugs_Added, sw, projectContext.Acid);
                                }

                                //add status footer
                                sr = new StreamReader(ReportPath + "Section_Footer_Fixed.html");
                                content = sr.ReadToEnd();
                                sw.Write(content);
                                sr.Close();
                            }
                        }
                        catch (Exception e)
                        {

                        }

                        //add project_footer
                        sr = new StreamReader(ReportPath + "Project_Footer_Fixed.html");
                        content = sr.ReadToEnd();
                        sw.Write(content);
                        sr.Close();
                    }
                }

            }
            else
            {
                //TO-DO: add infor to the report that no projects had been modified for this day.
            }

        }

        private static void AddRecordsToReport(List<UserStory> userStories, StreamWriter sw, String acid)
        {
            foreach (UserStory story in userStories)
            {
                int EntityID = story.Id;
                String EntityName = story.Name;
                String EntityType = "UserStory";
                //String EntityDeveloperAndEffort = " " + progressInt(story.EffortCompleted, story.Effort).ToString() + "% &emsp;&emsp; Разработчик: " + (story.Assignments.Items.Count > 0 ? story.Assignments.Items[0].GeneralUser.ToString() : "Не назначен");
                //String EntityPercentCompleted = progressInt(story.EffortCompleted, story.Effort).ToString();
                String EntityEffortCompleted = story.EffortCompleted.ToString();
                String EntityEffortRemain = story.EffortToDo.ToString();
                String EntityState = story.EntityState.Name;
                String EntityDeveoperName = "";                
                String AssignedUserID = "";
                if (story.Assignments.Items.Count > 0)
                {
                    AssignedUserID = story.Assignments.Items[0].GeneralUser.Id.ToString();
                    EntityDeveoperName = story.Assignments.Items[0].GeneralUser.ToString();
                }
                else
                {
                    EntityDeveoperName = "Не назначен";
                }

                string description = "";
                if (story.Description != null && story.Description.Length > 0)
                {
                    description = HttpUtility.HtmlDecode(StripHTML(story.Description));
                }

                WriteEntityToReport(EntityID, EntityName, EntityType, EntityState, EntityDeveoperName, EntityEffortCompleted, EntityEffortRemain, progressInt(story.EffortCompleted, story.Effort), AssignedUserID, description, sw, acid);
            }

        }
        private static void AddRecordsToReport(List<Bug> bugs, StreamWriter sw, String acid)
        {
            foreach (Bug bug in bugs)
            {
                int EntityID = bug.Id;
                String EntityName = bug.Name;
                String EntityType = "Bug";
                //String EntityDeveloperAndEffort = " " + progressInt(bug.EffortCompleted, bug.Effort).ToString() + "% &emsp;&emsp; Разработчик: " + (bug.Assignments.Items.Count > 0 ? bug.Assignments.Items[0].GeneralUser.ToString() : "Не назначен");
                //String EntityPercentCompleted = progressInt(bug.EffortCompleted, bug.Effort).ToString();
                String EntityEffortCompleted = bug.EffortCompleted.ToString();
                String EntityEffortRemain = bug.EffortToDo.ToString();
                String EntityState = bug.EntityState.Name;
                String EntityDeveoperName = "";                
                String AssignedUserID = "";
                if (bug.Assignments.Items.Count > 0)
                {
                    AssignedUserID = bug.Assignments.Items[0].GeneralUser.Id.ToString();
                    EntityDeveoperName = bug.Assignments.Items[0].GeneralUser.ToString();
                }
                else
                {
                    EntityDeveoperName = "Не назначен";
                }

                string description = "";
                if (bug.Description != null && bug.Description.Length > 0)
                {
                    description = HttpUtility.HtmlDecode(StripHTML(bug.Description));
                }

                WriteEntityToReport(EntityID, EntityName, EntityType, EntityState, EntityDeveoperName, EntityEffortCompleted, EntityEffortRemain, progressInt(bug.EffortCompleted, bug.Effort), AssignedUserID, description, sw, acid);
            }

        }

        private static void WriteEntityToReport(int EntityId, String EntityName, String EntityType, String EntityState, String EntityDeveloperName, String EntityEffortCompleted, String EntityEffortRemain, int Progress, String AssignedUserID, String description, StreamWriter sw, String acid)
        {
            StreamReader sr = new StreamReader(ReportPath + "Entity_Record_Fixed.html");
            String content = sr.ReadToEnd();

            content = content.Replace("##EntityName##", EntityId + " : " + EntityName);
            content = content.Replace("##EntityType##", EntityType);
            int imageWidth = 38;
            switch (EntityType)
            {
                case "Bug":
                    imageWidth = 26;
                    break;
            }
            content = content.Replace("##ImageWidth##", imageWidth.ToString());
            //content = content.Replace("##EntityDeveloperAndEffort##", EntityDeveloperAndEffort);
            content = content.Replace("##percentCompleted##", Progress.ToString());
            content = content.Replace("##developerName##", EntityDeveloperName);
            content = content.Replace("##effortCompleted##", EntityEffortCompleted);
            content = content.Replace("##effortRemain##", EntityEffortRemain);
            content = content.Replace("##EntityState##", EntityState);
            content = content.Replace("##AssignedUserID", AssignedUserID);
            content = content.Replace("##acid##", acid);
            content = content.Replace("##entityID##", EntityId.ToString()); 
            //content = content.Replace("##ProgressWidth##", Progress);
            //#728397
            if (Progress > 0)
            {
                content = content.Replace("##progressbgcolorandwidth##", "style=\"background-color:#bac3ca;\" width=\"" + ((int)(Progress * 0.5)).ToString() + "\"");
            }
            else
            {
                content = content.Replace("##progressbgcolorandwidth##", "");
            }

            content = content.Replace("##EntityDescriptrion##", (description.Length > 0 ? description.Substring(0, (description.Length > 255 ? 255 : description.Length)) + " ....." : ""));
            sw.Write(content);
            sr.Close();

        }

        private static int progressInt(double EffortCompleted, double Effort)
        {
            return (Effort == 0 ? 0 : (int)(EffortCompleted / Effort * 100));
        }
        private static DateTime reportDate()
        {
            int daysMinus = 1;
            switch (DateTime.Today.DayOfWeek)
            {
                case DayOfWeek.Saturday:
                    daysMinus = 1;
                    break;
                case DayOfWeek.Sunday:
                    daysMinus = 2;
                    break;
                case DayOfWeek.Monday:
                    daysMinus = 3;
                    break;
                default:
                    daysMinus = 1;
                    break;
            }
            return DateTime.Today.AddDays(-daysMinus); ;
        }
        private static string StripHTML(string HTMLText)
        {
            Regex reg = new Regex("<[^>]+>", RegexOptions.IgnoreCase);
            return reg.Replace(HTMLText, "");
        }

        private static void SendEmail()
        {
            String subject = "TP Daily Recap";
            String senderAddress = "tpnotification@dentsuaegis.ru";
            String recepinetsList = "dmitry.mironov@dentsuaegis.ru";
            String serviceName = "TPDailyRecap";
            StreamReader sr = new StreamReader(@"c:\temp\report.html");
            String bodyHTML = sr.ReadToEnd();

            var mailClient = new AMService.AMServiceClient();
            mailClient.AddToMailQueueAsIs(subject, bodyHTML, senderAddress, recepinetsList, 5, AMService.PriorityEnum.Normal, null, serviceName);
            mailClient.Close();
            sr.Close();
            sr.Dispose();
        }

    }
}

