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
        //private const string ReportPath = @"C:\Users\dmitry.mironov\Documents\Visual Studio 2013\Projects\DailyRecapLight\";
        private const string ReportPath = @"C:\Users\dmitry.mironov\Documents\Visual Studio 2013\Projects\TP_Daily_Recap_v2\TPDRv2\";
        private const bool DebugModeOn = true;

        static void Main(string[] args)
        {
            //create web client to get response from webserver
            var client = new WebClient();
            client.UseDefaultCredentials = true;
            client.Encoding = Encoding.UTF8;

            //create Streamwriter to write html to report file
            //open report top part
            StreamWriter reportFile = new StreamWriter(@"c:\temp\report2.html", false, Encoding.UTF8);

            //add report header
            StreamReader sr = new StreamReader(ReportPath + "Report_Header.html");

            //update report date in the output report file
            string content = sr.ReadToEnd();
            content = content.Replace("##ReportStartDate##", reportDate().ToString("dddd dd MMMM", CultureInfo.CreateSpecificCulture("ru-RU")));
            content = content.Replace("##ReportEndDate##", reportDate().ToString("dddd dd MMMM", CultureInfo.CreateSpecificCulture("ru-RU")));
            reportFile.Write(content);

            StartBuidingReport(client, reportDate(), reportDate(), reportFile);

            //add report footer
            sr = new StreamReader(ReportPath + "Report_Footer.html");
            content = sr.ReadToEnd();
            reportFile.Write(content);

            //save report file and dispose
            reportFile.Close();
            reportFile.Dispose();

            //SendEmail();
        }

        private static void StartBuidingReport(WebClient client, DateTime reportStartDate, DateTime reportEndDate, StreamWriter sw)
        {
            bool projectAdded2Report = false;
            //get collection of projects
            ProjectsCollection projectsCollection = JsonConvert.DeserializeObject<ProjectsCollection>(client.DownloadString(PathToTp + "projects?include=[name,id,owner]&where=(IsActive eq 'true') and (id ne 1710)&take=1000&format=json")); //excelude Dmitry Mironov's personal tasks

            //project acid property
            //String acid = "";

            ShowInConsole("Total projects: " + projectsCollection.Items.Count());
            
            //check that projec collection has at least one project
            if (projectsCollection.Items.Count > 0)
            {
                //for each project check if there are any  userstories or bugs 
                foreach (Project project in projectsCollection.Items)
                {
                    //get project acid
                    ShowInConsole("Getting project ACID.");
                    ProjectContext projectContext = JsonConvert.DeserializeObject<ProjectContext>(client.DownloadString(PathToTp + "Context/?ids=" + project.Id + "&format=json"));

                    ShowInConsole("Start processing project : " + project.Id + " : " + project.Name);                    

                    //preparing all collections we need for the report

                    ShowInConsole("Selecting US and Bugs that had been completed during report period.");                    
                    UserStoriesCollection userStoriesCollection_Done = JsonConvert.DeserializeObject<UserStoriesCollection>(client.DownloadString(PathToTp + "UserStories?include=[Id,Name,Description,StartDate,EndDate,CreateDate,ModifyDate,Effort,EffortCompleted,EffortToDo,Owner[id,FirstName,LastName],EntityState[id,name],Feature[id,name],Assignments[Role,GeneralUser[id,FirstName,LastName]]]&where=(Project.Id eq " + project.Id + ") and (ModifyDate gte '" + reportStartDate.ToString("yyyy-MM-dd") + "') and (ModifyDate lte '" + reportEndDate.ToString("yyyy-MM-dd") + "') and (EntityState.Name eq 'Done')&take=1000&format=json"));
                    BugsCollection bugsCollection_Done = JsonConvert.DeserializeObject<BugsCollection>(client.DownloadString(PathToTp + "Bugs?include=[Id,Name,Description,StartDate,EndDate,CreateDate,ModifyDate,Effort,EffortCompleted,EffortToDo,Owner[id,FirstName,LastName],EntityState[id,name],Assignments[Role,GeneralUser[id,FirstName,LastName]]]&where=(Project.Id eq " + project.Id + ") and (ModifyDate gte '" + reportStartDate.ToString("yyyy-MM-dd") + "') and (ModifyDate lte '" + reportEndDate.ToString("yyyy-MM-dd") + "') and (EntityState.Name eq 'Done')&take=1000&format=json"));
                    UserStoriesCollection userStoriesCollection_R2R = JsonConvert.DeserializeObject<UserStoriesCollection>(client.DownloadString(PathToTp + "UserStories?include=[Id,Name,Description,StartDate,EndDate,CreateDate,ModifyDate,Effort,EffortCompleted,EffortToDo,Owner[id,FirstName,LastName],EntityState[id,name],Feature[id,name],Assignments[Role,GeneralUser[id,FirstName,LastName]]]&where=(Project.Id eq " + project.Id + ") and (EntityState.Name eq 'Ready to release')&take=1000&format=json"));
                    BugsCollection bugsCollection_R2R = JsonConvert.DeserializeObject<BugsCollection>(client.DownloadString(PathToTp + "Bugs?include=[Id,Name,Description,StartDate,EndDate,CreateDate,ModifyDate,Effort,EffortCompleted,EffortToDo,Owner[id,FirstName,LastName],EntityState[id,name],Assignments[Role,GeneralUser[id,FirstName,LastName]]]&where=(Project.Id eq " + project.Id + ") and (EntityState.Name eq 'Ready to release')&take=1000&format=json"));
                    UserStoriesCollection userStoriesCollection_InProgress = JsonConvert.DeserializeObject<UserStoriesCollection>(client.DownloadString(PathToTp + "UserStories?include=[Id,Name,Description,StartDate,EndDate,CreateDate,ModifyDate,Effort,EffortCompleted,EffortToDo,Owner[id,FirstName,LastName],EntityState[id,name],Feature[id,name],Assignments[Role,GeneralUser[id,FirstName,LastName]]]&where=(Project.Id eq " + project.Id + ") and (EntityState.name ne 'Open') and (EntityState.name ne 'Planned') and (EntityState.name ne 'Ready to release') and (EntityState.name ne 'Done')&take=1000&format=json"));
                    BugsCollection bugsCollection_InProgress = JsonConvert.DeserializeObject<BugsCollection>(client.DownloadString(PathToTp + "Bugs?include=[Id,Name,Description,StartDate,EndDate,CreateDate,ModifyDate,Effort,EffortCompleted,EffortToDo,Owner[id,FirstName,LastName],EntityState[id,name],Assignments[Role,GeneralUser[id,FirstName,LastName]]]&where=(Project.Id eq " + project.Id + ") and (EntityState.name ne 'Open') and (EntityState.name ne 'Planned') and (EntityState.name ne 'Ready to release') and (EntityState.name ne 'Done')&take=1000&format=json"));
                    
                    ShowInConsole("Checking that we have at lease one entity to show in the report:");
                    //check that we got at least one user story or bug that had been completed or ready to release
                    if (userStoriesCollection_Done.Items.Count > 0 || bugsCollection_Done.Items.Count > 0 || userStoriesCollection_R2R.Items.Count > 0 || bugsCollection_R2R.Items.Count > 0 || userStoriesCollection_InProgress.Items.Count > 0 || bugsCollection_InProgress.Items.Count > 0)
                    {
                        ShowInConsole("At least one entity should be added to the report.");

                        AddProjectHeaderToReport(project, sw);
                        projectAdded2Report = true;

                        if (userStoriesCollection_Done.Items.Count > 0 || bugsCollection_Done.Items.Count > 0 || userStoriesCollection_R2R.Items.Count > 0 || bugsCollection_R2R.Items.Count > 0)
                        {
                            AddSectionHeaderToReport("Выполнено", sw);

                            //adding completed entities                       
                            if (userStoriesCollection_Done.Items.Count > 0)
                            {
                                AddRecordsToReport(userStoriesCollection_Done.Items.ToList<UserStory>(), sw, projectContext.Acid);
                            }

                            if (bugsCollection_Done.Items.Count > 0)
                            {
                                AddRecordsToReport(bugsCollection_Done.Items.ToList<Bug>(), sw, projectContext.Acid);
                            }

                            if (userStoriesCollection_R2R.Items.Count > 0)
                            {
                                AddRecordsToReport(userStoriesCollection_R2R.Items.ToList<UserStory>(), sw, projectContext.Acid);
                            }

                            if (bugsCollection_R2R.Items.Count > 0)
                            {
                                AddRecordsToReport(bugsCollection_R2R.Items.ToList<Bug>(), sw, projectContext.Acid);
                            }

                            AddSectionFooterToReport(sw);
                        }
                    }
                                                            
                    //check that we have at least one entity to add to the report
                    if (userStoriesCollection_InProgress.Items.Count > 0 || bugsCollection_InProgress.Items.Count > 0)
                    {
                        AddSectionHeaderToReport("В процессе", sw);
                        
                        if (userStoriesCollection_InProgress.Items.Count > 0)
                        { 
                            AddRecordsToReport(userStoriesCollection_InProgress.Items.ToList<UserStory>(), sw, projectContext.Acid); 
                        }

                        if (bugsCollection_InProgress.Items.Count > 0)
                        {
                            AddRecordsToReport(bugsCollection_InProgress.Items.ToList<Bug>(), sw, projectContext.Acid); 
                        }

                        AddSectionFooterToReport(sw);
                    }

                    if (projectAdded2Report)
                    { 
                        AddProjectFooterToReport(sw);
                        projectAdded2Report = false;
                    }
                    
                }            

            }
            else
            {
                //TO-DO: add infor to the report that no projects had been modified for this day.
            }
    }

        private static void AddProjectHeaderToReport(Project project, StreamWriter sw)
        {
            //Add report section to output file    

            StreamReader sr = new StreamReader(ReportPath + "Project_Header.html");
            String content = sr.ReadToEnd();

            content = content.Replace("##ProjectName##", project.Name);
            //add project owner
            content = content.Replace("##ProjectOwner##", project.Owner.ToString());
            sw.Write(content);
            sr.Close();
        }
        private static void AddProjectFooterToReport(StreamWriter sw)
        {
            //add project_footer
            StreamReader sr = new StreamReader(ReportPath + "Project_Footer.html");
            String content = sr.ReadToEnd();
            sw.Write(content);
            sr.Close();  
        }
        private static void AddSectionHeaderToReport(String SectionName, StreamWriter sw)
        {
            StreamReader sr = new StreamReader(ReportPath + "Section_Header.html");
            String content = sr.ReadToEnd();
            content = content.Replace("##SectionName##", SectionName);
            sw.Write(content);
            sr.Close();

            sr = new StreamReader(ReportPath + "Entities_Header.html");
            content = sr.ReadToEnd();
            sw.Write(content);
            sr.Close();
        }
        private static void AddSectionFooterToReport(StreamWriter sw)
        {
            StreamReader sr = new StreamReader(ReportPath + "Entities_Footer.html");
            String content = sr.ReadToEnd();
            sw.Write(content);
            sr.Close();
            //add status footer
            sr = new StreamReader(ReportPath + "Section_Footer.html");
            content = sr.ReadToEnd();
            sw.Write(content);
            sr.Close();
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
            StreamReader sr = new StreamReader(ReportPath + "Entity_Record.html");
            String content = sr.ReadToEnd();

            content = content.Replace("##EntityID##", EntityId.ToString());
            content = content.Replace("##EntityName##", EntityName);
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
            StreamReader sr = new StreamReader(@"c:\temp\report2.html");
            String bodyHTML = sr.ReadToEnd();

            var mailClient = new AMService.AMServiceClient();
            mailClient.AddToMailQueueAsIs(subject, bodyHTML, senderAddress, recepinetsList, 5, AMService.PriorityEnum.Normal, null, serviceName);
            mailClient.Close();
            sr.Close();
            sr.Dispose();
        }

        private static void ShowInConsole(String s)
        {
            if (DebugModeOn)
            {
                Console.WriteLine(s);
            }
        }
    }
}

