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
        //private const string ReportPath = @"C:\Users\dmiron01.EMEA-MEDIA\Documents\Visual Studio 2013\Projects\TP_Daily_Recap_v2\TPDRv2\";        
        //private const string ReportPath = Properties.Settings.Default.HTMLLocation;
        private static String ReportPath = Properties.Settings.Default.HTMLLocation.ToString();
        private static String sHTMLReportFile = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\report.html";
        private static String sHTMLSummaryFile = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Summary.html";
        private const bool DebugModeOn = false;
        
        private const String UserStorySelectionFields = "include=[Id,Name,Description,StartDate,EndDate,CreateDate,ModifyDate,Effort,EffortCompleted,EffortToDo,Owner[id,FirstName,LastName],EntityState[id,name],Feature[id,name],Assignments[Role,GeneralUser[id,FirstName,LastName]],TeamIteration[Id,Name,StartDate,EndDate]]";
        private const String BugSelectionFields = "include=[Id,Name,Description,StartDate,EndDate,CreateDate,ModifyDate,Effort,EffortCompleted,EffortToDo,Owner[id,FirstName,LastName],EntityState[id,name],Assignments[Role,GeneralUser[id,FirstName,LastName]],TeamIteration[Id,Name,StartDate,EndDate]]";
        private const String ProjectSelectionFields = "include=[name,id,owner,process]";
        private const string token = "&token=Mjo5OUJFOUVCQTNEMzNGOEI3MzcyRjg1MzEwRkNGNzZDRQ==";

        private static DateTime reportStartDate;
        private static DateTime reportEndDate;

        private static String reportType = "Daily";

        static void Main(string[] args)
        {              
            //process passed arguments
            //String rStartDate = DateTime.Today.ToString();
            //String rEndDate = DateTime.Today.ToString();
            // String reportType = "Daily";

            if (args.Length == 0)
            {

            }
            else
            {
                switch (args[0])
                { 
                    case "Daily":
                    Int16 daysMinus = 1;

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
                        reportStartDate = DateTime.Today.AddDays(-daysMinus).Date;
                        reportEndDate = DateTime.Today.AddDays(-daysMinus).Date;
                        reportType = "Daily";
                        break;
                    case "Weekly":
                        reportStartDate = DateTime.Today.AddDays(-7).Date;
                        reportEndDate = DateTime.Today.AddDays(-1).Date;                        
                        reportType = "Weekly";
                        break;                    
                }                
            }
            //create web client to get response from webserver
            var client = new WebClient();
            //client.UseDefaultCredentials = false;            
            client.Encoding = Encoding.UTF8;

            //create Streamwriter to write html to report file
            //open report top part
            StreamWriter reportFile = new StreamWriter(sHTMLReportFile, false, Encoding.UTF8);
            
            //add report header
            StreamReader sr = new StreamReader(ReportPath + "Report_Header.html");

            //update report date in the output report file
            string content = sr.ReadToEnd();
            content = content.Replace("##ReportStartDate##", reportStartDate.ToString("dddd dd MMMM", CultureInfo.CreateSpecificCulture("ru-RU")));
            content = content.Replace("##ReportEndDate##", reportEndDate.ToString("dddd dd MMMM", CultureInfo.CreateSpecificCulture("ru-RU")));
            reportFile.Write(content);

            StartBuidingReport(client, reportStartDate, reportEndDate, reportFile, reportType);

            //add report footer
            sr = new StreamReader(ReportPath + "Report_Footer.html");
            content = sr.ReadToEnd();
            reportFile.Write(content);

            //save report file and dispose
            reportFile.Close();
            reportFile.Dispose();

            //add summary to report
            StreamReader srSummary = new StreamReader(sHTMLSummaryFile);
            content = srSummary.ReadToEnd();
            StreamReader srReport = new StreamReader(sHTMLReportFile);
            String reportContent = srReport.ReadToEnd();            
            //reportContent = reportContent.Replace("##SummarySection##", reportType=="Weekly" ? content : "");
            reportContent = reportContent.Replace("##SummarySection##", ""); //убрал саммари по требованию Жени.
            reportContent = reportContent.Replace("##SummaryEntryRight##", "&nbsp;");
            srReport.Close();
            srReport.Dispose();
            srSummary.Close();
            srSummary.Dispose();
            reportFile = new StreamWriter(sHTMLReportFile, false, Encoding.UTF8);
            reportFile.Write(reportContent);
            reportFile.Close();
            reportFile.Dispose();


            SendEmail(reportStartDate, reportEndDate, reportType);
        }

        private static void StartBuidingReport(WebClient client, DateTime reportStartDate, DateTime reportEndDate, StreamWriter sw, String reportType)
        {
            //temp file to buid summary section
            StreamWriter summaryFile = new StreamWriter(sHTMLSummaryFile, false, Encoding.UTF8);

            StreamReader sr = new StreamReader(ReportPath + "Summary_Header.html");
            String content = sr.ReadToEnd();
            summaryFile.Write(content);
            summaryFile.Close();

            bool projectAdded2Report = false;
            //get collection of projects
            ProjectsCollection projectsCollection = JsonConvert.DeserializeObject<ProjectsCollection>(client.DownloadString(PathToTp + "projects?" + ProjectSelectionFields + "&where=(IsActive eq 'true') and (id ne 1710) and (id ne 497)&take=1000&format=json" + token)); //excelude Dmitry Mironov's personal tasks

            //project acid property
            //String acid = "";

            ShowInConsole("Total projects: " + projectsCollection.Items.Count());
            
            //check that projec collection has at least one project
            if (projectsCollection.Items.Count > 0)
            {
                bool AddProjectToSummaryLeft = true;
                //for each project check if there are any  userstories or bugs 
                foreach (Project project in projectsCollection.Items)
                {
                    //get project acid
                    ShowInConsole("Getting project ACID.");
                    ProjectContext projectContext = JsonConvert.DeserializeObject<ProjectContext>(client.DownloadString(PathToTp + "Context/?ids=" + project.Id + "&format=json" + token));
                    //TeamIterationsCollection tic = JsonConvert.DeserializeObject<TeamIterationsCollection>(client.DownloadString(PathToTp + "teamiterations?where=(IsCurrent eq 'true')&format=json" + token));
                    TeamIteration teamIteration = JsonConvert.DeserializeObject<TeamIterationsCollection>(client.DownloadString(PathToTp + "teamiterations?where=(IsCurrent eq 'true')&format=json" + token)).Items[0];

                    ShowInConsole("Start processing project : " + project.Id + " : " + project.Name);                    

                    //preparing all collections we need for the report

                    ShowInConsole("Start processing UserStories and Bugs for the project.");
                    string queryCondition = "";

                    #region Collections : Completed
                    //Completed
                    UserStoriesCollection userStoriesCollection_Done = JsonConvert.DeserializeObject<UserStoriesCollection>(client.DownloadString(PathToTp + "UserStories?" + UserStorySelectionFields + "&where=(Project.Id eq " + project.Id + ") and (ModifyDate gte '" + reportStartDate.ToString("yyyy-MM-dd") + "') and (ModifyDate lte '" + reportEndDate.ToString("yyyy-MM-dd") + "') and (EntityState.Name eq 'Done')&take=1000&format=json" + token));
                    BugsCollection bugsCollection_Done = JsonConvert.DeserializeObject<BugsCollection>(client.DownloadString(PathToTp + "Bugs?" + BugSelectionFields + "&where=(Project.Id eq " + project.Id + ") and (ModifyDate gte '" + reportStartDate.ToString("yyyy-MM-dd") + "') and (ModifyDate lte '" + reportEndDate.ToString("yyyy-MM-dd") + "') and (EntityState.Name eq 'Done')&take=1000&format=json" + token));
                    UserStoriesCollection userStoriesCollection_R2R = JsonConvert.DeserializeObject<UserStoriesCollection>(client.DownloadString(PathToTp + "UserStories?" + UserStorySelectionFields + "&where=(Project.Id eq " + project.Id + ") and (EntityState.Name eq 'Ready to release') and (ModifyDate gte '" + reportStartDate.ToString("yyyy-MM-dd") + "') and (ModifyDate lte '" + reportEndDate.ToString("yyyy-MM-dd") + "')&take=1000&format=json" + token));
                    BugsCollection bugsCollection_R2R = JsonConvert.DeserializeObject<BugsCollection>(client.DownloadString(PathToTp + "Bugs?" + BugSelectionFields + "&where=(Project.Id eq " + project.Id + ") and (EntityState.Name eq 'Ready to release') and (ModifyDate gte '" + reportStartDate.ToString("yyyy-MM-dd") + "') and (ModifyDate lte '" + reportEndDate.ToString("yyyy-MM-dd") + "')&take=1000&format=json" + token)); 
                    #endregion

                    #region Collections : Work In Progress (Weekly) / Modified (Daily)
                    
                    switch (reportType)
                    { 
                        case "Daily":
                            //queryCondition = "(Project.Id eq " + project.Id + ") and (EntityState.name ne 'Open') and (EntityState.name ne 'Planned') and (EntityState.name ne 'User Acceptance Testing') and (EntityState.name ne 'Ready to release') and (EntityState.name ne 'Done') and (ModifyDate gte '" + reportStartDate.ToString("yyyy-MM-dd") + "') and (ModifyDate lte '" + reportEndDate.ToString("yyyy-MM-dd") + "') and (CreateDate ne '" + reportStartDate.ToString("yyyy-MM-dd") + "')";
                            queryCondition = "(Project.Id eq " + project.Id + ") and (EntityState.name ne 'Open') and (EntityState.name ne 'Planned') and (EntityState.name ne 'User Acceptance Testing') and (EntityState.name ne 'Ready to release') and (EntityState.name ne 'Done') and (CreateDate ne '" + reportStartDate.ToString("yyyy-MM-dd") + "')";
                            break;
                        case "Weekly":
                            if (project.Process.Name.Equals("Scrum"))
                            {
                                queryCondition = "(Project.Id eq " + project.Id + ") and (EntityState.name ne 'Open') and (EntityState.name ne 'Planned') and (EntityState.name ne 'User Acceptance Testing') and (EntityState.name ne 'Ready to release') and (EntityState.name ne 'Done') and (TeamIteration.ID eq " + teamIteration.Id + ")";
                            }
                            else
                            {
                                queryCondition = "(Project.Id eq " + project.Id + ") and (EntityState.name ne 'Open') and (EntityState.name ne 'Planned') and (EntityState.name ne 'User Acceptance Testing') and (EntityState.name ne 'Ready to release') and (EntityState.name ne 'Done')";
                            }
                            
                            break;
                        //case "Custom":
                        //    queryCondition = "(Project.Id eq " + project.Id + ") and (EntityState.name ne 'Open') and (EntityState.name ne 'Planned') and (EntityState.name ne 'User Acceptance Testing') and (EntityState.name ne 'Ready to release') and (EntityState.name ne 'Done') and (ModifyDate gte '" + reportStartDate.ToString("yyyy-MM-dd") + "') and (ModifyDate lte '" + reportEndDate.ToString("yyyy-MM-dd") + "')";
                        //    break;
                    }

                    BugsCollection bugsCollection_InProgress = JsonConvert.DeserializeObject<BugsCollection>(client.DownloadString(PathToTp + "Bugs?" + BugSelectionFields + "&where=" + queryCondition + "&take=1000&format=json" + token));

                    UserStoriesCollection userStoriesCollection_InProgress = JsonConvert.DeserializeObject<UserStoriesCollection>(client.DownloadString(PathToTp + "UserStories?" + UserStorySelectionFields + "&where=" + queryCondition + "&take=1000&format=json" + token));
                    List<UserStory> us_WIP = new List<UserStory>();
                    if (userStoriesCollection_InProgress.Items.Count > 0)
                    {
                        //отбираем только измененные за отчетный период стори либо стори, у ктороых были изменены таски.
                        foreach (UserStory story in userStoriesCollection_InProgress.Items.ToList<UserStory>())
                        {
                            if (story.ModifyDate.Date >= reportStartDate.Date && story.ModifyDate.Date <= reportEndDate.Date)
                            {
                                us_WIP.Add(story);
                            }
                            else
                            {
                                //сама стори не менялась, проверяем ее таски на изменения
                                //string s = "tasks?where=(UserStory.id eq " + story.Id + ") and (ModifyDate gte '" + reportStartDate.ToString("yyyy-MM-dd") + "') and (ModifyDate lte '" + reportEndDate.ToString("yyyy-MM-dd") + "')&take=1000&format=json" + token;
                                TasksCollection tc = JsonConvert.DeserializeObject<TasksCollection>(client.DownloadString(PathToTp + "tasks?where=(UserStory.id eq " + story.Id + ") and (ModifyDate gte '" + reportStartDate.ToString("yyyy-MM-dd") + "') and (ModifyDate lte '" + reportEndDate.ToString("yyyy-MM-dd") + "')&take=1000&format=json" + token));
                                if (tc.Items.Count > 0)
                                {
                                    //есть таски измененные в отчетный период, добавляем стори в коллекцию                             
                                    us_WIP.Add(story);
                                }
                            }
                        }   
                    }
                    
                    #endregion "Work In Progress"

                    #region Collections : Planned
                    //Planned

                    switch (reportType)
                    { 
                        case "Daily":
                            queryCondition = "(Project.Id eq " + project.Id + ") and (EntityState.name eq 'Planned')";                            
                            break;
                        case "Weekly":
                            if (project.Process.Name.Equals("Scrum"))
                            {
                                queryCondition = "(Project.Id eq " + project.Id + ") and (EntityState.name eq 'Open') and (TeamIteration.ID eq " + teamIteration.Id + ")";
                            }
                            else
                            {
                                //queryCondition = "(Project.Id eq " + project.Id + ") and (EntityState.name eq 'Open')";
                                //Женя попросил сделать 
                                queryCondition = "(Project.Id eq " + project.Id + ") and (EntityState.name eq 'Planned')";
                            }
                            
                            break;
                    }

                    UserStoriesCollection userStoriesCollection_Planned = JsonConvert.DeserializeObject<UserStoriesCollection>(client.DownloadString(PathToTp + "UserStories?" + UserStorySelectionFields + "&where=" + queryCondition + "&take=1000&format=json" + token));
                    BugsCollection bugsCollection_Planned = JsonConvert.DeserializeObject<BugsCollection>(client.DownloadString(PathToTp + "Bugs?" + BugSelectionFields + "&where=" + queryCondition + "&take=1000&format=json" + token));
                    //TO-DO: add selection for Scrum projects as well. 
                    #endregion

                    #region Collections : UAT
                    //User Accesptance testing
                    UserStoriesCollection userStoriesCollection_UAT = JsonConvert.DeserializeObject<UserStoriesCollection>(client.DownloadString(PathToTp + "UserStories?" + UserStorySelectionFields + "&where=(Project.Id eq " + project.Id + ") and (EntityState.name eq 'User Acceptance testing')&take=1000&format=json" + token));
                    BugsCollection bugsCollection_UAT = JsonConvert.DeserializeObject<BugsCollection>(client.DownloadString(PathToTp + "Bugs?" + BugSelectionFields + "&where=(Project.Id eq " + project.Id + ") and (EntityState.name eq 'User Acceptance testing')&take=1000&format=json" + token)); 
                    #endregion
                    
                    #region Collections : Added
                    UserStoriesCollection userStoriesCollection_Added = JsonConvert.DeserializeObject<UserStoriesCollection>(client.DownloadString(PathToTp + "UserStories?" + UserStorySelectionFields + "&where=(Project.Id eq " + project.Id + ") and (EntityState.name ne 'Done') and (EntityState.name ne 'Ready to release') and (CreateDate gte '" + reportStartDate.ToString("yyyy-MM-dd") + "') and (CreateDate lte '" + reportEndDate.ToString("yyyy-MM-dd") + "')&take=1000&format=json" + token));
                    BugsCollection bugsCollection_Added = JsonConvert.DeserializeObject<BugsCollection>(client.DownloadString(PathToTp + "Bugs?" + BugSelectionFields + "&where=(Project.Id eq " + project.Id + ") and (EntityState.name ne 'Done') and (EntityState.name ne 'Ready to release') and (CreateDate gte '" + reportStartDate.ToString("yyyy-MM-dd") + "') and (CreateDate lte '" + reportEndDate.ToString("yyyy-MM-dd") + "')&take=1000&format=json" + token));
                    #endregion


                    //*********************************************
                    //      Start generating report HTML
                    //*********************************************

                    int bugsCompleted = 0;
                    int bugsInProgress = 0;
                    int bugsPlanned = 0;
                    int bugsUAT = 0;
                    int userStoryCompleted = 0;
                    int userStoryInProgress = 0;
                    int userStoryPlanned = 0;
                    int userStoryUAT = 0;

                    ShowInConsole("Checking that we have at lease one entity to show in the report:");
                    //check that we got at least one user story or bug that need to be added to the report
                    Boolean addProjectToReport = false;

                    if (reportType == "Daily")
                    {
                        if (userStoriesCollection_Done.Items.Count > 0 || bugsCollection_Done.Items.Count > 0 || userStoriesCollection_R2R.Items.Count > 0 || bugsCollection_R2R.Items.Count > 0 || us_WIP.Count > 0 || bugsCollection_InProgress.Items.Count > 0)
                        {
                            addProjectToReport = true;
                        }
                    }
                    else
                    {
                        if (userStoriesCollection_Done.Items.Count > 0 || bugsCollection_Done.Items.Count > 0 || userStoriesCollection_R2R.Items.Count > 0 || bugsCollection_R2R.Items.Count > 0 || us_WIP.Count > 0 || bugsCollection_InProgress.Items.Count > 0 || userStoriesCollection_Planned.Items.Count > 0 || bugsCollection_Planned.Items.Count > 0)
                        {
                            addProjectToReport = true;                            
                        }
                    }

                    if (addProjectToReport)
                    {
                        ShowInConsole("At least one entity should be added to the report.");

                        AddProjectHeaderToReport(project, sw);
                        projectAdded2Report = true;


                        //Section = Выполнено
                        if (userStoriesCollection_Done.Items.Count > 0 || bugsCollection_Done.Items.Count > 0 || userStoriesCollection_R2R.Items.Count > 0 || bugsCollection_R2R.Items.Count > 0)
                        {
                            AddSectionHeaderToReport("Выполнено", sw);
                            

                            //adding completed entities                       
                            if (userStoriesCollection_Done.Items.Count > 0)
                            {
                                AddRecordsToReport(userStoriesCollection_Done.Items.ToList<UserStory>(), sw, projectContext.Acid);
                                userStoryCompleted += userStoriesCollection_Done.Items.Count;
                            }

                            if (bugsCollection_Done.Items.Count > 0)
                            {
                                AddRecordsToReport(bugsCollection_Done.Items.ToList<Bug>(), sw, projectContext.Acid);
                                bugsCompleted += bugsCollection_Done.Items.Count;
                            }

                            if (userStoriesCollection_R2R.Items.Count > 0)
                            {
                                AddRecordsToReport(userStoriesCollection_R2R.Items.ToList<UserStory>(), sw, projectContext.Acid);
                                userStoryCompleted += userStoriesCollection_Done.Items.Count;
                            }

                            if (bugsCollection_R2R.Items.Count > 0)
                            {
                                AddRecordsToReport(bugsCollection_R2R.Items.ToList<Bug>(), sw, projectContext.Acid);
                                bugsCompleted += bugsCollection_Done.Items.Count;
                            }

                            AddSectionFooterToReport(sw);
                        }

                        //Section = В работе (для еженедельного отчета)
                        //Section = Измененные (для ежедневного отчета)
                        if (us_WIP.Count > 0 || bugsCollection_InProgress.Items.Count > 0)
                        {

                            AddSectionHeaderToReport(reportType=="Daily"?"Изменено":"В работе", sw);
                            
                            if (us_WIP.Count > 0)
                            {
                                AddRecordsToReport(us_WIP, sw, projectContext.Acid);
                                userStoryInProgress += us_WIP.Count;
                            }

                            if (bugsCollection_InProgress.Items.Count > 0)
                            {
                                AddRecordsToReport(bugsCollection_InProgress.Items.ToList<Bug>(), sw, projectContext.Acid);
                                bugsInProgress += bugsCollection_InProgress.Items.Count;
                            }
                        }                           
                        
                        AddSectionFooterToReport(sw);

                        
                        //Section = Запланировано (только для еженедельного отчета)
                        if (reportType == "Weekly")
                        { 
                            if (userStoriesCollection_Planned.Items.Count > 0 || bugsCollection_Planned.Items.Count > 0)
                            {
                                AddSectionHeaderToReport("Запланировано", sw);

                                if (userStoriesCollection_Planned.Items.Count > 0)
                                {
                                    AddRecordsToReport(userStoriesCollection_Planned.Items.ToList<UserStory>(), sw, projectContext.Acid);
                                    userStoryPlanned += userStoriesCollection_Planned.Items.Count;
                                }

                                if (bugsCollection_Planned.Items.Count > 0)
                                {
                                    AddRecordsToReport(bugsCollection_Planned.Items.ToList<Bug>(), sw, projectContext.Acid);
                                    bugsPlanned += bugsCollection_Planned.Items.Count;
                                }

                                AddSectionFooterToReport(sw);
                            }
                        }

                        //Section = Добавлено (только для ежедневного отчета)
                        if (reportType == "Daily")
                        {
                            if (userStoriesCollection_Added.Items.Count > 0 || bugsCollection_Added.Items.Count > 0)
                            {
                                AddSectionHeaderToReport("Добавлено", sw);

                                if (userStoriesCollection_Added.Items.Count > 0)
                                {
                                    AddRecordsToReport(userStoriesCollection_Added.Items.ToList<UserStory>(), sw, projectContext.Acid);
                                }

                                if (bugsCollection_Added.Items.Count > 0)
                                {
                                    AddRecordsToReport(bugsCollection_Added.Items.ToList<Bug>(), sw, projectContext.Acid);
                                }

                                AddSectionFooterToReport(sw);
                            }
                        }

                        //Section = UAT (только для ежедневного отчета)
                        if (reportType == "Weekly")
                        {
                            if (userStoriesCollection_UAT.Items.Count > 0 || bugsCollection_UAT.Items.Count > 0)
                            {
                                AddSectionHeaderToReport("На тестировании у заказчика", sw);

                                if (userStoriesCollection_UAT.Items.Count > 0)
                                {
                                    AddRecordsToReport(userStoriesCollection_UAT.Items.ToList<UserStory>(), sw, projectContext.Acid);
                                    userStoryUAT += userStoriesCollection_UAT.Items.Count;
                                }

                                if (bugsCollection_UAT.Items.Count > 0)
                                {
                                    AddRecordsToReport(bugsCollection_UAT.Items.ToList<Bug>(), sw, projectContext.Acid);
                                    bugsUAT += bugsCollection_UAT.Items.Count;
                                }

                                AddSectionFooterToReport(sw);
                            }
                        }

                        //Add project footer
                        if (projectAdded2Report)
                        {
                            AddProjectFooterToReport(sw);
                            projectAdded2Report = false;
                        }

                        //Add project to Summary section
                        AddProjectToSummary(project, userStoryCompleted, userStoryInProgress, userStoryPlanned, userStoryUAT, bugsCompleted, bugsInProgress, bugsPlanned, bugsUAT,AddProjectToSummaryLeft);

                        if(AddProjectToSummaryLeft)
                        {
                            AddProjectToSummaryLeft=false;
                        }
                        else
                        {
                            AddProjectToSummaryLeft=true;
                        }
                        
                        
                    }
                    
                }            

            }
            else
            {
                //TO-DO: add infor to the report that no projects had been modified for this day.
            }

            //add summary Footer
            StreamReader srSummary = new StreamReader(ReportPath + "Summary_Footer.html");
            content = srSummary.ReadToEnd();
            
            StreamWriter swSummary = new StreamWriter(sHTMLSummaryFile, true, Encoding.UTF8);
            swSummary.Write(content);
            srSummary.Close();
            swSummary.Close();            
    }

        private static void AddProjectToSummary(Project project, int userStoryCompleted, int userStoryInProgress, int userStoryPlanned, int userStoryUAT, int bugsCompleted, int bugsInProgress, int bugsPlanned, int bugsUAT, bool Add2Left)
        {
            StreamReader srEntry = new StreamReader(ReportPath + "SummaryEntry.html");
            String entry = srEntry.ReadToEnd();

            entry = entry.Replace("##ProjectName##", project.Name);
            entry = entry.Replace("##BugsCompleted##",bugsCompleted.ToString());
            entry = entry.Replace("##UserStoryCompleted##",userStoryCompleted.ToString());
            entry = entry.Replace("##BugsInProgress##",bugsInProgress.ToString());
            entry = entry.Replace("##UserStoryInProgress##",userStoryInProgress.ToString());
            entry = entry.Replace("##BugsPlanned##",bugsPlanned.ToString());
            entry = entry.Replace("##UserStoryPlanned##",userStoryPlanned.ToString());
            entry = entry.Replace("##BugsUAT##",bugsUAT.ToString());
            entry = entry.Replace("##UserStoryUAT##",userStoryUAT.ToString());
            entry = entry.Replace("##ProjectID##", project.Id.ToString());

            if (Add2Left)
            {
                StreamReader sr = new StreamReader(ReportPath + "Summary_record.html");
                String content = sr.ReadToEnd();

                content = content.Replace("##SummaryEntryLeft##",entry);

                StreamWriter sw = new StreamWriter(sHTMLSummaryFile, true, Encoding.UTF8);
                sw.Write(content);
                sw.Close();
                sr.Close();
            }
            else
            {
                StreamReader sr = new StreamReader(sHTMLSummaryFile);
                String content = sr.ReadToEnd();

                content = content.Replace("##SummaryEntryRight##", entry);

                sr.Close();

                StreamWriter sw = new StreamWriter(sHTMLSummaryFile, false, Encoding.UTF8);
                sw.Write(content);
                sw.Close();
            }
            
        }
        
        private static void AddProjectHeaderToReport(Project project, StreamWriter sw)
        {
            //Add report section to output file    

            StreamReader sr = new StreamReader(ReportPath + "Project_Header.html");
            String content = sr.ReadToEnd();

            content = content.Replace("##ProjectName##", project.Name);
            content = content.Replace("##ProjectID##", project.Id.ToString());
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

                String EntityEffortCompleted = "";
                String EntityEffortRemain = "";
                String description = "";
                String PlannedDates = "";  

                String EntityState = story.EntityState.Name;

                if (reportType == "Weekly" && (EntityState == "Done" || EntityState == "Ready to release"))
                {
                                
                }
                else
                {
                    EntityEffortCompleted = story.EffortCompleted.ToString();
                    EntityEffortRemain = story.EffortToDo.ToString();                                     

                    description = "";
                    if (story.Description != null && story.Description.Length > 0)
                    {
                        String descriptionText = HttpUtility.HtmlDecode(StripHTML(story.Description));
                        description = @"<tr><td colspan=""2"" style=""text-wrap:normal;border-bottom:1px solid #000000;"" width=""90"">&nbsp;</td><td style=""border-bottom:1px solid #000000;"">"
                                        + (descriptionText.Length > 0 ? descriptionText.Substring(0, (descriptionText.Length > 255 ? 255 : descriptionText.Length)) + " ....." : "Описание отсутствует.") + 
                                        "</td></tr>";
                    }

                    PlannedDates = "";
                    if (story.TeamIteration != null)
                    {
                        PlannedDates = @"<tr>
                                        <td colspan=""2"" width=""90"">
                                            &nbsp;
                                        </td>
                                        <td style=""font-size:5pt;"">
                                            <span style=""font-style:italic;font-size:10pt;"">"
                                                    + story.TeamIteration.StartDate.ToShortDateString() + " : " + story.TeamIteration.EndDate.ToShortDateString() +
                                               "</span><br />&nbsp;</td></tr>";
                    }
                }

                

                WriteEntityToReport(EntityID, EntityName, EntityType, EntityState, EntityDeveoperName, EntityEffortCompleted, EntityEffortRemain, progressInt(story.EffortCompleted, story.Effort), AssignedUserID, description, sw, acid, PlannedDates);
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

                String EntityEffortCompleted = "";
                String EntityEffortRemain = "";
                string description = "";                
                String PlannedDates = "";

                String EntityState = bug.EntityState.Name;

                if (reportType == "Weekly" && (EntityState == "Done" || EntityState == "Ready to release"))
                {

                }
                else
                {
                    EntityEffortCompleted = bug.EffortCompleted.ToString();
                    EntityEffortRemain = bug.EffortToDo.ToString();

                    description = "";
                    if (bug.Description != null && bug.Description.Length > 0)
                    {
                        String descriptionText = HttpUtility.HtmlDecode(StripHTML(bug.Description));
                        description = @"<tr><td colspan=""2"" style=""text-wrap:normal;border-bottom:1px solid #000000;"" width=""90"">&nbsp;</td><td style=""border-bottom:1px solid #000000;"">"
                                        + (descriptionText.Length > 0 ? descriptionText.Substring(0, (descriptionText.Length > 255 ? 255 : descriptionText.Length)) + " ....." : "Описание отсутствует.") +
                                        "</td></tr>";
                    }

                    PlannedDates = "";
                    if (bug.TeamIteration != null)
                    {
                        PlannedDates = @"<tr><td colspan=""2"" width=""90"">&nbsp;</td><td style=""font-size:5pt;""><span style=""font-style:italic;font-size:10pt;"">" + bug.TeamIteration.StartDate.ToShortDateString() + " : " + bug.TeamIteration.EndDate.ToShortDateString() + "</span><br />&nbsp;</td></tr>";
                    }
                }                

                WriteEntityToReport(EntityID, EntityName, EntityType, EntityState, EntityDeveoperName, EntityEffortCompleted, EntityEffortRemain, progressInt(bug.EffortCompleted, bug.Effort), AssignedUserID, description, sw, acid, PlannedDates);
            }

        }

        private static void WriteEntityToReport(int EntityId, String EntityName, String EntityType, String EntityState, String EntityDeveloperName, String EntityEffortCompleted, String EntityEffortRemain, int Progress, String AssignedUserID, String description, StreamWriter sw, String acid, String PlannedDates)
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

            content = content.Replace("##PlannedDates##", PlannedDates);
            content = content.Replace("##EntityDescription##", description);
           
            //    content = content.Replace("##EntityDescription##", (description.Length > 0 ? description.Substring(0, (description.Length > 255 ? 255 : description.Length)) + " ....." : "Описание отсутствует."));
           
            sw.Write(content);
            sr.Close();

        }

        private static int progressInt(double EffortCompleted, double Effort)
        {
            return (Effort == 0 ? 0 : (int)(EffortCompleted / Effort * 100));
        }
        //private static DateTime reportDate(String sDate)
        //{
        //    int daysMinus = 1;
        //    switch (DateTime.Parse(sDate).DayOfWeek)
        //    {
        //        case DayOfWeek.Saturday:
        //            daysMinus = 1;
        //            break;
        //        case DayOfWeek.Sunday:
        //            daysMinus = 2;
        //            break;
        //        case DayOfWeek.Monday:
        //            daysMinus = 3;
        //            break;
        //        default:
        //            daysMinus = 1;
        //            break;
        //    }
        //    return DateTime.Today.AddDays(-daysMinus); ;
        //}
        private static string StripHTML(string HTMLText)
        {
            Regex reg = new Regex("<[^>]+>", RegexOptions.IgnoreCase);
            return reg.Replace(HTMLText, "");
        }

        private static void SendEmail(DateTime reportStartDate, DateTime reportEndDate, String reportType)
        {
            String subject = "TP " + reportType + " Recap : " + reportStartDate.ToString("dddd dd MMMM", CultureInfo.CreateSpecificCulture("ru-RU")) + (reportType == "Weekly"?" - " + reportEndDate.ToString("dddd dd MMMM", CultureInfo.CreateSpecificCulture("ru-RU")):"");
            String senderAddress = "tpnotification@dentsuaegis.ru";
            String recepinetsList = Properties.Settings.Default.ResepientsList;
            String serviceName = "TPRecap";
            StreamReader sr = new StreamReader(sHTMLReportFile);
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

