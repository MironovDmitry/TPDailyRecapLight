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
            StreamReader sr = new StreamReader(ReportPath + "Report_Header.html");
            
            //update report date in the output report file
            string content = sr.ReadToEnd();
            content = content.Replace("##ReportDate##", reportDate().ToString("dddd dd MMMM", CultureInfo.CreateSpecificCulture("ru-RU")));
            reportFile.Write(content);

            StartBuidingReport(client, reportDate(), reportFile);

            //add report footer
            sr = new StreamReader(ReportPath + "Report_Footer.html");
            content = sr.ReadToEnd();
            reportFile.Write(content);

            //save report file and dispose
            reportFile.Close();
            reportFile.Dispose();
        }

        private static void StartBuidingReport(WebClient client, DateTime reportDate, StreamWriter sw)
        {
            //get collection of projects
            ProjectsCollection projectsCollection = JsonConvert.DeserializeObject<ProjectsCollection>(client.DownloadString(PathToTp + "projects?include=[name,id,owner]&where=IsActive%20eq%20%27true%27&take=1000&format=json"));

            Console.WriteLine("Total projects: " + projectsCollection.Items.Count());
            //check that projec collection has at least one project
            if (projectsCollection.Items.Count > 0)
            {
                //for each project check if there are any  userstories or bugs modified on this report date
                foreach(Project project in projectsCollection.Items)
                {
                    //StreamReader sr = new StreamReader(ReportPath + "Project_top.html");
                    StreamReader sr = new StreamReader(ReportPath + "Project_Header.html");
                    String content = sr.ReadToEnd();


                    UserStoriesCollection userStoriesCollection = JsonConvert.DeserializeObject<UserStoriesCollection>(client.DownloadString(PathToTp + "UserStories?include=[Id,Name,Description,StartDate,EndDate,CreateDate,ModifyDate,Effort,EffortCompleted,EffortToDo,Owner[id,FirstName,LastName],EntityState[id,name],Feature[id,name],Assignments[Role,GeneralUser[id,FirstName,LastName]]]&where=(Project.Id eq " + project.Id + ") and (ModifyDate eq '" + reportDate.ToString("yyyy-MM-dd") + "')&take=1000&format=json"));    
                    //check that we got at least one user story
                    if (userStoriesCollection.Items.Count > 0)
                    { 
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
                            List<UserStory> uc_Completed = userStoriesCollection.Items.Where(us => (us.EndDate.HasValue ? us.EndDate.Value.Date.Equals(reportDate.Date):false)).ToList<UserStory>();
                            if (uc_Completed.Count > 0)
                            {
                                //add Section to report
                                //sr = new StreamReader(ReportPath + "Section_Header.html");
                                sr = new StreamReader(ReportPath + "Section_Header.html");
                                content = sr.ReadToEnd();
                                content = content.Replace("##SectionName##", "Выполнено");
                                sw.Write(content);
                                sr.Close();

                                AddRecordsToReport(uc_Completed, project, sw);

                                //add status footer
                                sr = new StreamReader(ReportPath + "Section_Footer.html");
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
                            List<UserStory> uc_Modified = userStoriesCollection.Items.Where(us => (us.ModifyDate.HasValue ? us.ModifyDate.Value.Date.Equals(reportDate.Date):false) && us.EntityState.Name != "Done" && us.CreateDate.Value.Date.CompareTo(reportDate.Date)!=0).ToList<UserStory>();
                            if (uc_Modified.Count > 0)
                            {
                                //add Section to report
                                sr = new StreamReader(ReportPath + "Section_Header.html");
                                content = sr.ReadToEnd();
                                content = content.Replace("##SectionName##", "Изменено");
                                sw.Write(content);
                                sr.Close();

                                AddRecordsToReport(uc_Modified, project, sw);

                                //add status footer
                                sr = new StreamReader(ReportPath + "Section_Footer.html");
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
                            List<UserStory> uc_Added = userStoriesCollection.Items.Where(us => us.CreateDate.Value.Date.Equals(reportDate.Date)).ToList<UserStory>();
                            if (uc_Added.Count > 0)
                            {
                                //add Section to report
                                sr = new StreamReader(ReportPath + "Section_Header.html");
                                content = sr.ReadToEnd();
                                content = content.Replace("##SectionName##", "Добавлено");
                                sw.Write(content);
                                sr.Close();

                                AddRecordsToReport(uc_Added, project, sw);

                                //add status footer
                                sr = new StreamReader(ReportPath + "Section_Footer.html");
                                content = sr.ReadToEnd();
                                sw.Write(content);
                                sr.Close();
                            }
                        }
                        catch (Exception e)
                        {

                        }

                        //add project_footer
                        sr = new StreamReader(ReportPath + "project_footer.html");
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

        private static void AddRecordsToReport(List<UserStory> userStories, Project project, StreamWriter sw)
        {
            foreach (UserStory story in userStories)
            {
                //StreamReader sr = new StreamReader(ReportPath + "UserStory.html");
                StreamReader sr = new StreamReader(ReportPath + "EntityRecord.html");
                String content = sr.ReadToEnd();

                content = content.Replace("##EntityName##", story.Id + " : " + story.Name);
                content = content.Replace("##EntityType##", "UserStory");
                content = content.Replace("##EntityDeveloperAndEffort##", "Прогресс: " + progressInt(story.EffortCompleted, story.Effort).ToString() + "% &emsp;&emsp; Разработчик: " + (story.Assignments.Items.Count > 0 ? story.Assignments.Items[0].GeneralUser.ToString() : "Не назначен"));

                string description = "";
                if (story.Description != null && story.Description.Length > 0)
                {
                     description = HttpUtility.HtmlDecode(StripHTML(story.Description));
                }
                
                //content = content.Replace("##UserStoryDescription##", (story.Description.Length>0? story.Description.Substring(0,(story.Description.Length> 255?255:story.Description.Length)) + " .....":""));
                content = content.Replace("##EntityDescriptrion##", (description.Length > 0 ? description.Substring(0, (description.Length > 255 ? 255 : description.Length)) + " ....." : ""));
                sw.Write(content);
                sr.Close();
            }
            
        }

       
        private static int progressInt(double EffortCompleted, double Effort)
        {
            return (Effort == 0 ?0:(int)(EffortCompleted/Effort*100));
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
            return DateTime.Today.AddDays(-daysMinus);;
        }
        private static string StripHTML(string HTMLText)
        {
            Regex reg = new Regex("<[^>]+>", RegexOptions.IgnoreCase);
            return reg.Replace(HTMLText, "");
        }
    }
}

