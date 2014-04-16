using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace TPDailyRecapLight
{
    class Program
    {
        private const string PathToTp = "https://targetprocess.aemedia.ru/TargetProcess2/api/v1/";
        private const string ReportPath = @"C:\Users\dmitry.mironov\Documents\Visual Studio 2013\Projects\dailyrecap\dailyrecap\";


        static void Main(string[] args)
        {
            //create web client to get response from webserver
            var client = new WebClient();
            client.UseDefaultCredentials = true;
            client.Encoding = Encoding.UTF8;

            //create Streamwriter to write html to report file
            //open report top part
            StreamWriter reportFile = new StreamWriter(@"c:\temp\report.html", false, Encoding.UTF8);            
            StreamReader sr = new StreamReader(ReportPath + "Report_Header.html");
            //update report date in the output report file
            string content = sr.ReadToEnd();
            content = content.Replace("##ReportDate##", reportDate().ToString("dddd dd MMMM", CultureInfo.CreateSpecificCulture("ru-RU")));
            reportFile.Write(content);

            StartBuidingReport(client, reportDate(), reportFile);

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
                    StreamReader sr = new StreamReader(ReportPath + "Project_top.html");
                    String content = sr.ReadToEnd();


                    UserStoriesCollection userStoriesCollection = JsonConvert.DeserializeObject<UserStoriesCollection>(client.DownloadString(PathToTp + "UserStories?include=[Id,Name,Description,StartDate,EndDate,CreateDate,ModifyDate,Effort,EffortCompleted,EffortToDo,Owner[id,FirstName,LastName],EntityState[id,name],Feature[id,name],Assignments[Role,GeneralUser[id,FirstName,LastName]]]&where=(Project.Id eq " + project.Id + ") and (ModifyDate eq '" + reportDate.ToString("yyyy-MM-dd") + "')&take=1000&format=json"));    
                    //check that we got at least one user story
                    if (userStoriesCollection.Items.Count > 0)
                    { 
                        //Add report section to output file
                        
                        content = content.Replace("##ProjectName##", project.Name);
                        content = content.Replace("##ProjectOwner##", project.Owner.ToString());
                        sw.Write(content);
                        sr.Close();

                        //process project and userstories
                        //get stories lists by status                        
                        try
                        { 
                            List<UserStory> uc_Completed = userStoriesCollection.Items.Where(us => us.EndDate.Value.Date.Equals(reportDate.Date)).ToList<UserStory>();
                            if (uc_Completed.Count > 0)
                            {
                                //add Section to report
                                sr = new StreamReader(ReportPath + "status_header.html");
                                content = sr.ReadToEnd();
                                content = content.Replace("##Status##", "Выполнено");
                                sw.Write(content);
                                sr.Close();

                                AddRecordsToReport(uc_Completed, project, sw);

                                //add status footer
                                sr = new StreamReader(ReportPath + "status_footer.html");
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
                            List<UserStory> uc_Modified = userStoriesCollection.Items.Where(us => us.ModifyDate.Value.Date.Equals(reportDate.Date) && us.EntityState.Name != "Done" && us.CreateDate.Value.Date.CompareTo(reportDate.Date)!=0).ToList<UserStory>();
                            if (uc_Modified.Count > 0)
                            {
                                //add Section to report
                                sr = new StreamReader(ReportPath + "status_header.html");
                                content = sr.ReadToEnd();
                                content = content.Replace("##Status##", "Изменено");
                                sw.Write(content);
                                sr.Close();

                                AddRecordsToReport(uc_Modified, project, sw);

                                //add status footer
                                sr = new StreamReader(ReportPath + "status_footer.html");
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
                                sr = new StreamReader(ReportPath + "status_header.html");
                                content = sr.ReadToEnd();
                                content = content.Replace("##Status##", "Добавлено");
                                sw.Write(content);
                                sr.Close();

                                AddRecordsToReport(uc_Added, project, sw);

                                //add status footer
                                sr = new StreamReader(ReportPath + "status_footer.html");
                                content = sr.ReadToEnd();
                                sw.Write(content);
                                sr.Close();
                            }
                        }
                        catch (Exception e)
                        {

                        }
                    }

                    //add report_footer
                    sr = new StreamReader(ReportPath + "project_footer.html");
                    content = sr.ReadToEnd();
                    sw.Write(content);
                    sr.Close();
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
                StreamReader sr = new StreamReader(ReportPath + "UserStory.html");
                String content = sr.ReadToEnd();

                content = content.Replace("##UserStoryName##", story.Id + " : " + story.Name);
               
                content = content.Replace("##UserStoryDeveloperAndEffort##", "(Прогресс: " + progressInt(story.EffortCompleted, story.Effort).ToString() + "%    Разработчик: " + story.Assignments.Items[0].GeneralUser.ToString() + ")");
                content = content.Replace("##UserStoryDescription##", story.Description);
                sw.Write(content);
                sr.Close();
            }
            
        }

        private static int progressInt(double EffortCompleted, double Effort)
        {            
            return (int)(EffortCompleted/Effort*100);
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
    }
}
