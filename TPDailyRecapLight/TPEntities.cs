using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TPDailyRecapLight
{    
    public class Owner
    {
        public int Id { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public override String ToString()
        {
            return FirstName + " " + LastName;
        }
    }

    public class EntityState
    {
        public int Id { get; set; }
        public string Name { get; set; }
    }
    

    public class AssignmentsCollection
    {
        public List<Assignment> Items { get; set; }
    }

    public class Role
    {
        public int Id { get; set; }
        public string Name { get; set; }
    }

    public class GeneralUser
    {
        public string Kind { get; set; }
        public int Id { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public override String ToString()
        {
            return FirstName + " " + LastName;
        }
    }

    public class Assignment
    {
        public int Id { get; set; }
        public Role Role { get; set; }
        public GeneralUser GeneralUser { get; set; }
    }
    public class UserStory
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public DateTime? StartDate { get; set; }
        public DateTime? EndDate { get; set; }
        public DateTime CreateDate { get; set; }
        public DateTime ModifyDate { get; set; }
        public object LastCommentDate { get; set; }
        public string Tags { get; set; }
        public double NumericPriority { get; set; }
        public double Effort { get; set; }
        public double EffortCompleted { get; set; }
        public double EffortToDo { get; set; }
        public double Progress { get; set; }
        public double TimeSpent { get; set; }
        public double TimeRemain { get; set; }
        public object PlannedStartDate { get; set; }
        public object PlannedEndDate { get; set; }
        public double InitialEstimate { get; set; }
        public EntityType EntityType { get; set; }
        public Owner Owner { get; set; }
        public object LastCommentedUser { get; set; }
        public Project Project { get; set; }
        public Release Release { get; set; }
        public object Iteration { get; set; }
        public object TeamIteration { get; set; }
        //public Team Team { get; set; }
        public Priority Priority { get; set; }
        public EntityState EntityState { get; set; }
        public Feature Feature { get; set; }
        public List<object> CustomFields { get; set; }
        public AssignmentsCollection Assignments { get; set; }
    }

    public class UserStoriesCollection
    {
        public string Next { get; set; }
        public List<UserStory> Items { get; set; }
    }

    public class Project
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public DateTime? StartDate { get; set; }
        public DateTime? EndDate { get; set; }
        public DateTime CreateDate { get; set; }
        public DateTime ModifyDate { get; set; }
        public DateTime? LastCommentDate { get; set; }
        public string Tags { get; set; }
        public double NumericPriority { get; set; }
        public bool IsActive { get; set; }
        public bool IsProduct { get; set; }
        public string Abbreviation { get; set; }
        public string MailReplyAddress { get; set; }
        public string Color { get; set; }
        public double Progress { get; set; }
        public EntityType EntityType { get; set; }
        public Owner Owner { get; set; }
        //public LastCommentedUser LastCommentedUser { get; set; }
        //public object Project { get; set; }
        //public Program Program { get; set; }
        public Process Process { get; set; }
        //public Company Company { get; set; }
        public List<object> CustomFields { get; set; }       
    }

    public class ProjectContext
    {
        public string Acid { get; set; }
    }

    public class ProjectsCollection
    {
        public string Next { get; set; }
        public List<Project> Items { get; set; }
    }

    public class Process
    {
        public int Id { get; set; }
        public string Name { get; set; }
    }

    public class Bug
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public DateTime? StartDate { get; set; }
        public DateTime? EndDate { get; set; }
        public DateTime? CreateDate { get; set; }
        public DateTime? ModifyDate { get; set; }
        public DateTime? LastCommentDate { get; set; }
        public string Tags { get; set; }
        public double NumericPriority { get; set; }
        public double Effort { get; set; }
        public double EffortCompleted { get; set; }
        public double EffortToDo { get; set; }
        public double TimeSpent { get; set; }
        public double? TimeRemain { get; set; }
        public object PlannedStartDate { get; set; }
        public object PlannedEndDate { get; set; }
        public EntityType EntityType { get; set; }
        public Owner Owner { get; set; }
        //public LastCommentedUser LastCommentedUser { get; set; }
        public Project Project { get; set; }
        //public Release Release { get; set; }
        public object Iteration { get; set; }
        //public TeamIteration TeamIteration { get; set; }
        //public Team Team { get; set; }
        //public Priority Priority { get; set; }
        public EntityState EntityState { get; set; }
        public object Build { get; set; }
        public UserStory UserStory { get; set; }
        //public Severity Severity { get; set; }
        public List<object> CustomFields { get; set; }
        public AssignmentsCollection Assignments { get; set; } 
    }
    
    public class EntityType
    {
        public int Id { get; set; }
        public string Name { get; set; }
    }

    public class BugsCollection
    {
        public string Next { get; set; }
        public List<Bug> Items { get; set; }
    }    

    public class Release
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public object Description { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public DateTime CreateDate { get; set; }
        public DateTime ModifyDate { get; set; }
        public object LastCommentDate { get; set; }
        public string Tags { get; set; }
        public double NumericPriority { get; set; }
        public bool IsCurrent { get; set; }
        public double Progress { get; set; }
        public EntityType EntityType { get; set; }
        public Owner Owner { get; set; }
        public object LastCommentedUser { get; set; }
        public Project Project { get; set; }
        public List<object> CustomFields { get; set; }
    }

    public class ReleasesCollection
    {
        public string Next { get; set; }
        public List<Release> Items { get; set; }
    }

   
    public class Priority
    {
        public int Id { get; set; }
        public string Name { get; set; }
    }

   
    public class Feature
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public DateTime? StartDate { get; set; }
        public DateTime? EndDate { get; set; }
        public DateTime CreateDate { get; set; }
        public DateTime ModifyDate { get; set; }
        public DateTime? LastCommentDate { get; set; }
        public string Tags { get; set; }
        public double NumericPriority { get; set; }
        public double Effort { get; set; }
        public double EffortCompleted { get; set; }
        public double EffortToDo { get; set; }
        public double TimeSpent { get; set; }
        public double? TimeRemain { get; set; }
        public DateTime? PlannedStartDate { get; set; }
        public DateTime? PlannedEndDate { get; set; }
        public double InitialEstimate { get; set; }
        public EntityType EntityType { get; set; }
        public Owner Owner { get; set; }        
        public Project Project { get; set; }
        public Release Release { get; set; }
        public object Iteration { get; set; }
        public object TeamIteration { get; set; }        
        public Priority Priority { get; set; }
        public EntityState EntityState { get; set; }
        public List<object> CustomFields { get; set; }
    }

    public class FeaturesCollection
    {
        public string Next { get; set; }
        public List<Feature> Items { get; set; }
    }

    public class Task
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public object Description { get; set; }
        public DateTime? StartDate { get; set; }
        public DateTime? EndDate { get; set; }
        public DateTime CreateDate { get; set; }
        public DateTime ModifyDate { get; set; }
        public DateTime? LastCommentDate { get; set; }
        public string Tags { get; set; }
        public double NumericPriority { get; set; }
        public double Effort { get; set; }
        public double EffortCompleted { get; set; }
        public double EffortToDo { get; set; }
        public double Progress { get; set; }
        public double TimeSpent { get; set; }
        public double TimeRemain { get; set; }
        public object PlannedStartDate { get; set; }
        public object PlannedEndDate { get; set; }
        public EntityType EntityType { get; set; }
        public Owner Owner { get; set; }
        //public LastCommentedUser LastCommentedUser { get; set; }
        public Project Project { get; set; }
        public Release Release { get; set; }
        //public Iteration Iteration { get; set; }
        public object TeamIteration { get; set; }
        //public Team Team { get; set; }
        public Priority Priority { get; set; }
        public EntityState EntityState { get; set; }
        public UserStory UserStory { get; set; }
        public List<object> CustomFields { get; set; }
    }

    public class TasksCollection
    {
        public string Next { get; set; }
        public List<Task> Items { get; set; }
    }
}

