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
    }

    public class EntityState
    {
        public int Id { get; set; }
        public string Name { get; set; }
    }

    public class Feature
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
        public double Effort { get; set; }
        public double EffortCompleted { get; set; }
        public double EffortToDo { get; set; }
        public Owner Owner { get; set; }
        public EntityState EntityState { get; set; }
        public Feature Feature { get; set; }
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
        public Owner Owner { get; set; }
    }

    public class ProjectsCollection
    {
        public string Next { get; set; }
        public List<Project> Items { get; set; }
    }
}

