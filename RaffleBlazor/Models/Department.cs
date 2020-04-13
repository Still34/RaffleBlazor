using System.Collections.Concurrent;
using System.Collections.Generic;

namespace RaffleBlazor.Models
{
    public class Department
    {
        private readonly List<Student> _students = new List<Student>();

        public string Name { get; }

        public IReadOnlyCollection<Student> Students
            => _students;

        public int MaxStudents { get; }

        public bool HasReachedLimits
            => _students.Count >= MaxStudents;

        public Department(string name, int maxStudents)
        {
            Name = name ?? "Unknown";
            MaxStudents = maxStudents > 0 ? maxStudents : 99;
        }

        public void AddStudent(Student student)
        {
            student.Department = this;
            _students.Add(student);
        }

        public override string ToString()
            => Name;
    }
}