using System;
using System.Collections.Concurrent;
using System.Collections.Generic;

namespace RaffleBlazor.Models
{
    public class RaffleEntry
    {
        public DateTime SubmittedAt { get; }
        public Student Student { get; }

        public RaffleEntry(DateTime submittedAt, Student student)
        {
            Student = student;
            SubmittedAt = submittedAt;
        }

        public IReadOnlyCollection<Department> Preferences
            => _preferences;

        private readonly List<Department> _preferences = new List<Department>();

        internal RaffleEntry AddPreference(Department department)
        {
            _preferences.Add(department);
            return this;
        }
    }
}