namespace RaffleBlazor.Models
{
    public class Student
    {
        public string Name { get; }
        public string Id { get; }
        public Department Department { get; set; } = null;

        public Student(string name, string id)
        {
            Name = name ?? "Unknown";
            Id = id ?? "Unknown";
        }

        public override string ToString()
            => $"{Name} ({Id})";
    }
}