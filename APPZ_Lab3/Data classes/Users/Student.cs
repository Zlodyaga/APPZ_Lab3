namespace APPZ_Lab3.Data_classes
{
    public class Student : IUser
    {
        public int Id { get; set; }
        public string Username { get; set; }
        public string LastName { get; set; }
        public string FirstName { get; set; }
        public string Email { get; set; }
        public string Phone { get; set; }
        public string sex { get; set; }

        public Student(int id, string username, string lastName, string firstName, string email, string phone, string sex)
        {
            Id = id;
            Username = username;
            LastName = lastName;
            FirstName = firstName;
            Email = email;
            Phone = phone;
            this.sex = sex;
        }
    }
}
