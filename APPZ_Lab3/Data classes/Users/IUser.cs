namespace APPZ_Lab3.Data_classes
{
    public interface IUser
    {
        int Id { get; set; } // Ідентифікатор користувача
        string Username { get; set; } // Логін користувача
        string LastName { get; set; } // Прізвище користувача
        string FirstName { get; set; } // Ім'я користувача
        string Email { get; set; } // Електронна пошта користувача
        string Phone { get; set; } // Номер телефону користувача
        string sex { get; set; } // Стать користувача
    }
}
