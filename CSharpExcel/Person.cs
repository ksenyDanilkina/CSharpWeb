namespace CSharpExcel
{
    class Person
    {
        public string Name { get; }

        public string Surname { get; }

        public int Age { get; }

        public string PhoneNumber { get; }

        public Person(string name, string surname, int age, string phoneNumber)
        {
            Name = name;
            Surname = surname;
            Age = age;
            PhoneNumber = phoneNumber;
        }
    }
}
