namespace RegistrationUtility.ModelObjects
{
    internal class HostApplication : Enumeration
    {
        private HostApplication(int value, string name) : base(value, name)
        {

        }

        public static HostApplication Access { get; } = new(0, nameof(Access));

        public static HostApplication Excel { get; } = new(1, nameof(Excel));

        public static HostApplication Onenote { get; } = new(2, nameof(Onenote));

        public static HostApplication Outlook { get; } = new(3, nameof(Outlook));

        public static HostApplication Word { get; } = new(4, nameof(Word));

        public override string ToString()
        {
            return name;
        }
    }
}
