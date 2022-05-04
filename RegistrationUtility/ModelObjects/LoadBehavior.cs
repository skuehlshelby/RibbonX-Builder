
using System;

namespace RegistrationUtility.ModelObjects
{
    internal sealed class LoadBehavior : Enumeration
    {
        private LoadBehavior(int value, string name, string description) : base(value, name)
        {
            Description = description;
        }

        public int Value => value;
        public string Name => name;
        public string Description { get; }

        public static LoadBehavior UnloadedDontLoadAutomatically { get; } = new(0, "Unloaded, Don't Load Automatically", string.Join(Environment.NewLine + Environment.NewLine, "The add-in is not loaded automatically. The user can try to manually load the add-in, or the add-in can be loaded programmatically.", "If the add-in is successfully loaded, the LoadBehavior value remains 0, but the status of the add-in in the 'COM Add-Ins' dialog box is updated to indicate that the add-in is loaded."));

        public static LoadBehavior LoadedDontLoadAutomatically { get; } = new(1, "Loaded, Don't Load Automatically", string.Join(Environment.NewLine + Environment.NewLine, "The add-in is not loaded automatically. The user can try to manually load the add-in, or the add-in can be loaded programmatically.", "Although the 'COM Add-ins' dialog box will indicate that the add-in is loaded, the add-in isn't actually loaded until it is loaded manually or programmatically.", "If the application successfully loads the add-in, the LoadBehavior changes to 0."));

        public static LoadBehavior UnloadedLoadAtStartUp { get; } = new(2, "Unloaded, Load At StartUp", string.Join(Environment.NewLine + Environment.NewLine, "The host application does not try to load the add-in automatically. The user can try to manually load the add-in, or the add-in can be loaded programmatically.", "If the application successfully loads the add-in, the LoadBehavior changes to 3."));

        public static LoadBehavior LoadedLoadAtStartUp { get; } = new(3, "Loaded, Load At StartUp", string.Join(Environment.NewLine + Environment.NewLine, "The host application tries to load the add-in when the application starts. This is the most commonly used LoadBehavior.", "If the add-in is loaded successfully, the LoadBehavior is unchanged. If an error occurs while loading the add-in, the LoadBehavior changes to 2."));

        public static LoadBehavior UnloadedLoadOnDemand { get; } = new(8, "Unloaded, Load On Demand", string.Join(Environment.NewLine + Environment.NewLine, "The add-in is not loaded automatically. The user can try to manually load the add-in, or the add-in can be loaded programmatically.", "If the application successfully loads the add-in, the LoadBehavior value changes to 9."));

        public static LoadBehavior LoadedLoadOnDemand { get; } = new(9, "Loaded, Load On Demand", string.Join(Environment.NewLine + Environment.NewLine, "The add-in is loaded only when the host application requires it, such as when a user selects a UI element that uses functionality in the add-in (for example, a custom button in the Ribbon).", "If the host application successfully loads the add-in, the LoadBehavior value remains 9, but the status of the add-in in the COM Add-ins dialog box is updated to indicate that the add-in is currently loaded. If an error occurs when loading the add-in, the LoadBehavior changes to 8."));

        public static LoadBehavior ConnectFirstTime { get; } = new(16, "Connect First Time", string.Join(Environment.NewLine + Environment.NewLine, "Set this value if you want your add-in to be loaded on demand. The host application loads the add-in when the user runs the host application for the first time after registering the add-in. The next time the user runs the host application, the host application loads any UI elements that are defined by the add-in. However, the add-in isn't loaded until the user selects a UI element that is associated with the add-in.", "When the host application successfully loads the add-in for the first time, the LoadBehavior value remains 16 while the add-in is loaded. After the host application closes, the LoadBehavior changes to 9."));

    }
}
