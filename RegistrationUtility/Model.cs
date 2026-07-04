using Microsoft.Win32;
using RegistrationUtility.ModelObjects;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Input;

namespace RegistrationUtility
{
    internal class Model : Bindable
    {
        public Model()
        {
            Add(nameof(DLL), string.Empty, file => IsDll(file));
            Add<Type>(nameof(TargetType), null, t => IsRegisterable(t));
            Add(nameof(LoadBehavior), LoadBehavior.UnloadedDontLoadAutomatically);
            Add(nameof(Application), HostApplication.Access);
            Add(nameof(FriendlyName), string.Empty);
            Add(nameof(Description), string.Empty);

            Register = new RegisterCommand(this);
            Unregister = new UnregisterCommand(this);
        }

        public string DotNetCategoryId { get; } = "{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}";

        public string DLL
        {
            get { return Get<string>(); }
            set { Set(value); }
        }

        public Type TargetType
        {
            get { return Get<Type>(); }
            set { Set(value); }
        }

        public LoadBehavior LoadBehavior
        {
            get { return Get<LoadBehavior>(); }
            set { Set(value); }
        }

        public HostApplication Application
        {
            get { return Get<HostApplication>(); }
            set { Set(value); }
        }

        public string FriendlyName
        {
            get { return Get<string>(); }
            set { Set(value); }
        }

        public string Description
        {
            get { return Get<string>(); }
            set { Set(value); }
        }

        public ICollection<Type> ComVisibleClasses()
        {
            if (!string.IsNullOrWhiteSpace(DLL))
            {
                return Assembly
                    .LoadFrom(DLL)
                    .GetTypes()
                    .Where(t => IsRegisterable(t))
                    .ToArray();
            }
            else
            {
                return Array.Empty<Type>();
            }
        }

        private static bool IsDll(string filePath)
        {
            return !string.IsNullOrWhiteSpace(filePath) && File.Exists(filePath) && filePath.EndsWith(".dll", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsRegisterable(Type type)
        {
            return type != null && type.IsDefined(typeof(ComVisibleAttribute)) && type.IsDefined(typeof(ProgIdAttribute)) && type.IsDefined(typeof(GuidAttribute));
        }

        public ICommand Register { get; }

        private sealed class RegisterCommand : ICommand
        {
            private readonly Model model;

            public RegisterCommand(Model model)
            {
                this.model = model;
                model.PropertyChanged += OnPropertyChange;
            }

            private void OnPropertyChange(object sender, System.ComponentModel.PropertyChangedEventArgs e)
            {
                if (new string[] { nameof(DLL), nameof(TargetType), nameof(FriendlyName), nameof(Description) }.Contains(e.PropertyName))
                {
                    CanExecuteChanged?.Invoke(this, EventArgs.Empty);
                }
            }

            public event EventHandler CanExecuteChanged;

            bool ICommand.CanExecute(object parameter)
            {
                return IsDll(model.DLL) && IsRegisterable(model.TargetType) && !string.IsNullOrEmpty(model.FriendlyName) && !string.IsNullOrEmpty(model.Description);
            }

            void ICommand.Execute(object parameter)
            {
                try
                {
                    string guid = $"{{{model.TargetType.GetCustomAttribute<GuidAttribute>().Value}}}";
                    string progId = model.TargetType.GetCustomAttribute<ProgIdAttribute>().Value;
                    string assemblyName = model.TargetType.Assembly.FullName;

                    Registry.SetValue($"HKEY_CURRENT_USER\\SOFTWARE\\Classes\\{progId}", string.Empty, progId, RegistryValueKind.String);


                    Registry.SetValue($"HKEY_CURRENT_USER\\SOFTWARE\\Classes\\{progId}\\CLSID", string.Empty, guid);


                    Registry.SetValue($"HKEY_CURRENT_USER\\SOFTWARE\\Classes\\CLSID\\{guid}", string.Empty, progId, RegistryValueKind.String);


                    Registry.SetValue($"HKEY_CURRENT_USER\\SOFTWARE\\Classes\\CLSID\\{guid}\\ProgId", string.Empty, progId, RegistryValueKind.String);


                    Registry.SetValue($"HKEY_CURRENT_USER\\SOFTWARE\\Classes\\Wow6432Node\\CLSID\\{guid}\\ProgId", string.Empty, progId, RegistryValueKind.String);

                    using (RegistryKey inProcessServerEntries = Registry.CurrentUser.CreateSubKey($"SOFTWARE\\Classes\\CLSID\\{guid}\\inProcServer32"))
                    {
                        inProcessServerEntries.SetValue(string.Empty, "mscoree.dll");
                        inProcessServerEntries.SetValue("ThreadingModel", "Both");
                        inProcessServerEntries.SetValue("Class", progId);
                        inProcessServerEntries.SetValue("Assembly", assemblyName);
                        inProcessServerEntries.SetValue("RuntimeVersion", "v4.0.30319");
                        inProcessServerEntries.SetValue("CodeBase", $"file:///{model.DLL}");
                    }

                    using (RegistryKey inProcessServerEntries = Registry.CurrentUser.CreateSubKey($"SOFTWARE\\Classes\\CLSID\\{guid}\\inProcServer32\\1.0.0.0"))
                    {
                        inProcessServerEntries.SetValue("Class", progId);
                        inProcessServerEntries.SetValue("Assembly", assemblyName);
                        inProcessServerEntries.SetValue("RuntimeVersion", "v4.0.30319");
                        inProcessServerEntries.SetValue("CodeBase", $"file:///{model.DLL}");
                    }

                    using (RegistryKey implementedCategoryEntry = Registry.CurrentUser.CreateSubKey($"SOFTWARE\\Classes\\CLSID\\{guid}\\Implemented Categories\\{model.DotNetCategoryId}"))
                    {

                    }

                    using (RegistryKey implementedCategoryEntry = Registry.CurrentUser.CreateSubKey($"SOFTWARE\\Classes\\Wow6432Node\\CLSID\\{guid}\\Implemented Categories\\{model.DotNetCategoryId}"))
                    {

                    }

                    using (RegistryKey classRecord = Registry.CurrentUser.CreateSubKey("Software\\Classes\\Record\\{289E9AF1-4973-11D1-AE81-00A0C90F26F4}\\1.0.0.0"))
                    {
                        classRecord.SetValue("Class", "Extensibility.ext_ConnectMode");
                        classRecord.SetValue("Assembly", assemblyName);
                        classRecord.SetValue("RuntimeVersion", "v4.0.30319");
                        classRecord.SetValue("CodeBase", $"file:///{model.DLL}");
                    }

                    using (RegistryKey inProcessServerEntries = Registry.CurrentUser.CreateSubKey($"SOFTWARE\\Classes\\Wow6432Node\\CLSID\\{guid}\\inProcServer32"))
                    {
                        inProcessServerEntries.SetValue(string.Empty, "mscoree.dll");
                        inProcessServerEntries.SetValue("ThreadingModel", "Both");
                        inProcessServerEntries.SetValue("Class", progId);
                        inProcessServerEntries.SetValue("Assembly", assemblyName);
                        inProcessServerEntries.SetValue("RuntimeVersion", "v4.0.30319");
                        inProcessServerEntries.SetValue("CodeBase", $"file:///{model.DLL}");
                    }

                    using (RegistryKey inProcessServerEntries = Registry.CurrentUser.CreateSubKey($"SOFTWARE\\Classes\\Wow6432Node\\CLSID\\{guid}\\inProcServer32\\1.0.0.0"))
                    {
                        inProcessServerEntries.SetValue("Class", progId);
                        inProcessServerEntries.SetValue("Assembly", assemblyName);
                        inProcessServerEntries.SetValue("RuntimeVersion", "v4.0.30319");
                        inProcessServerEntries.SetValue("CodeBase", $"file:///{model.DLL}");
                    }

                    using (RegistryKey applicationSpecific = Registry.CurrentUser.CreateSubKey($"SOFTWARE\\Microsoft\\Office\\{model.Application}\\Addins\\{progId}"))
                    {
                        applicationSpecific.SetValue("FriendlyName", model.FriendlyName, RegistryValueKind.String);
                        applicationSpecific.SetValue("Description", model.Description, RegistryValueKind.String);
                        applicationSpecific.SetValue("LoadBehavior", model.LoadBehavior.Value, RegistryValueKind.DWord);
                    }
                }
                catch (Exception)
                {

                }
            }
        }

        public ICommand Unregister { get; }

        private sealed class UnregisterCommand : ICommand
        {
            private readonly Model model;

            public UnregisterCommand(Model model)
            {
                this.model = model;
                this.model.PropertyChanged += OnPropertyChanged;
            }

            private void OnPropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
            {
                if (new string[] { nameof(TargetType), nameof(Application) }.Contains(e.PropertyName))
                {
                    CanExecuteChanged?.Invoke(this, EventArgs.Empty);
                }
            }

            public event EventHandler CanExecuteChanged;

            bool ICommand.CanExecute(object parameter)
            {
                return IsRegisterable(model.TargetType) && model.Application != null;
            }

            void ICommand.Execute(object parameter)
            {
                try
                {
                    string progId = model.TargetType.GetCustomAttribute<ProgIdAttribute>().Value;
                    string guid = $"{{{model.TargetType.GetCustomAttribute<GuidAttribute>().Value}}}";

                    using (RegistryKey currentUser = Registry.CurrentUser)
                    {
                        if (currentUser.OpenSubKey($"SOFTWARE\\Classes\\{progId}") != null)
                            currentUser.DeleteSubKeyTree($"SOFTWARE\\Classes\\{progId}");

                        if (currentUser.OpenSubKey($"SOFTWARE\\Classes\\CLSID\\{guid}") != null)
                            currentUser.DeleteSubKeyTree($"SOFTWARE\\Classes\\CLSID\\{guid}");

                        if (currentUser.OpenSubKey($"SOFTWARE\\Classes\\Wow6432Node\\CLSID\\{guid}") != null)
                            currentUser.DeleteSubKeyTree($"SOFTWARE\\Classes\\Wow6432Node\\CLSID\\{guid}");

                        if (currentUser.OpenSubKey($"SOFTWARE\\Microsoft\\Office\\{model.Application}\\Addins\\{progId}") != null)
                            currentUser.DeleteSubKeyTree($"SOFTWARE\\Microsoft\\Office\\{model.Application}\\Addins\\{progId}");
                    }
                }
                catch (Exception)
                {

                }
            }
        }
    }
}
