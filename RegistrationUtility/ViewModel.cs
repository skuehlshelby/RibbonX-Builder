using RegistrationUtility.ModelObjects;
using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using OpenFileDialog = System.Windows.Forms.OpenFileDialog;

namespace RegistrationUtility
{
    internal sealed class ViewModel : Model
    {
        public ViewModel()
        {
            LoadBehaviors = Enumeration.Values<LoadBehavior>();
            Applications = Enumeration.Values<HostApplication>();
            SelectDll = new SelectDllCommand(this);
            Add(nameof(FolderIconVisibility), Visibility.Hidden);
            Add(nameof(AvailableTypes), Array.Empty<Type>());
        }

        public Visibility FolderIconVisibility
        {
            get { return Get<Visibility>(); }
            set { Set(value); }
        }

        public Type[] AvailableTypes
        {
            get { return Get<Type[]>(); }
            set { Set(value); }
        }

        public ICommand SelectDll { get; }

        public LoadBehavior[] LoadBehaviors { get; }

        public HostApplication[] Applications { get; }

        private class SelectDllCommand : ICommand
        {
            private readonly ViewModel model;

            public SelectDllCommand(ViewModel model)
            {
                this.model = model;
            }

            public event EventHandler CanExecuteChanged;

            public bool CanExecute(object parameter)
            {
                return true;
            }

            public void Execute(object parameter)
            {
                using (var folderDialog = new OpenFileDialog())
                {
                    folderDialog.Filter = "DLL (*.dll)| *.dll";

                    if (folderDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        model.DLL = folderDialog.FileName;
                        model.AvailableTypes = model.ComVisibleClasses().ToArray();
                        model.Reset(nameof(TargetType));
                    }
                }
            }
        }
    }
}
