using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;

namespace RegistrationUtility.ModelObjects
{
    public class Bindable : INotifyPropertyChanged
    {
        private readonly HashSet<IBindableProperty> bindableProperties = new HashSet<IBindableProperty>();

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            InvokePropertyChanged(e.PropertyName);
        }

        protected virtual void InvokePropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        protected void Add<T>(string name, T value)
        {
            IBindableProperty<T> property = new BindableProperty<T>(name, value);
            property.PropertyChanged += OnPropertyChanged;
            bindableProperties.Add(property);
        }

        protected void Add<T>(string name, T value, Predicate<T> validation)
        {
            IBindableProperty<T> property = new BindableProperty<T>(name, value, validation);
            property.PropertyChanged += OnPropertyChanged;
            bindableProperties.Add(property);
        }

        protected T Get<T>([CallerMemberName] string name = null)
        {
            return ((IBindableProperty<T>)bindableProperties.First(bp => bp.Name.Equals(name))).Value;          
        }

        protected void Set<T>(T value, [CallerMemberName] string name = null)
        {
            ((IBindableProperty<T>)bindableProperties.First(bp => bp.Name.Equals(name))).Value = value;
        }

        protected void Reset(string propertyName)
        {
            bindableProperties.First(bp => bp.Name.Equals(propertyName)).Reset();
        }



        private interface IBindableProperty : INotifyPropertyChanged
        {
            string Name { get; }

            void Reset();
        }

        private interface IBindableProperty<T> : IBindableProperty
        {
            T Value { get; set; }            
        }

        private class BindableProperty<T> : IBindableProperty<T>, IEquatable<IBindableProperty>
        {
            private readonly string name;
            private readonly T originalValue;
            private T value;
            private readonly Predicate<T> validation;

            public BindableProperty(string name, T value) : this(name, value, (v) => true)
            {

            }

            public BindableProperty(string name, T value, Predicate<T> validation)
            {
                this.name = name;
                this.value = value;
                this.originalValue = value;
                this.validation = validation;
            }

            public event PropertyChangedEventHandler PropertyChanged;

            string IBindableProperty.Name { get => name; }

            T IBindableProperty<T>.Value
            {
                get => value;
                set
                {
                    if (!Equals(this.value, value) && validation.Invoke(value))
                    {
                        this.value = value;
                        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
                    }
                }
            }

            public void Reset()
            {
                value = originalValue;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
            }

            public override int GetHashCode()
            {
                return name.GetHashCode(StringComparison.OrdinalIgnoreCase);
            }

            public override bool Equals(object obj)
            {
                return Equals(obj as IBindableProperty);
            }

            public bool Equals(IBindableProperty other)
            {
                return other != null && other.Name.Equals(name, StringComparison.OrdinalIgnoreCase);
            }
        }
    }
}
