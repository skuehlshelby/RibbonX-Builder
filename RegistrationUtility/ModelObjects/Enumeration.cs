using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace RegistrationUtility.ModelObjects
{
    internal abstract class Enumeration : IEquatable<Enumeration>, IComparable<Enumeration>
    {
        protected readonly int value;
        protected readonly string name;

        protected Enumeration(int value, string name)
        {
            this.value = value;
            this.name = name;
        }

        public static T[] Values<T>() where T : Enumeration
        {
            return typeof(T)
                .GetProperties(BindingFlags.Public | BindingFlags.Static | BindingFlags.DeclaredOnly)
                .Select(p => p.GetValue(null))
                .OfType<T>()
                .ToArray();
        }

        public static T Parse<T>(int value) where T : Enumeration
        {
            return Values<T>().Single(v => v.value == value);
        }

        public static T Parse<T>(string name) where T : Enumeration
        {
            return Values<T>().Single(v => string.Equals(v.name, name, StringComparison.OrdinalIgnoreCase));
        }

        #region Overrides And Comparison

        public override string ToString()
        {
            return $"({value}) {name}";
        }

        public override int GetHashCode()
        {
            return value;
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as Enumeration);
        }

        public bool Equals(Enumeration other)
        {
            return other != null && value == other.value;
        }

        public int CompareTo(Enumeration other)
        {
            return value.CompareTo(other?.value);
        }

        #endregion

        #region Operators

        public static bool operator ==(Enumeration left, Enumeration right)
        {
            return EqualityComparer<Enumeration>.Default.Equals(left, right);
        }

        public static bool operator !=(Enumeration left, Enumeration right)
        {
            return !(left == right);
        }

        public static bool operator <(Enumeration left, Enumeration right)
        {
            return left.CompareTo(right) < 0;
        }

        public static bool operator <=(Enumeration left, Enumeration right)
        {
            return left.CompareTo(right) <= 0;
        }

        public static bool operator >(Enumeration left, Enumeration right)
        {
            return left.CompareTo(right) > 0;
        }

        public static bool operator >=(Enumeration left, Enumeration right)
        {
            return left.CompareTo(right) >= 0;
        }

        #endregion

    }
}
