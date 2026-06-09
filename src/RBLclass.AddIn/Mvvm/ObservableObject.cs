using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace RBLclass.AddIn.Mvvm
{
    /// <summary>
    /// Minimal INotifyPropertyChanged base. Hand-rolled (not
    /// CommunityToolkit.Mvvm) because that package drags System.Memory/Unsafe
    /// versions that break the SQLite stack in this redirect-less COM host.
    /// </summary>
    public abstract class ObservableObject : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        protected bool SetProperty<T>(ref T field, T value,
                                      [CallerMemberName] string propertyName = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value))
                return false;

            field = value;
            OnPropertyChanged(propertyName);
            return true;
        }

        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
