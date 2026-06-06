using System;
using System.Runtime.InteropServices;

namespace RBLclass.Outlook.Adapter
{
    /// <summary>
    /// Owns a single Outlook COM object and releases it deterministically via
    /// <see cref="Marshal.ReleaseComObject"/> on <see cref="Dispose"/>. This is
    /// the CLAUDE.md "wrap every COM object in a ComRef&lt;T&gt;" rule made
    /// concrete: long-lived references leak and crash Outlook, so every folder,
    /// store and collection touched during the walk lives inside one of these
    /// in a <c>using</c> scope.
    /// </summary>
    public sealed class ComRef<T> : IDisposable where T : class
    {
        private T _value;

        public ComRef(T value)
        {
            _value = value;
        }

        /// <summary>The wrapped COM object. Throws once disposed.</summary>
        public T Value =>
            _value ?? throw new ObjectDisposedException(typeof(T).Name);

        public void Dispose()
        {
            var v = _value;
            _value = null;
            if (v != null && Marshal.IsComObject(v))
                Marshal.ReleaseComObject(v);
        }
    }
}
