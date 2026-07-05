using System;

namespace RBLclass.Core.Logging
{
    /// <summary>Default <see cref="IAppLog"/> for callers that don't inject one (e.g. existing tests).</summary>
    public sealed class NullAppLog : IAppLog
    {
        public static readonly NullAppLog Instance = new NullAppLog();

        private NullAppLog() { }

        public void Debug(string template, params object[] args) { }
        public void Information(string template, params object[] args) { }
        public void Warning(string template, params object[] args) { }
        public void Warning(Exception ex, string template, params object[] args) { }
        public void Error(Exception ex, string template, params object[] args) { }
    }
}
