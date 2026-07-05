using System;
using RBLclass.Core.Logging;
using Serilog;

namespace RBLclass.AddIn.Logging
{
    /// <summary>Forwards <see cref="IAppLog"/> calls to the add-in's static Serilog logger.</summary>
    public sealed class SerilogAppLog : IAppLog
    {
        public void Debug(string template, params object[] args) => Log.Debug(template, args);
        public void Information(string template, params object[] args) => Log.Information(template, args);
        public void Warning(string template, params object[] args) => Log.Warning(template, args);
        public void Warning(Exception ex, string template, params object[] args) => Log.Warning(ex, template, args);
        public void Error(Exception ex, string template, params object[] args) => Log.Error(ex, template, args);
    }
}
