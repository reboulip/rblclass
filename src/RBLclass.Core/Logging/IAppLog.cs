namespace RBLclass.Core.Logging
{
    /// <summary>
    /// Narrow logging seam so <c>RBLclass.Core</c> can record actions and
    /// swallowed failures without referencing Serilog (kept out of the
    /// netstandard2.0 core - see CLAUDE.md "Architecture"). Mirrors the
    /// Serilog message-template calls so <c>SerilogAppLog</c> in the AddIn is
    /// a thin forward.
    /// </summary>
    public interface IAppLog
    {
        void Debug(string template, params object[] args);
        void Information(string template, params object[] args);
        void Warning(string template, params object[] args);
        void Warning(System.Exception ex, string template, params object[] args);
        void Error(System.Exception ex, string template, params object[] args);
    }
}
