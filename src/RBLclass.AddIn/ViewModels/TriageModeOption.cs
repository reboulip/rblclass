using RBLclass.Core;

namespace RBLclass.AddIn.ViewModels
{
    /// <summary>A sent-item triage mode paired with a friendly label for the settings dropdown.</summary>
    public sealed class TriageModeOption
    {
        public TriageModeOption(SentItemTriageMode mode, string label)
        {
            Mode = mode;
            Label = label;
        }

        public SentItemTriageMode Mode { get; }
        public string Label { get; }
    }
}
