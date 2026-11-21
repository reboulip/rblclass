namespace RBLclass.Core
{
    /// <summary>
    /// File format written by the pane's Save-as button (v2.13.0.0 C1).
    /// </summary>
    public enum SaveAsFormat
    {
        /// <summary>Outlook message (.msg), the historical behaviour.</summary>
        Msg = 0,

        /// <summary>PDF, converted through Word (late-bound, optional at runtime).</summary>
        Pdf = 1,

        /// <summary>Setting only: the Save-as picker asks for the format each time.</summary>
        Ask = 2
    }
}
