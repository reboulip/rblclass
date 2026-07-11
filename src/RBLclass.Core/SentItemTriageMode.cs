namespace RBLclass.Core
{
    /// <summary>
    /// What to do with a mail that has just landed in Sent Items (legacy 6c,
    /// reworked). A fixed value is applied automatically with no prompt;
    /// <see cref="AskEveryTime"/> shows the triage dialog. Replaces the former
    /// on/off <c>SentItemTriagePrompt</c> boolean. Conversation widening was
    /// dropped from triage in this rework - it acts on the single sent item.
    /// </summary>
    public enum SentItemTriageMode
    {
        /// <summary>Do nothing (the mail stays in Sent Items).</summary>
        Leave = 0,

        /// <summary>Move the sent item to the Inbox automatically.</summary>
        MoveToInbox = 1,

        /// <summary>Delete the sent item automatically.</summary>
        Delete = 2,

        /// <summary>Show the prompt so the user picks each time.</summary>
        AskEveryTime = 3,

        /// <summary>Move the sent item to the user-configured archive folder automatically (v2.7.0.0 B2).</summary>
        MoveToArchive = 4,

        /// <summary>Copy the sent item to the Inbox, leaving the original in Sent Items (v2.9.0.0 C1).</summary>
        CopyToInbox = 5
    }
}
