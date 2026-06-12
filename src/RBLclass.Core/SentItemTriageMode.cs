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
        AskEveryTime = 3
    }
}
