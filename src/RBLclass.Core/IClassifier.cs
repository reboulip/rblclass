namespace RBLclass.Core
{
    /// <summary>
    /// The flagship action (legacy 5b): file mail items into one or more
    /// folders. All decision logic lives here and is unit-tested with a faked
    /// <see cref="IMailStore"/>; the Outlook mechanics live in the adapter.
    /// </summary>
    public interface IClassifier
    {
        ClassifyResult Classify(ClassifyRequest request);
    }
}
