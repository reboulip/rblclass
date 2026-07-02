using System.Threading;
using System.Threading.Tasks;

namespace RBLclass.Core
{
    /// <summary>
    /// Fetches the "latest available version" marker for the opt-in update check
    /// (v2.7.0.0 E1). The concrete implementation (HTTP) lives in the add-in shell
    /// so Core stays network-free; this seam keeps the decision logic testable.
    /// </summary>
    public interface IUpdateCheckService
    {
        /// <summary>
        /// Fetch the latest-version string from <paramref name="url"/>, or null on
        /// any failure (network unreachable, timeout, non-success status, parse
        /// error). Implementations must be entirely non-throwing.
        /// </summary>
        Task<string> GetLatestVersionAsync(string url, CancellationToken ct);
    }
}
