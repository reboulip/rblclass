using System;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using RBLclass.Core;
using Serilog;

namespace RBLclass.AddIn
{
    /// <summary>
    /// HTTP implementation of the opt-in update check (v2.7.0.0 E1). Issues a
    /// single anonymous GET for a plain-text version marker on the internal HTTPS
    /// share. Uses a shared static <see cref="HttpClient"/> (correct for .NET 4.8),
    /// which picks up the system/corporate proxy from
    /// <c>WebRequest.DefaultWebProxy</c> automatically. Entirely non-throwing.
    /// </summary>
    internal sealed class HttpUpdateCheckService : IUpdateCheckService
    {
        private static readonly HttpClient _client = new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(5)
        };

        public async Task<string> GetLatestVersionAsync(string url, CancellationToken ct)
        {
            try
            {
                using (var response = await _client.GetAsync(url, ct))
                {
                    if (!response.IsSuccessStatusCode) return null;
                    var body = await response.Content.ReadAsStringAsync();
                    return body?.Trim();
                }
            }
            catch (Exception ex)
            {
                Log.Information(ex, "Update check GET failed (non-critical).");
                return null;
            }
        }
    }
}
