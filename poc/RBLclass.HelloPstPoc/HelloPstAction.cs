using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Microsoft.Data.Sqlite;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace RBLclass.HelloPstPoc
{
    internal static class HelloPstAction
    {
        public static void Run(Outlook.Application app)
        {
            var report = new StringBuilder();

            report.AppendLine("Process bitness");
            report.AppendLine("  Is64BitProcess : " + Environment.Is64BitProcess);
            report.AppendLine("  IntPtr.Size    : " + IntPtr.Size);
            report.AppendLine();

            var pstSummary = EnumeratePstStores(app);
            report.Append(pstSummary);
            report.AppendLine();

            var sqliteSummary = ExerciseSqlite();
            report.Append(sqliteSummary);

            MessageBox.Show(
                report.ToString(),
                "RBLclass Hello PST POC",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private static string EnumeratePstStores(Outlook.Application app)
        {
            var sb = new StringBuilder();
            Outlook.NameSpace session = null;
            Outlook.Stores stores = null;

            try
            {
                session = app.Session;
                stores = session.Stores;
                int storeCount = stores.Count;
                sb.AppendLine("Outlook stores (total: " + storeCount + ")");

                Outlook.Store firstPst = null;
                string firstPstName = null;

                for (int i = 1; i <= storeCount; i++)
                {
                    Outlook.Store store = null;
                    try
                    {
                        try
                        {
                            store = stores[i];
                        }
                        catch (Exception ex)
                        {
                            // E.g. PST on OneDrive that hasn't materialised,
                            // an Exchange archive that's offline, etc.
                            // Don't let one bad store kill the whole report.
                            sb.AppendLine("  [" + i + "] (could not open: " +
                                          ex.Message.Trim() + ")");
                            continue;
                        }

                        bool isPst = SafeIsDataFileStore(store);
                        sb.AppendLine(
                            "  [" + i + "] " + store.DisplayName +
                            (isPst ? "  (PST)" : ""));

                        if (firstPst == null && isPst)
                        {
                            firstPst = store;
                            firstPstName = store.DisplayName;
                            store = null;
                        }
                    }
                    finally
                    {
                        if (store != null)
                            Marshal.ReleaseComObject(store);
                    }
                }

                if (firstPst == null)
                {
                    sb.AppendLine();
                    sb.AppendLine("No PST store found.");
                    return sb.ToString();
                }

                Outlook.Folder rootFolder = null;
                Outlook.Folders rootSubfolders = null;
                Outlook.Folder firstSubfolder = null;
                Outlook.Items firstSubfolderItems = null;

                try
                {
                    rootFolder = (Outlook.Folder)firstPst.GetRootFolder();
                    rootSubfolders = rootFolder.Folders;
                    int subfolderCount = rootSubfolders.Count;

                    sb.AppendLine();
                    sb.AppendLine("First PST: " + firstPstName);
                    sb.AppendLine("  Root subfolder count: " + subfolderCount);

                    if (subfolderCount == 0)
                    {
                        sb.AppendLine("  (root has no subfolders)");
                    }
                    else
                    {
                        firstSubfolder = (Outlook.Folder)rootSubfolders[1];
                        firstSubfolderItems = firstSubfolder.Items;
                        int itemCount = firstSubfolderItems.Count;
                        sb.AppendLine("  First subfolder    : " + firstSubfolder.Name);
                        sb.AppendLine("  Items in subfolder : " + itemCount);
                    }
                }
                finally
                {
                    if (firstSubfolderItems != null) Marshal.ReleaseComObject(firstSubfolderItems);
                    if (firstSubfolder != null) Marshal.ReleaseComObject(firstSubfolder);
                    if (rootSubfolders != null) Marshal.ReleaseComObject(rootSubfolders);
                    if (rootFolder != null) Marshal.ReleaseComObject(rootFolder);
                    Marshal.ReleaseComObject(firstPst);
                }
            }
            finally
            {
                if (stores != null) Marshal.ReleaseComObject(stores);
                if (session != null) Marshal.ReleaseComObject(session);
            }

            return sb.ToString();
        }

        private static bool SafeIsDataFileStore(Outlook.Store store)
        {
            try { return store.IsDataFileStore; }
            catch { return false; }
        }

        private static string ExerciseSqlite()
        {
            var sb = new StringBuilder();

            string localAppData = Environment.GetFolderPath(
                Environment.SpecialFolder.LocalApplicationData);
            string dir = Path.Combine(localAppData, "RBLclass");
            Directory.CreateDirectory(dir);
            string dbPath = Path.Combine(dir, "hello-pst-poc.db");

            sb.AppendLine("SQLite");
            sb.AppendLine("  Database     : " + dbPath);

            string connStr = "Data Source=" + dbPath;
            using (var conn = new SqliteConnection(connStr))
            {
                conn.Open();

                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText =
                        "CREATE TABLE IF NOT EXISTS PocPing(" +
                        "Id INTEGER PRIMARY KEY AUTOINCREMENT, " +
                        "Stamp TEXT NOT NULL);";
                    cmd.ExecuteNonQuery();
                }

                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "INSERT INTO PocPing(Stamp) VALUES($s);";
                    cmd.Parameters.AddWithValue("$s",
                        DateTime.UtcNow.ToString("o"));
                    cmd.ExecuteNonQuery();
                }

                long rowCount;
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "SELECT COUNT(*) FROM PocPing;";
                    rowCount = (long)cmd.ExecuteScalar();
                }

                sb.AppendLine("  Row count    : " + rowCount);
            }

            sb.AppendLine("  e_sqlite3.dll: " + ResolveLoadedNativeSqlitePath());

            return sb.ToString();
        }

        private static string ResolveLoadedNativeSqlitePath()
        {
            try
            {
                Process current = Process.GetCurrentProcess();
                foreach (ProcessModule module in current.Modules)
                {
                    if (string.Equals(
                            module.ModuleName,
                            "e_sqlite3.dll",
                            StringComparison.OrdinalIgnoreCase))
                    {
                        return module.FileName;
                    }
                }
                return "(not loaded as a named module)";
            }
            catch (Exception ex)
            {
                return "(lookup failed: " + ex.Message + ")";
            }
        }
    }
}
