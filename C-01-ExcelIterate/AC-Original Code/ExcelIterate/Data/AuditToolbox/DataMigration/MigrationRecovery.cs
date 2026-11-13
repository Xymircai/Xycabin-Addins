using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Excel_Iterate
{
    internal sealed class MigrationRecovery
    {
        internal sealed class Point { internal string Path, Backup, Hash, Displaced; internal bool Restored, Cleared; }
        private readonly List<Point> points = new List<Point>();
        internal int RestorableCount { get { return points.Count(p => !p.Restored); } }
        internal List<Point> Backups { get { return points.Where(p => !p.Cleared).ToList(); } }
        internal List<Point> Pending { get { return points.Where(p => !p.Restored && !p.Cleared).Reverse().ToList(); } }
        internal void Record(string path, string backup)
        {
            if (string.IsNullOrEmpty(backup) || !File.Exists(backup) || points.Any(p => string.Equals(p.Backup, backup, StringComparison.OrdinalIgnoreCase))) return;
            points.Add(new Point { Path = path, Backup = backup, Hash = DataMigrationHost.Digest(path) });
        }
        internal static IEnumerable<string> Files(Point point)
        {
            if (!string.IsNullOrEmpty(point.Backup)) yield return point.Backup;
            if (!string.IsNullOrEmpty(point.Displaced)) yield return point.Displaced;
        }
        internal static List<KeyValuePair<Point, string>> Restore(IList<Point> pending)
        { return Apply(pending, p => { p.Displaced = MigrationOutputTransaction.Restore(p.Path, p.Backup, p.Hash); p.Restored = true; }); }
        internal static List<KeyValuePair<Point, string>> Clear(IList<Point> pending)
        { return Apply(pending, p => { MigrationOutputTransaction.DeleteBackups(Files(p).ToList()); p.Cleared = true; }); }
        private static List<KeyValuePair<Point, string>> Apply(IList<Point> pending, Action<Point> action)
        {
            var results = new List<KeyValuePair<Point, string>>();
            foreach (var point in pending)
            {
                string failure = "";
                try { action(point); } catch (Exception error) { failure = error.Message; }
                results.Add(new KeyValuePair<Point, string>(point, failure));
            }
            return results;
        }
    }
}
