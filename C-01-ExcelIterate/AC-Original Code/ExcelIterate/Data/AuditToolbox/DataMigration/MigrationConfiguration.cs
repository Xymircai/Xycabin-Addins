using System.Collections.Generic;

namespace Excel_Iterate
{
    // A detached graph: shared source/target identities survive the copy,
    // while later edits to the window cannot alter a prepared plan.
    internal sealed class MigrationConfiguration
    {
        internal readonly List<MigrationPair> Pairs = new List<MigrationPair>();
        internal readonly List<MigrationBook> SourceBooks = new List<MigrationBook>(), TargetBooks = new List<MigrationBook>();
        internal readonly Dictionary<MigrationPair, MigrationPair> Owners = new Dictionary<MigrationPair, MigrationPair>();
        internal readonly bool ReplaceTargets, AllowEmptySource;
        internal MigrationConfiguration(IList<MigrationPair> pairs, IList<MigrationBook> sources, IList<MigrationBook> targets, bool replace, bool allowEmpty)
        {
            ReplaceTargets = replace; AllowEmptySource = allowEmpty;
            CopyBooks(sources, SourceBooks); CopyBooks(targets, TargetBooks);
            var sides = new Dictionary<MigrationSide, MigrationSide>();
            foreach (var pair in pairs)
            {
                var copy = new MigrationPair { Source = CopySide(pair.Source, sides), Target = CopySide(pair.Target, sides), Enabled = pair.Enabled, ValuesOnly = pair.ValuesOnly, KeepFormat = pair.KeepFormat, FormatOnly = pair.FormatOnly };
                Pairs.Add(copy); Owners.Add(copy, pair);
            }
        }
        private static void CopyBooks(IList<MigrationBook> source, IList<MigrationBook> target)
        {
            foreach (var book in source)
            {
                var copy = new MigrationBook { Path = book.Path, Error = book.Error };
                copy.Sheets.AddRange(book.Sheets); target.Add(copy);
            }
        }
        private static MigrationSide CopySide(MigrationSide side, IDictionary<MigrationSide, MigrationSide> copies)
        {
            MigrationSide copy;
            if (copies.TryGetValue(side, out copy)) return copy;
            copies.Add(side, copy = side.Copy());
            if (side.Trigger != null)
            {
                copy.Trigger.Error = side.Trigger.Error;
                copy.Trigger.Signature = side.Trigger.Signature;
                foreach (var row in side.Trigger.Rows)
                {
                    var item = new MigrationTriggerRow { Row = row.Row, Value = row.Value, Name = row.Name, Path = row.Path, Kind = row.Kind };
                    item.Candidates.AddRange(row.Candidates); copy.Trigger.Rows.Add(item);
                }
            }
            return copy;
        }
        internal void Attach(MigrationPlan plan)
        {
            foreach (var owner in Owners.Values) owner.Instances.Clear();
            foreach (var mapping in plan.Mappings)
            {
                mapping.Pair = Owners[mapping.Pair];
                mapping.Pair.Instances.Add(mapping);
            }
        }
    }
}
