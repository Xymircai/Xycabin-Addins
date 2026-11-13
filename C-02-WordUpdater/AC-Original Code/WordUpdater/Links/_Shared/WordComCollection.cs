namespace WordUpdater
{
    internal static class WordComCollection
    {

        // ---------- COM 集合索引访问：先试 coll[i]，再试 coll.Item(i) ----------
        internal static object Item(object coll, int i)
        {
            dynamic d = coll;
            try { return d[i]; }
            catch { }
            return d.Item(i);
        }
    }
}
