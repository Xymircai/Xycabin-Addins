using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text.RegularExpressions;
using Excel_Iterate;

namespace Excel_Iterate.TestHarness
{
    // Managed host fixtures exercise the real adapter and disk copy pipeline.
    // These are NOT WPS or format-compatibility acceptance tests.
    public static class DataMigrationProbe
    {
        private static int count;
        private static void Check(bool value, string message) { if (!value) throw new Exception(message); count++; Console.WriteLine("PASS: " + message); }
        private static void Reject(Action action, string message)
        { bool rejected = false; try { action(); } catch (InvalidOperationException) { rejected = true; } Check(rejected, message); }
        private static MigrationFormulaContext Context(int row, int column)
        {
            var c = new MigrationFormulaContext { SourcePath = @"C:\fixtures\源.xlsx", SourceSheet = "源表", TargetSheet = "目标表", RowDelta = row, ColumnDelta = column };
            c.TargetSheets.UnionWith(new[] { "目标表", "明细", "O'Brien" }); c.ExternalFiles["第三方.xlsx"] = @"C:\third\第三方.xlsx"; return c;
        }
        private static void NameMatchingChecks()
        {
            // Generic names retain numbering, punctuation, partial-match and duplicate-path cases.
            const string mergedName = "！00-示例集团（合并）";
            Check(MigrationNameIndex.NormalizeName(mergedName + ".xlsx") == MigrationNameIndex.NormalizeName("！00-示例集团(合并)")
                && MigrationNameIndex.NormalizeName("03-01-示例 文化.xlsx") == "示例文化" && MigrationNameIndex.NormalizeName("2024年报") == "2024年报", "names normalise extension, width, numbering, ！ and spaces only");
            var pool = new[] { @"D:\目标\！00-示例集团（合并）.xlsx", @"D:\目标\！01-示例集团.xlsx", @"D:\目标\！03-01-测试市示例传媒有限公司.xlsx", @"D:\目标\！02-示例子公司.xlsx" };
            var exact = new MigrationTriggerRow { Row = 2, Value = "x", Name = mergedName }; new MigrationNameIndex(pool).Match(exact);
            Check(exact.Kind == "精确" && exact.Path == pool[0] && exact.Ready, "an exact name picks its workbook even when a near name exists");
            var fuzzy = new MigrationTriggerRow { Row = 3, Value = "x", Name = "示例传媒有限公司" }; new MigrationNameIndex(pool).Match(fuzzy);
            Check(fuzzy.Kind == "缺失" && fuzzy.Path.Length == 0 && fuzzy.Candidates.Count == 0 && !fuzzy.Ready, "a partial name no longer matches (fuzzy matching removed)");
            var missing = new MigrationTriggerRow { Row = 4, Value = "x", Name = "完全无关的单位" }; new MigrationNameIndex(pool).Match(missing);
            Check(missing.Kind == "缺失" && missing.Path.Length == 0, "an unrelated name matches nothing rather than the closest file");
            var twins = new MigrationTriggerRow { Row = 5, Value = "x", Name = "示例集团" }; new MigrationNameIndex(new[] { @"D:\a\示例集团.xlsx", @"D:\b\01-示例集团.xlsx", @"D:\a\示例集团乙.xlsx" }).Match(twins);
            Check(twins.Kind == "多个候选" && twins.Path.Length == 0 && twins.Candidates.Count == 2, "the same name in two folders requires a choice; a longer name is not a candidate");
            var spare = new MigrationPlan(); spare.Unused.Add(@"D:\a\多余.xlsx");
            Check(DataMigrationForm.PlanSummary(spare).Contains("1 个目标工作簿没有触发行匹配，不会写入"), "added targets no trigger row matched are reported, not written");
        }
        private static void FormulaChecks()
        {
            var c = Context(99, 2);
            var examples = new Dictionary<string, string> {
                {"=A1", "=C100"}, {"=$A1+A$1+$A$1", "=$A100+C$1+$A$1"},
                {"=SUM(A1:B2)", "=SUM(C100:D101)"}, {"=SUM(A:A,$B:$B)", "=SUM(C:C,$B:$B)"},
                {"=SUM(1:3,$2:$4)", "=SUM(100:102,$2:$4)"}, {"=明细!A1", "='明细'!C100"},
                {"=源表!A1", "='目标表'!C100"}, {"='O''Brien'!$A1", "='O''Brien'!$A100"},
                {"='[第三方.xlsx]明细'!A1:B2", "='C:\\third\\[第三方.xlsx]明细'!A1:B2"},
                {"=A1+'[第三方.xlsx]明细'!A1", "=C100+'C:\\third\\[第三方.xlsx]明细'!A1"},
                {"='[源.xlsx]源表'!A1", "='目标表'!C100"},
                {"=IF(A1=\"A1\",\"a\"\"B2\",1.5E10)", "=IF(C100=\"A1\",\"a\"\"B2\",1.5E10)"},
                {"=LOG10 (A1)+TRUE", "=LOG10 (C100)+TRUE"}
            };
            foreach (var pair in examples) Check(DataMigrationEngine.Translate(pair.Key, c, "源表").Text == pair.Value, "formula " + pair.Key);
            Reject(delegate { DataMigrationEngine.Translate("=缺表!A1", c, "源表"); }, "missing internal sheet blocks migration");
            string missingSheet = null;
            try { DataMigrationEngine.Translate("=缺表!A1", c, "源表"); } catch (MigrationMissingSheetException error) { missingSheet = error.Sheet; }
            Check(missingSheet == "缺表", "a sheet the target lacks is reported as such, so that cell can be pasted as a value");
            Reject(delegate { DataMigrationEngine.Translate("=A1", Context(-1, 0), "源表"); }, "negative row blocks migration");
            Reject(delegate { DataMigrationEngine.Translate("=A1", Context(0, -1), "源表"); }, "negative column blocks migration");
            c.MaxRow = 65536; Reject(delegate { DataMigrationEngine.Translate("=A65536", c, "源表"); }, "legacy target row limit enforced");
            foreach (string formula in new[] { "=Table1[金额]", "='明细:源表'!A1", "=A1#", "={1,2}", "=INDIRECT(\"明细!A1\")", "=未知名称", "='[2]明细'!A1", "='[不存在.xlsx]明细'!A1", "=A1:明细!B2" })
                Reject(delegate { DataMigrationEngine.Translate(formula, c, "源表"); }, "unsupported formula rejected: " + formula);
            c = Context(99, 0);
            c.Names.Add(new MigrationName { Name = "Rate", Formula = "='源表'!$A$1" });
            c.Names.Add(new MigrationName { Name = "Amount", Formula = "=Rate*2" });
            c.Names.Add(new MigrationName { Scope = "源表", Name = "LocalRate", Formula = "=Rate" });
            var required = DataMigrationEngine.RequiredNames(new[] { DataMigrationEngine.Translate("=Amount+LocalRate", c, "源表") }, c);
            Check(required.Count == 3 && required[0].Name == "Rate", "name dependency closure ordered before users");
            Check(required[0].Formula == "='目标表'!$A$1" && required[2].Scope == "目标表", "names and scopes map to target");
            c.Names.Add(new MigrationName { Name = "XFE1", Formula = "=2" });
            Check(DataMigrationEngine.Translate("=XFE1", c, "源表").Dependencies.Count == 1, "legal defined name beyond A1 column limit is not rewritten");
            c.ExternalFiles["源.xlsx"] = @"C:\other\源.xlsx";
            Check(DataMigrationEngine.Translate("='[源.xlsx]明细'!A1", c, "源表").Text == "='C:\\other\\[源.xlsx]明细'!A1", "same filename in another directory stays external");
            HashSet<string> ambiguous;
            var map = DataMigrationHost.MapExternalFiles(new[] { @"C:\a\同名源", @"C:\b\同名源", @"C:\third\第三方.xlsx" }, out ambiguous);
            Check(!map.ContainsKey("同名源") && ambiguous.Contains("同名源") && map["第三方.xlsx"] == @"C:\third\第三方.xlsx", "a file name linked under two paths is marked ambiguous instead of failing the workbook");
            var d = Context(99, 2); d.AmbiguousFiles.Add("同名源");
            Check(DataMigrationEngine.Translate("=A1+'[第三方.xlsx]明细'!A1", d, "源表").Text == "=C100+'C:\\third\\[第三方.xlsx]明细'!A1", "an unused same-name link no longer blocks migration");
            Reject(delegate { DataMigrationEngine.Translate("='[同名源]明细'!A1", d, "源表"); }, "a formula that uses a same-name link is still blocked");
            NameMatchingChecks();
            c.Names.Add(new MigrationName { Name = "Relative", Formula = "=源表!A1" });
            Reject(delegate { DataMigrationEngine.RequiredNames(new[] { DataMigrationEngine.Translate("=Relative", c, "源表") }, c); }, "relative defined name fails closed");
            c.Names.Add(new MigrationName { Name = "Cycle1", Formula = "=Cycle2" }); c.Names.Add(new MigrationName { Name = "Cycle2", Formula = "=Cycle1" });
            Reject(delegate { DataMigrationEngine.RequiredNames(new[] { DataMigrationEngine.Translate("=Cycle1", c, "源表") }, c); }, "name cycles rejected");
            Check(MigrationArea.Parse("$B$2:D4").Address == "B2:D4", "rectangular addresses normalized");
            Check(!MigrationArea.Parse("B2:C3").Contains(MigrationArea.Parse("A2:C3")), "partial merged area does not fit selection");
            Reject(delegate { MigrationArea.Parse("B2:A1"); }, "reversed area rejected");
            Check(DataMigrationEngine.OutputPath(@"C:\fixtures\a.xlsm", p => p.EndsWith("_复制结果.xlsm")) == @"C:\fixtures\a_复制结果_2.xlsm", "output collision preserves extension and existing file");
            c = Context(0, 0);
            Check(DataMigrationEngine.ComparisonFormula("=SUM($C$6:B5)", c) == DataMigrationEngine.ComparisonFormula("=SUM(B5:$C$6)", c), "reversed range endpoints compare equal after WPS normalization");
            Check(DataMigrationEngine.ComparisonFormula("=SUM($C5:B$6)", c) == DataMigrationEngine.ComparisonFormula("=SUM(B5:$C$6)", c), "range normalization preserves each axis dollar locks");
            Check(DataMigrationEngine.ComparisonFormula("=SUM($C$6:B5)", c) != DataMigrationEngine.ComparisonFormula("=SUM($B$5:C6)", c), "different dollar locks are not hidden by comparison");
            Check(DataMigrationEngine.ComparisonFormula("=\"$C$6:B5\"", c) == "=\"$C$6:B5\"", "formula string literals are not canonicalized as references");
            Check(DataMigrationEngine.DirectSelfReference("=$C4", "目标表", 4, 3) && !DataMigrationEngine.DirectSelfReference("='明细'!$C4", "目标表", 4, 3), "direct destination self-reference is distinguished from another sheet");
            Check(DataMigrationEngine.OutputPath(@"C:\fixtures\审计底稿（甲）.xlsx", p => false) == @"C:\fixtures\审计底稿（甲）_复制结果.xlsx", "formal output preserves Unicode workbook name without temporary identifiers");
        }
        private static void PackageChecks(string root)
        {
            string original = Path.Combine(root, "ISO日期.xlsx"), copy = Path.Combine(root, "owned-date-copy.xlsx");
            using (var file = File.Create(original)) using (var zip = new ZipArchive(file, ZipArchiveMode.Create))
            {
                using (var writer = new StreamWriter(zip.CreateEntry("xl/workbook.xml").Open())) writer.Write("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><workbookPr date1904=\"0\"/></workbook>");
                using (var writer = new StreamWriter(zip.CreateEntry("xl/worksheets/sheet1.xml").Open())) writer.Write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData><row r=\"18\"><c r=\"D18\" t=\"d\"><v>2026-10-29T12:00:00.000Z</v></c></row></sheetData></worksheet>");
                using (var writer = new StreamWriter(zip.CreateEntry("xl/vbaProject.bin").Open())) writer.Write("macro payload unchanged");
            }
            string hash = DataMigrationHost.Digest(original); File.Copy(original, copy);
            Check(MigrationPackageCompatibility.HasIsoDates(original) && MigrationPackageCompatibility.NormalizeOwnedCopy(copy) == 1 && !MigrationPackageCompatibility.HasIsoDates(copy), "ISO date is normalized only inside the owned copy");
            using (var stream = File.OpenRead(copy)) using (var zip = new ZipArchive(stream, ZipArchiveMode.Read))
            {
                var doc = new System.Xml.XmlDocument(); using (var xml = zip.GetEntry("xl/worksheets/sheet1.xml").Open()) doc.Load(xml);
                double actual = double.Parse(doc.GetElementsByTagName("v")[0].InnerText, System.Globalization.CultureInfo.InvariantCulture);
                Check(actual == new DateTime(2026, 10, 29, 12, 0, 0).ToOADate(), "normalized date keeps year and fractional time rather than year-only number");
                using (var reader = new StreamReader(zip.GetEntry("xl/vbaProject.bin").Open())) Check(reader.ReadToEnd() == "macro payload unchanged", "date normalization retains macro bytes");
            }
            Check(hash == DataMigrationHost.Digest(original), "ISO date compatibility does not modify the original package");
            Check(MigrationPackageCompatibility.DateSerial("1900-02-28", false) == 59 && MigrationPackageCompatibility.DateSerial("1900-03-01", false) == 61, "Excel 1900 leap-day boundary is preserved");
            Check(MigrationPackageCompatibility.DateSerial("1904-01-01", true) == 0, "1904 workbook epoch is supported");
        }
        private static void Drain(IEnumerable<MigrationProgress> steps) { foreach (var step in steps) { } }
        private static MigrationRequest Request(string source, params MigrationTarget[] targets)
        { return new MigrationRequest { SourcePath = source, SourceSheet = "源表", SourceRange = "B1:C2", Targets = targets.ToList() }; }
        private static void PickerChecks(string root)
        {
            string path = Path.Combine(root, "选择源.xlsx");
            using (var file = File.Create(path)) using (var zip = new ZipArchive(file, ZipArchiveMode.Create))
            using (var text = new StreamWriter(zip.CreateEntry("xl/workbook.xml").Open()))
                text.Write("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheets><sheet name=\"迁移源\" sheetId=\"1\"/><sheet name=\"辅助 &amp; 明细\" sheetId=\"2\"/></sheets></workbook>");
            var app = new FakeApp(); string hash = DataMigrationHost.Digest(path);
            using (var host = new DataMigrationHost(app))
                Check(host.Sheets(path).SequenceEqual(new[] { "迁移源", "辅助 & 明细" }), "xlsx picker reads Chinese and escaped sheet names without opening WPS workbooks");
            string macroPath = Path.Combine(root, "选择源.xlsm"); File.Copy(path, macroPath);
            using (var host = new DataMigrationHost(app)) Check(host.Sheets(macroPath).Count == 2, "xlsm picker also reads the package without a COM open");
            Check(app.Workbooks.OpenCalls == 0 && app.Workbooks.Count == 0, "choosing OOXML source creates no host workbooks or active-window switch");
            Check(hash == DataMigrationHost.Digest(path) && !Directory.GetFiles(root, ".xy-migration-read-*").Any(), "file picker neither changes original bytes nor creates read copies");
            using (var form = new DataMigrationForm())
            {
                Check(form.Mappings.Count == 0 && form.SourceBooks.Count == 0 && form.TargetBooks.Count == 0, "new migration window starts without an implicit source, target or range");
                Reject(delegate { form.BuildPlan(); }, "source and target ranges must be explicitly chosen before preflight");
                var pair = form.AddPair(new MigrationPair());
                form.AddBook(false, path, new[] { "迁移源", "辅助 & 明细" });
                Check(pair.Source.Books.Count == 0 && pair.Source.Sheet.Length == 0 && pair.Source.Range.Length == 0, "adding a workbook never fills existing ranges or guesses a sheet or an A1 range");
                Reject(delegate { form.BuildPlan(); }, "a row without sheets and ranges cannot be prechecked");
            }
            string sample = Path.Combine(root, "mouse-source.xls");
            var book = app.Workbooks.Create(sample, "源表"); var sheet = book.Sheets[0]; book.Saved = false;
            string original = DataMigrationHost.Digest(sample);
            Func<object, object> choose = s => new PickerSelection { Parent = s, Row = 5, Column = 2, Rows = new FakeCount { Count = 18 }, Columns = new FakeCount { Count = 8 } };
            using (var host = new DataMigrationHost(app))
            {
                Check(host.PickSourceRange(sample, "源表", "", choose) == "B5:I22", "mouse selection reads a rectangular range from the requested source");
                Check(host.PickSourceRange(sample, "源表", "B5:I22", s => false) == null, "native picker cancellation returns no replacement address");
                Check(host.PickSourceRange(sample, "源表", "invalid", choose) == "B5:I22", "an invalid previously typed range does not prevent picking again");
            }
            Check(!book.Saved && app.Workbooks.Items.Contains(book) && DataMigrationHost.Digest(sample) == original, "coordinate selection neither saves nor closes an unsaved borrowed workbook");
            var selected = (PickerSelection)choose(sheet);
            Reject(delegate { DataMigrationHost.SourceRangeAddress(selected, book, new FakeSheet(book, "其他表")); }, "picker rejects a different source sheet");
            Reject(delegate { DataMigrationHost.SourceRangeAddress(selected, new object(), sheet); }, "picker rejects a different source workbook");
            selected.Areas.Count = 2;
            Reject(delegate { DataMigrationHost.SourceRangeAddress(selected, book, sheet); }, "picker rejects multiple disjoint Ctrl selections");
            selected.Areas.Count = 1; selected.Rows.Count = 100001; selected.Columns.Count = 1;
            Reject(delegate { DataMigrationHost.SourceRangeAddress(selected, book, sheet); }, "picker enforces the migration cell limit");
            selected.Rows.Count = 1; selected.Row = 0;
            Reject(delegate { DataMigrationHost.SourceRangeAddress(selected, book, sheet); }, "picker rejects invalid coordinates");
            book.Close(false);
            using (var host = new DataMigrationHost(app))
                Check(host.PickSourceRange(sample, "源表", "", choose) == "B5:I22", "closed source range selection uses an owned read copy");
            Check(app.Workbooks.Count == 0 && !Directory.GetFiles(root, ".xy-migration-read-*").Any() && DataMigrationHost.Digest(sample) == original, "picker closes and removes only its own copy and preserves original bytes");
            using (var host = new DataMigrationHost(app)) Check(host.PickSourceRange(sample, "源表", "", s => false) == null, "cancelling a closed source picker is supported");
            Check(app.Workbooks.Count == 0 && !Directory.GetFiles(root, ".xy-migration-read-*").Any(), "cancelled source picker also cleans up its own copy");
        }
        private static IEnumerable<MigrationProgress> ReentrantStep(Action callback)
        { callback(); yield return new MigrationProgress("test step"); }
        private static void TargetSelectionChecks(string root)
        {
            string folder = Path.Combine(root, "批量目标"), sub = Path.Combine(folder, "子文件夹"); Directory.CreateDirectory(sub);
            foreach (string name in new[] { "甲.xls", "乙.XLSX", "丙.xlsm", "丁.et", "说明.txt", "~$锁文件.xlsx", ".xy-migration-read-test.xlsx", "甲_复制结果.xlsx", "甲_复制结果_2.xls", "甲_复制修正结果.xlsx", "源.xlsx" }) File.WriteAllText(Path.Combine(folder, name), "fixture");
            File.WriteAllText(Path.Combine(sub, "不递归.xlsx"), "fixture");
            var scan = MigrationTargetScan.Expand(new[] { folder, Path.Combine(folder, "甲.xls"), folder }, Path.Combine(folder, "源.xlsx"));
            Check(scan.Files.Count == 4 && scan.Errors.Count == 0, "folder selection discovers four supported workbook formats and deduplicates overlapping selections");
            Check(scan.Files.All(p => Path.GetDirectoryName(p) == folder) && !scan.Files.Any(p => p.Contains("复制结果") || p.Contains("复制修正结果") || Path.GetFileName(p).StartsWith("~$") || Path.GetFileName(p).StartsWith(".xy-") || Path.GetFileName(p) == "源.xlsx"), "folder scan excludes subfolders, source, locks, temporary files and generated results");
            Check(MigrationTargetScan.Expand(new[] { sub, Path.Combine(folder,"不存在") }, "").Files.Count == 1 && MigrationTargetScan.Expand(new[] { Path.Combine(folder,"不存在") }, "").Errors.Count == 1, "one bad selection does not discard files discovered elsewhere");
            Check(MigrationTargetScan.Expand(new[] { Path.Combine(folder,"甲_复制结果.xlsx") }, "").Files.Count == 1, "explicit file selection can still use an earlier result intentionally");
            var app = new FakeApp(); string sourcePath=Path.Combine(root,"check-source.xls"), first=Path.Combine(root,"check-first.xls"), second=Path.Combine(root,"check-second.xls");
            var source=app.Workbooks.Create(sourcePath,"源表"); source.Sheets[0].Cell("B1").Value2=88.0; source.Persist();
            app.Workbooks.Create(first,"目标表"); app.Workbooks.Create(second,"目标表"); string secondHash=DataMigrationHost.Digest(second);
            using(var form=new DataMigrationForm())
            {
                form.AddBook(false,sourcePath,new[]{"源表"}); form.AddBook(true,first,new[]{"目标表"}); form.AddBook(true,second,new[]{"目标表"});
                var pair=form.AddPair(new MigrationPair()); pair.Source.Books.Clear(); pair.Source.Books.Add(sourcePath); pair.Source.Sheet="源表"; pair.Source.Range="B1:C2";
                pair.Target.Books.Add(first); pair.Target.Books.Add(second); pair.Target.Sheet="目标表"; pair.Target.Range="B100"; form.RefreshResults();
                Check(form.Mappings.Count==1 && form.Mappings[0].Enabled && form.Mappings[0].SourceUsed && form.Mappings[0].TargetUsed, "a complete mapping row starts enabled with both sides");
                pair.Target.Books.Remove(second); form.RefreshResults();
                var plan=form.BuildPlan(); var request=plan.Passes.Single().Request;
                Check(request.Targets.Count==1 && request.Targets[0].Path==first, "workbooks not chosen for a row are absent from the executable plan");
                app.Workbooks.Find(first).Close(false);
                using(var host=new DataMigrationHost(app)){Drain(host.Preview(request));Drain(host.Execute());}
                Check(request.Targets[0].Status=="成功" && DataMigrationHost.Digest(second)==secondHash && !File.Exists(Path.Combine(root,"check-second_复制结果.xls")), "only the chosen target produces an output and excluded original bytes remain unchanged");
                int invalidated=0; form.InputsChanged+=delegate{invalidated++;};
                pair.Enabled=false; form.RefreshResults();
                Reject(delegate{form.BuildPlan();}, "a disabled row blocks preview rather than falling back to anything else");
                pair.Enabled=true; pair.Target.Books.Add(second); form.RefreshResults();
                form.AddBook(true,Path.Combine(root,"unreadable.xlsx"),new string[0],"fixture read failure");
                Check(invalidated==1 && pair.Target.Books.Count==2, "adding a workbook invalidates the preview once and never changes existing rows");
                pair.Target.Books.Add(Path.GetFullPath(Path.Combine(root,"unreadable.xlsx"))); form.RefreshResults();
                plan=form.BuildPlan(); var unreadable=plan.Mappings.Single(m=>m.TargetPath.EndsWith("unreadable.xlsx"));
                Check(unreadable.Result.Status=="失败" && unreadable.Result.Error.Contains("fixture read failure"), "an unreadable target stays listed and fails with its reason");
                app.Workbooks.Find(second).Close(false);
                request=plan.Passes.Single().Request;
                using(var host=new DataMigrationHost(app))Drain(host.Preview(request));
                Check(request.Targets.Count==2 && request.Targets.All(t=>t.Ready), "the other targets of the row remain ready while one target fails independently");
            }
        }
        private static void SelectionPopupChecks(string root)
        {
            using (var form = new DataMigrationForm())
            {
                form.AddBook(true, Path.Combine(root, "目标甲.xlsx"), new[] { "目标表" });
                form.AddBook(true, Path.Combine(root, "目标乙.xlsx"), new[] { "目标表" });
                form.AddBook(true, Path.Combine(root, "子目录", "目标乙.xlsx"), new[] { "辅助明细" });
                var pair = form.AddPair(new MigrationPair()); pair.Target.Books.AddRange(form.TargetBooks.Select(b => b.Path)); pair.Target.Books.RemoveAt(1); form.RefreshResults();
                var row = form.PairItems()[0];
                Check(!DialogFormatting.Walk(form).OfType<System.Windows.Forms.Button>().Any(b => b.Text == "全选" || b.Text == "全不选"), "standalone all/none buttons are removed from the main window");
                int changed = 0; form.InputsChanged += delegate { changed++; };
                using (var menu = form.CreateBookPopup(row, true))
                {
                    menu.Search.Text = "甲"; Check(menu.MatchCount == 1, "workbook search matches Chinese workbook names");
                    menu.SelectMatching(false);
                    Check(pair.Target.Books.Count == 2 && changed == 0, "draft selections do not change the row before confirmation");
                }
                Check(pair.Target.Books.Count == 2 && changed == 0, "cancel or dismiss discards the entire selection draft");
                using (var menu = form.CreateBookPopup(row, true))
                {
                    menu.Search.Text = "乙"; Check(menu.MatchCount == 2, "equal filenames from different directories remain separate choices");
                    menu.SelectMatching(false); menu.Search.Text = "甲";
                    Check(menu.List.GetItemChecked(0), "batch changes to search matches preserve unmatched selections");
                    menu.Search.Text = "子目录"; Check(menu.MatchCount == 1 && !menu.List.GetItemChecked(0), "full-path search finds a specific duplicate and retains draft changes");
                    menu.InvertMatching(); menu.Search.Text = "辅助明细";
                    Check(menu.MatchCount == 1 && menu.List.GetItemChecked(0), "invert changes only matching rows and worksheet names are searchable");
                    menu.Search.Text = "乙"; menu.SelectMatching(false); menu.ApplySelection();
                }
                Check(changed == 1 && pair.Target.Books.Count == 1 && pair.Target.Books[0].EndsWith("目标甲.xlsx"), "confirm applies the full draft once and only checked workbooks enter the row");
                Check(form.Mappings.Count == 1 && form.PairItems()[0] == row, "workbook search does not hide or reorder the table rows");
                using (var menu = form.CreateBookPopup(row, true)) { menu.ApplySelection(); }
                Check(changed == 1, "confirming unchanged choices preserves the existing preview identity");
                using (var menu = form.CreateBookPopup(row, true))
                {
                    menu.SelectMatching(false); form.Run(Enumerable.Repeat(new MigrationProgress("fixture"), 2), delegate { }, delegate(Exception error) { throw error; });
                    Reject(delegate { form.CreateBookPopup(row, true); }, "running migration prevents opening the workbook dropdown");
                    menu.ApplySelection();
                    Check(pair.Target.Books.Count == 1 && changed == 1, "a stale dropdown cannot change a row during execution");
                }
            }
        }
        private static void PumpChecks()
        {
            using (var form = new DataMigrationForm())
            {
                bool completed = false, premature = false; Exception failure = null;
                var tick = typeof(DataMigrationForm).GetMethod("Tick", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                form.Run(ReentrantStep(delegate { tick.Invoke(form, new object[0]); premature = completed; }), delegate { completed = true; }, delegate(Exception error) { failure = error; });
                tick.Invoke(form, new object[0]);
                Check(!premature && completed && failure == null, "COM message pumping cannot reenter MoveNext and dispose a running migration session");
            }
        }
        public sealed class EventSheet { public object Parent { get; set; } }
        private static void WatchChecks()
        {
            object book=new object(); var sheet=new EventSheet{Parent=book};
            var notify=typeof(SourceChangeWatch).GetMethod("OnChanged",System.Reflection.BindingFlags.NonPublic|System.Reflection.BindingFlags.Instance);
            using(var watch=new SourceChangeWatch(new object())) {watch.WatchBook(book);notify.Invoke(watch,new object[]{sheet,false,"SheetCalculate"});Reject(delegate{watch.Validate();},"legacy monitor reproduces invalidation on calculation without data editing");}
            using(var watch=new SourceChangeWatch(new object(),true)) {watch.WatchBook(book);notify.Invoke(watch,new object[]{sheet,false,"SheetCalculate"});watch.Validate();Check(true,"migration monitor ignores calculation-only notifications");notify.Invoke(watch,new object[]{sheet,false,"SheetChange"});bool precise=false;try{watch.Validate();}catch(InvalidOperationException error){precise=error.Message.Contains("SheetChange");}Check(precise,"real source edits still abort with the precise event type");}
            using(var watch=new SourceChangeWatch(new object(),true))
            {
                watch.WatchBook(book);var other=new object();var worker=new System.Threading.Thread(delegate(){notify.Invoke(watch,new object[]{other,true,"WorkbookBeforeClose"});});worker.Start();worker.Join();watch.Validate();Check(true,"foreign-thread unrelated close is checked on the host rather than aborting all books");
                worker=new System.Threading.Thread(delegate(){notify.Invoke(watch,new object[]{book,true,"WorkbookBeforeClose"});});worker.Start();worker.Join();bool precise=false;try{watch.Validate();}catch(InvalidOperationException error){precise=error.Message.Contains("WorkbookBeforeClose");}Check(precise,"foreign-thread close of the monitored book still aborts");
            }
        }
        private static void FidelityChecks(string root)
        {
            string source=Path.Combine(root,"fidelity-source.xlsx"),target=Path.Combine(root,"fidelity-output.xlsx");
            using(var file=File.Create(source))using(var zip=new ZipArchive(file,ZipArchiveMode.Create))
            {
                var parts=new Dictionary<string,string>{
                    {"xl/workbook.xml","<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"源表\" sheetId=\"1\" r:id=\"r1\"/></sheets></workbook>"},
                    {"xl/_rels/workbook.xml.rels","<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"r1\" Target=\"worksheets/sheet1.xml\"/></Relationships>"},
                    {"xl/styles.xml","<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><numFmts count=\"1\"><numFmt numFmtId=\"164\" formatCode=\"yyyy-mm-dd\"/></numFmts><cellXfs count=\"2\"><xf numFmtId=\"0\"/><xf numFmtId=\"164\"/></cellXfs></styleSheet>"},
                    {"xl/sharedStrings.xml","<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><si><t>hidden text</t></si></sst>"},
                    {"xl/worksheets/sheet1.xml","<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData><row r=\"1\"><c r=\"A1\" t=\"str\"><v>anchor</v></c><c r=\"B1\" s=\"1\"><v>46324</v></c></row><row r=\"2\"><c r=\"B2\" t=\"s\"><v>0</v></c></row></sheetData><mergeCells><mergeCell ref=\"A1:B2\"/></mergeCells></worksheet>"}
                };
                foreach(var part in parts)using(var text=new StreamWriter(zip.CreateEntry(part.Key).Open()))text.Write(part.Value);
            }
            string hash=DataMigrationHost.Digest(source);File.Copy(source,target);
            using(var file=File.Open(target,FileMode.Open,FileAccess.ReadWrite))using(var zip=new ZipArchive(file,ZipArchiveMode.Update))
            {
                foreach(string name in new[]{"xl/styles.xml","xl/worksheets/sheet1.xml"})
                {
                    var entry=zip.GetEntry(name);string text;using(var reader=new StreamReader(entry.Open()))text=reader.ReadToEnd();
                    text=text.Replace("yyyy-mm-dd","yyyy/mm/dd").Replace("<c r=\"B2\" t=\"s\"><v>0</v></c>","");
                    using(var stream=entry.Open()){stream.SetLength(0);using(var writer=new StreamWriter(stream))writer.Write(text);}
                }
            }
            var snapshot=MigrationPackagePreservation.Capture(source,"源表",MigrationArea.Parse("A1:B2"));var destination=new MigrationTarget{Sheet="源表",Area=MigrationArea.Parse("A1:B2")};
            bool rejected=false;try{snapshot.Verify(target,destination);}catch(IOException){rejected=true;}Check(rejected,"package-level audit rejects the date-format and merged-hidden-data differences");
            string contentOnly = Path.Combine(root, "content-only.xlsx"); File.Copy(target, contentOnly);
            var contentSnapshot = MigrationPackagePreservation.Capture(source, "源表", MigrationArea.Parse("A1:B2"), false);
            contentSnapshot.Apply(contentOnly, destination); contentSnapshot.Verify(contentOnly, destination);
            using (var file = File.OpenRead(contentOnly)) using (var zip = new ZipArchive(file, ZipArchiveMode.Read))
            {
                var styles = new System.Xml.XmlDocument(); using (var xml = zip.GetEntry("xl/styles.xml").Open()) styles.Load(xml);
                Check(styles.GetElementsByTagName("numFmt")[0].Attributes["formatCode"].Value == "yyyy/mm/dd", "content-only package preservation retains the target number format instead of restoring the source format");
                var sheet = new System.Xml.XmlDocument(); using (var xml = zip.GetEntry("xl/worksheets/sheet1.xml").Open()) sheet.Load(xml);
                Check(contentSnapshot.HasHiddenValues && sheet.GetElementsByTagName("c").Cast<System.Xml.XmlElement>().Any(c => c.GetAttribute("r") == "B2" && c.InnerText == "hidden text"), "content-only package preservation still restores merged hidden values");
            }
            snapshot.Apply(target,destination);snapshot.Verify(target,destination);Check(snapshot.HasHiddenValues,"raw hidden merged values and original number-format code are restored and verified");Check(hash==DataMigrationHost.Digest(source),"fidelity restoration never edits the source file");
        }
        private static void TransactionChecks(string root)
        {
            string folder = Path.Combine(root, "原文件更新"); Directory.CreateDirectory(folder);
            string target = Path.Combine(folder, "审计底稿.xlsm"), stage = Path.Combine(folder, ".xy-migration-stage.xlsm");
            File.WriteAllText(target, "original"); File.WriteAllText(stage, "verified");
            string hash = DataMigrationHost.Digest(target);
            string backup = MigrationOutputTransaction.Commit(stage, target, hash, "batch1");
            Check(File.ReadAllText(target) == "verified" && File.ReadAllText(backup) == "original" && !File.Exists(stage), "atomic replacement retains exact original backup and original Unicode filename");
            Check(Path.GetDirectoryName(backup) == Path.Combine(folder, "复制备份", "batch1") && !Directory.GetFiles(folder, "*_复制结果*").Any(), "replacement produces no result-suffix workbook in the target directory");
            File.WriteAllText(stage, "next"); bool failed = false;
            try { MigrationOutputTransaction.Commit(stage, target, hash, "stale"); } catch (IOException) { failed = true; }
            Check(failed && File.ReadAllText(target) == "verified" && File.Exists(stage), "changed original is rejected without consuming the verified staging file");
            hash = DataMigrationHost.Digest(target); failed = false;
            using (File.Open(target, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                try { MigrationOutputTransaction.Commit(stage, target, hash, "locked"); } catch (IOException) { failed = true; }
            Check(failed && File.ReadAllText(target) == "verified", "locked target fails before replacement");
            failed = false; try { MigrationOutputTransaction.Commit(stage, target, hash, "batch1"); } catch (IOException) { failed = true; }
            Check(failed && File.ReadAllText(backup) == "original" && File.ReadAllText(target) == "verified", "existing recovery file is never overwritten");
            string nextBackup = MigrationOutputTransaction.Commit(stage, target, hash, "batch2");
            Check(File.ReadAllText(target) == "next" && File.ReadAllText(nextBackup) == "verified" && File.ReadAllText(backup) == "original", "repeated migration preserves each preceding version independently");
        }
        private static void QuietReplaceChecks(string root)
        {
            string folder = Path.Combine(root, "逐项静默"); Directory.CreateDirectory(folder);
            var app = new FakeApp(); var source = app.Workbooks.Create(Path.Combine(folder, "源.xls"), "源表");
            source.Sheets[0].Cell("B1").Formula = "=A1"; source.Sheets[0].Cell("C1").Value2 = "=text"; source.Persist();
            app.ActiveWorkbook = source; app.ActiveSheet = source.Sheets[0]; string sourceHash = DataMigrationHost.Digest(source.FullName);
            var targets = new List<MigrationTarget>(); var hashes = new Dictionary<string, string>();
            for (int i = 0; i < 12; i++)
            {
                var book = app.Workbooks.Create(Path.Combine(folder, "目标" + i + ".xls"), "目标表");
                book.Sheets[0].Cell("B100").Value2 = "old"; book.FailSave = i == 1; book.Persist(); book.Close(false);
                targets.Add(new MigrationTarget { Path = book.FullName, Sheet = "目标表", Start = "B100" }); hashes[book.FullName] = DataMigrationHost.Digest(book.FullName);
            }
            var request = Request(source.FullName, targets.ToArray()); request.ReplaceTargets = true;
            bool bounded = true, hidden = true; int peak = 0;
            Action inspect = delegate
            {
                peak = Math.Max(peak, app.Workbooks.Count); bounded &= app.Workbooks.Count <= 2;
                hidden &= app.Workbooks.Items.Where(b => Path.GetFileName(b.FullName).StartsWith(".xy-migration-")).All(b => !b.Windows.Item(1).Visible);
            };
            using (var host = new DataMigrationHost(app))
            {
                foreach (var step in host.Preview(request)) inspect();
                Check(targets.All(t => t.Ready && t.Handle == null && t.Gateway == null) && app.Workbooks.Count == 1, "preflight closes each target reader before retaining preview data");
                foreach (var step in host.Execute()) inspect();
                Check(targets.Count(t => t.Status == "成功") == 11 && targets[1].Status == "失败", "replacement save failure is independent across a twelve-target batch");
            }
            Check(bounded && hidden && peak == 2 && app.Workbooks.Count == 1, "temporary workbooks remain hidden at every yield and do not accumulate with batch size");
            Check(source.Windows.Item(1).Visible && DataMigrationHost.Digest(source.FullName) == sourceHash, "user source visibility and bytes are unchanged");
            Check(targets[1].Output.Length == 0 && targets[1].Backup.Length == 0 && DataMigrationHost.Digest(targets[1].Path) == hashes[targets[1].Path], "failed save leaves the target original unchanged");
            Check(targets.Where(t => t.Status == "成功").All(t => t.Output == t.Path && File.Exists(t.Backup) && DataMigrationHost.Digest(t.Backup) == hashes[t.Path]), "every successful original update reports its exact recovery path");
            var updated = app.Workbooks.Open(targets[0].Path, 0, true);
            Check(updated.Sheets[0].Cell("B100").Formula == "=A100" && updated.Windows.Item(1).Visible, "updated original reopens at a visible workbook view with migrated formula");
            updated.Close(false);
            Check(!Directory.GetFiles(folder, "*_复制结果*").Any() && !Directory.GetFiles(folder, ".xy-migration-*").Any(), "successful and failed processing leave no result-suffix or reader files in the working directory");
            var open = app.Workbooks.Open(targets[0].Path, 0, false); var blocked = new MigrationTarget { Path = targets[0].Path, Sheet = "目标表", Start = "B100" };
            request = Request(source.FullName, blocked); request.ReplaceTargets = true;
            string openedHash = DataMigrationHost.Digest(blocked.Path);
            using (var host = new DataMigrationHost(app))
            {
                Drain(host.Preview(request));
                Check(blocked.Ready && app.Workbooks.Items.Contains(open) && DataMigrationHost.Digest(blocked.Path) == openedHash, "preflight reads a saved open target without closing or changing it");
                Drain(host.Execute());
                Check(!blocked.Ready && blocked.Error.Contains("关闭") && blocked.Output.Length == 0 && blocked.Backup.Length == 0 && app.Workbooks.Items.Contains(open) && DataMigrationHost.Digest(blocked.Path) == openedHash, "execution rejects an open replacement target with save-and-close guidance and leaves the original unchanged");
            }
            open.Close(false);
            var cancelled = new MigrationTarget { Path = targets[2].Path, Sheet = "目标表", Start = "B100" };
            request = Request(source.FullName, cancelled); request.ReplaceTargets = true; string before = DataMigrationHost.Digest(cancelled.Path);
            using (var host = new DataMigrationHost(app))
            {
                Drain(host.Preview(request)); using (var steps = host.Execute().GetEnumerator())
                { Check(steps.MoveNext(), "replacement execution yields before committing"); host.Cancelled = true; while (steps.MoveNext()) { } }
                Check(cancelled.Status == "已取消" && cancelled.Output.Length == 0 && DataMigrationHost.Digest(cancelled.Path) == before, "cancellation before commit preserves the original file");
            }
            using (var form = new DataMigrationForm())
            {
                var option = (System.Windows.Forms.CheckBox)form.Controls.Find("migrationReplaceTargets", true).Single();
                form.AddBook(false, source.FullName, new[] { "源表" }); form.AddBook(true, targets[0].Path, new[] { "目标表" });
                var pair = form.AddPair(new MigrationPair()); pair.Source.Books.Add(Path.GetFullPath(source.FullName)); pair.Source.Sheet = "源表"; pair.Source.Range = "B1:C2";
                pair.Target.Books.Add(Path.GetFullPath(targets[0].Path)); pair.Target.Sheet = "目标表"; pair.Target.Range = "B100"; form.RefreshResults();
                Check(option.Checked && form.BuildPlan().ReplaceTargets, "production window defaults to updating originals with backup");
                option.Checked = false; Check(!form.BuildPlan().ReplaceTargets, "user can explicitly retain the earlier save-copy behavior");
            }
        }
        private static void SnapshotRaceChecks(string root)
        {
            string folder = Path.Combine(root, "读取快照竞争"); Directory.CreateDirectory(folder);
            var app = new FakeApp(); var source = app.Workbooks.Create(Path.Combine(folder, "源.xls"), "源表"); source.Sheets[0].Cell("B1").Value2 = 9.0; source.Persist();
            app.ActiveWorkbook = source; app.ActiveSheet = source.Sheets[0];
            var book = app.Workbooks.Create(Path.Combine(folder, "目标.xls"), "目标表"); book.Close(false);
            app.Workbooks.Opened = delegate(string path) { if (Path.GetFileName(path).StartsWith(".xy-migration-read-")) File.WriteAllText(book.FullName, "concurrent target edit"); };
            var target = new MigrationTarget { Path = book.FullName, Sheet = "目标表", Start = "B100" }; var request = Request(source.FullName, target); request.ReplaceTargets = true;
            using (var host = new DataMigrationHost(app)) { Drain(host.Preview(request)); Check(!target.Ready && target.Error.Contains("变化") && File.ReadAllText(book.FullName) == "concurrent target edit", "target edits during read-copy opening are rejected instead of adopting a new fingerprint"); }
            app.Workbooks.Opened = null; source.Close(false);
            app.Workbooks.Opened = delegate(string path) { if (Path.GetFileName(path).StartsWith(".xy-migration-read-")) File.WriteAllText(source.FullName, "concurrent source edit"); };
            using (var host = new DataMigrationHost(app)) Reject(delegate { Drain(host.Preview(Request(source.FullName, new MigrationTarget { Path = book.FullName, Sheet = "目标表", Start = "B100" }))); }, "source edits during read-copy opening cannot be accepted as the source snapshot");
            app.Workbooks.Opened = null;
        }
        private static void HostChecks(string root)
        {
            var app = new FakeApp(); var source = app.Workbooks.Create(Path.Combine(root, "源.xlsx"), "源表");
            source.Sheets[0].Cell("B1").Formula = "=A1+'[第三方.xlsx]明细'!$A$1+Rate";
            source.Sheets[0].Cell("C1").Value2 = "=普通文本";
            source.Sheets[0].Cell("B2").Value2 = 12.5;
            source.Names.Add("Rate", "='源表'!$A$1", true);
            source.Links.Add(Path.Combine(root, "第三方.xlsx")); source.Persist();
            app.ActiveWorkbook = source; app.ActiveSheet = source.Sheets[0];
            var targets = new List<MigrationTarget>(); var originals = new Dictionary<string, string>();
            foreach (string ext in new[] { ".xlsx", ".xlsm", ".xls", ".et" })
            {
                var book = app.Workbooks.Create(Path.Combine(root, "目标" + ext), "目标表"); book.Sheets[0].Cell("B100").Value2 = 77.0; book.Persist();
                targets.Add(new MigrationTarget { Path = book.FullName, Sheet = "目标表", Start = "B100" }); originals[book.FullName] = DataMigrationHost.Digest(book.FullName);
            }
            var conflict = app.Workbooks.Create(Path.Combine(root, "冲突.xlsx"), "目标表"); conflict.Names.Add("Rate", "=99", true); conflict.Persist();
            targets.Insert(1, new MigrationTarget { Path = conflict.FullName, Sheet = "目标表", Start = "B100" });
            var saveFailure = app.Workbooks.Create(Path.Combine(root, "保存失败.xlsm"), "目标表"); saveFailure.FailSave = true; saveFailure.Persist();
            var badSaveTarget = new MigrationTarget { Path = saveFailure.FullName, Sheet = "目标表", Start = "B100" }; targets.Insert(2, badSaveTarget);
            originals[saveFailure.FullName] = DataMigrationHost.Digest(saveFailure.FullName);
            originals[source.FullName] = DataMigrationHost.Digest(source.FullName); originals[conflict.FullName] = DataMigrationHost.Digest(conflict.FullName);
            var emptyRequest = new MigrationRequest { SourcePath = source.FullName, SourceSheet = "源表", SourceRange = "Z99", Targets = new List<MigrationTarget> { targets[0] } };
            using (var host = new DataMigrationHost(app)) Reject(delegate { Drain(host.Preview(emptyRequest)); }, "accidental all-blank source cannot silently produce a successful empty result");
            emptyRequest.AllowEmptySource = true;
            using (var host = new DataMigrationHost(app)) { Drain(host.Preview(emptyRequest)); Check(targets[0].Ready && emptyRequest.NonEmptyCells == 0, "intentional all-blank clearing requires explicit opt-in"); }
            Check(DataMigrationForm.Summary(new MigrationRequest { SourceSheet = "源表", SourceRange = "B1:C2", NonEmptyCells = 3, Targets = new List<MigrationTarget> { new MigrationTarget { Status = "成功" }, new MigrationTarget { Status = "失败" } } }).Contains("成功 1，失败 1"), "partial batch failure is visible in the completion summary");
            using (var host = new DataMigrationHost(app))
            {
                Drain(host.Preview(Request(source.FullName, targets.ToArray())));
                Check(targets.Count(t => t.Ready) == 5 && targets[1].Status == "失败", "one name conflict does not block other targets");
                Check(targets[0].Overwritten == 1, "preflight counts overwritten content");
                Drain(host.Execute());
                Check(targets.Count(t => t.Status == "成功") == 4, "real adapter publishes four independent managed fixture copies");
                Check(badSaveTarget.Status == "失败" && badSaveTarget.Output.Length == 0, "save failure never publishes an unverified result and later targets continue");
                Check(!Directory.GetFiles(root, ".xy-migration-*.xlsm").Any() && Directory.GetFiles(host.DiagnosticDirectory, "*_未通过校验_*.xlsm").Any(), "failed copies use readable diagnostic names outside the user output directory");
                foreach (var target in targets.Where(t => t.Output.Length > 0))
                {
                    var output = app.Workbooks.Open(target.Output, 0, true); var sheet = output.Sheets[0];
                    Check(sheet.Cell("B100").Formula.StartsWith("=A100+") && sheet.Cell("B100").Formula.Contains("!$A$1"), "internal shift and frozen external address survive reopen");
                    Check(!sheet.Cell("C100").HasFormula && (string)sheet.Cell("C100").Value2 == "=普通文本", "leading equals remains literal text");
                    Check(sheet.Cell("C101").Value2 == null && Convert.ToDouble(sheet.Cell("B101").Value2) == 12.5, "blank and numeric values survive reopen"); output.Close(false);
                }
            }
            foreach (var original in originals) Check(DataMigrationHost.Digest(original.Key) == original.Value, "original file bytes unchanged: " + Path.GetFileName(original.Key));
            Check(app.EnableEvents && app.ScreenUpdating && app.DisplayAlerts && app.AskToUpdateLinks && app.AutomationSecurity == 1 && app.Calculation == -4105, "host event, alert, macro and calculation settings restored");
            var stale = new MigrationTarget { Path = targets[0].Path, Sheet = "目标表", Start = "B100" };
            using (var host = new DataMigrationHost(app))
            {
                Drain(host.Preview(Request(source.FullName, stale))); app.Workbooks.Find(stale.Path).Saved = false; Drain(host.Execute());
                Check(stale.Status == "失败" && stale.Output.Length == 0, "new unsaved edits prevent delivery"); app.Workbooks.Find(stale.Path).Saved = true;
            }
            // One target, two source ranges: read once, written into one copy, saved once.
            using (var host = new DataMigrationHost(app) { HoldTargets = true })
            {
                var left = new MigrationTarget { Path = targets[0].Path, Sheet = "目标表", Start = "B100" };
                var right = new MigrationTarget { Path = targets[0].Path, Sheet = "目标表", Start = "E100" };
                Drain(host.Preview(Request(source.FullName, left))); host.Keep();
                Drain(host.Preview(Request(source.FullName, right))); host.Keep();
                Check(host.KeptTargets == 2, "batched passes keep what they read");
                Drain(host.ExecuteKept());
                Check(left.Status == "成功" && right.Status == "成功" && left.Output.Length > 0 && left.Output == right.Output, "two source ranges for one target are delivered as one result");
                var output = app.Workbooks.Open(left.Output, 0, true); var sheet = output.Sheets[0];
                Check(Convert.ToDouble(sheet.Cell("B101").Value2) == 12.5 && Convert.ToDouble(sheet.Cell("E101").Value2) == 12.5, "both ranges survive the single save and reopen"); output.Close(false);
                File.Delete(left.Output);
            }
            Check(!Directory.GetFiles(root, ".xy-migration-read-*").Any() && app.Workbooks.Items.All(b => !Path.GetFileName(b.FullName).StartsWith(".xy-migration-")), "held target read copies are closed and deleted after a batched run");
            using (var host = new DataMigrationHost(app))
            {
                var cancelled = new MigrationTarget { Path = targets[0].Path, Sheet = "目标表", Start = "B100" }; Drain(host.Preview(Request(source.FullName, cancelled)));
                var steps = host.Execute().GetEnumerator(); Check(steps.MoveNext(), "execution yields between cells"); host.Cancelled = true;
                while (steps.MoveNext()) { } steps.Dispose(); Check(cancelled.Output.Length == 0 && cancelled.Status == "已取消", "cancelled copy is not published");
            }
            Check(app.Workbooks.Items.All(b => !Path.GetFileName(b.FullName).StartsWith(".xy-migration-")), "temporary copy workbooks closed after cancellation");
            var blocked = new MigrationTarget { Path = targets[0].Path, Sheet = "目标表", Start = "B100" };
            app.Workbooks.Find(blocked.Path).Sheets[0].ProtectContents = true;
            using (var host = new DataMigrationHost(app)) { Drain(host.Preview(Request(source.FullName, blocked))); Check(!blocked.Ready, "protected target rejected before writing"); }
            app.Workbooks.Find(blocked.Path).Sheets[0].ProtectContents = false;
            blocked = new MigrationTarget { Path = targets[0].Path, Sheet = "缺表", Start = "B100" };
            using (var host = new DataMigrationHost(app)) { Drain(host.Preview(Request(source.FullName, blocked))); Check(!blocked.Ready, "missing destination sheet fails independently"); }
            source.Saved = false;
            using (var host = new DataMigrationHost(app)) Reject(delegate { Drain(host.Preview(Request(source.FullName, targets[0]))); }, "unsaved user-owned source still blocks preflight");
            source.Saved = true;
            // Reproduce binary-reader metadata writes without opening any original.
            source.Close(false); app.Workbooks.Find(targets[0].Path).Close(false);
            app.Workbooks.DirtyReadCopies = true;
            var copyOnly = new MigrationTarget { Path = targets[0].Path, Sheet = "目标表", Start = "B100" };
            using (var host = new DataMigrationHost(app))
            {
                Drain(host.Preview(Request(source.FullName, copyOnly)));
                Check(app.Workbooks.Find(source.FullName) == null && app.Workbooks.Find(copyOnly.Path) == null, "unopened originals are inspected through disposable read copies");
                Check(app.Workbooks.Items.Where(b => Path.GetFileName(b.FullName).StartsWith(".xy-migration-read-")).All(b => !b.Saved), "fixture reproduces WPS marking readonly copies unsaved on open");
                Drain(host.Execute()); Check(copyOnly.Status == "成功", "read-copy snapshot supports execution");
            }
            var legacy = app.Workbooks.Create(Path.Combine(root, "选择旧格式.xls"), "旧格式源表"); legacy.Close(false);
            using (var host = new DataMigrationHost(app)) Check(host.Sheets(legacy.FullName).SequenceEqual(new[] { "旧格式源表" }), "legacy picker tolerates owned readonly import recalculation");
            app.Workbooks.DirtyReadCopies = false;
            Check(!Directory.GetFiles(root, ".xy-migration-read-*").Any(), "owned read copies removed after closing their workbooks");
            Check(originals.All(p => DataMigrationHost.Digest(p.Key) == p.Value), "all original bytes remain unchanged after closing snapshot readers");
            string macro = Path.Combine(root, "macro-fixture.xlsm");
            using (var file = File.Create(macro)) using (var zip = new ZipArchive(file, ZipArchiveMode.Create)) using (var writer = new StreamWriter(zip.CreateEntry("xl/vbaProject.bin").Open())) writer.Write("fixture VBA bytes");
            Check(DataMigrationHost.MacroFingerprints(macro).Count == 1, "VBA payload fingerprint captured independently of workbook edits");
        }
        [STAThread] public static int Main(string[] args)
        {
            string root = Path.Combine(Path.GetTempPath(), "migration-probe-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(root);
            DataMigrationHost.DiagnosticRootOverride = Path.Combine(root, "diagnostics");
            try
            {
                if(args.Length==1 && args[0]=="--target-picker-ui") { Console.WriteLine("PICKER_BEGIN"); HostDialogService.Register(); var picked=MigrationTargetPicker.Show(null); foreach(string path in picked)Console.WriteLine("PICKED: "+path); Console.WriteLine("PICKER_END count="+picked.Length+"\n"+UiDiagnostics.Text); Directory.Delete(root,true); return 0; }
                MappingSummaryChecks(); FormulaChecks(); PackageChecks(root); FidelityChecks(root); WatchChecks(); PickerChecks(root); TargetSelectionChecks(root); SelectionPopupChecks(root); PumpChecks(); HostChecks(root); TransactionChecks(root); QuietReplaceChecks(root); SnapshotRaceChecks(root); LinkChecks(root); LinkConfigurationChecks();
                using (var form = new DataMigrationForm()) Check(!form.Visible && form.Mappings.Count == 0 && form.Sources.Columns.Count == 7 && form.Targets.Columns.Count == 6 && form.Sources.Columns.Contains("pick") && form.Sources.Columns.Contains("trigger") && !form.Targets.Columns.Contains("trigger"), "production form constructs without host access");
                LayoutChecks(); RecalculationTargetChecks(root); PresetChecks(); ArchitectureChecks();
                if (args.Length == 2 && args[0] == "--render") Render(root, Path.GetFullPath(args[1]));
                Console.WriteLine("PASS: DataMigrationProbe " + count + " assertions; managed fixtures only, no WPS connection");
                Directory.Delete(root, true); return 0;
            }
            catch (Exception error) { Console.WriteLine("FAILED: " + error + "\nFixtures retained: " + root); return 1; }
            finally { DataMigrationHost.DiagnosticRootOverride = null; }
        }
        // Three pages, no top toolbar, a collapsible range panel, and
        // re-pairing by swapping one side in 映射.
        private static void LayoutChecks()
        {
            using (var form = new DataMigrationForm())
            {
                var all = DialogFormatting.Walk(form).ToList();
                var tabs = all.OfType<MigrationTabStrip>().Single(c => c.Name == "migrationTabs");
                Check(tabs.TabCount == 3 && tabs.Title(0) == "源区域" && tabs.Title(1) == "目标区域" && tabs.Title(2) == "映射", "source, target and mapping are separate pages");
                Check(!all.OfType<System.Windows.Forms.TabControl>().Any() && !all.OfType<System.Windows.Forms.LinkLabel>().Any(), "pages and add actions use the window's own controls");
                Check(!all.Any(c => c.Text == "收起区域选择" || c.Text == "新增一行" || c.Text == "复制此行" || c.Text == "删除行"), "the old top toolbar actions are removed");
                // 添加工作簿 now sits at the bottom of each range page, left of 新增区域.
                foreach (string pair in new[] { "添加源工作簿…|＋ 新增源区域", "添加目标工作簿…|＋ 新增目标区域" })
                {
                    var names = pair.Split('|'); var add = all.Single(c => c.Text == names[0]); var range = all.Single(c => c.Text == names[1]);
                    Check(add.Parent == range.Parent && add.Parent.Controls.GetChildIndex(add) < add.Parent.Controls.GetChildIndex(range) && !(add.Parent.Parent is System.Windows.Forms.FlowLayoutPanel && add.Parent.Parent.Name == "migrationToolbar"),
                        "“" + names[0] + "” sits next to “" + names[1] + "” at the bottom of its page, not in a top toolbar");
                }
                Check(all.Count(c => c.Text == "＋ 新增映射") == 1, "the mapping page has its own 新增映射 button");
                var split = all.OfType<System.Windows.Forms.SplitContainer>().Single(c => c.Name == "migrationSplit");
                Check(all.Any(c => c.Name == "migrationSidebarToggle") && all.Any(c => c.Name == "migrationSidebarClose"), "range panel has an edge toggle and a close button");
                Check(split.Panel2Collapsed, "range panel starts closed");
                form.ShowSidebar(true); form.ShowSidebar(false); Check(split.Panel2Collapsed, "range panel collapses so the pages fill the width");
                form.ShowSidebar(true); Check(!split.Panel2Collapsed, "range panel reopens from the edge toggle");
                var s1 = new MigrationSide { Sheet = "表", Range = "A1:B2" }; s1.Books.Add(@"D:\a.xlsx");
                var s2 = new MigrationSide { Sheet = "表", Range = "C1:D2" }; s2.Books.Add(@"D:\a.xlsx");
                var t1 = new MigrationSide { Sheet = "表", Range = "E1" }; t1.Books.Add(@"D:\b.xlsx");
                var p1 = form.AddSide(false, s1); var p2 = form.AddSide(false, s2); var p3 = form.AddSide(true, t1);
                Check(p1.ValuesOnly && p2.ValuesOnly, "rows created in the window paste values by default");
                Check(!p1.KeepFormat && MigrationPair.ModeText(true, false) == "只写数值" && MigrationPair.ModeText(true, true) == "数值+格式" && MigrationPair.ModeText(false, false) == "只写公式" && MigrationPair.ModeText(false, true) == "公式+格式", "new rows write values without formats; four write modes");
                var board = form.Mappings;
                Check(board.BeginModeSelection(0) && !board.ModeMenuOpen
                    && board.ModeChoices.SequenceEqual(new[] { "只写数值", "数值+格式", "只写公式", "公式+格式", "只写格式", "粘贴链接" }), "mapping dropdown has all six ordered modes");
                board.PendingMode = "粘贴链接"; board.CloseModeSelection();
                Check(p1.Mode == "只写数值", "cancelling a pending dropdown selection preserves the row");
                board.BeginModeSelection(0); board.PendingMode = "粘贴链接"; board.CloseModeSelection(); CommitMode(board);
                Check(p1.Mode == "只写数值", "closing the mode menu (Escape) cancels a pending mode without changing the mapping");
                board.BeginModeSelection(0); board.PendingMode = "粘贴链接"; board.Width++; CommitMode(board);
                Check(p1.Mode == "只写数值", "resizing closes a pending editor without applying it");
                board.BeginModeSelection(0); board.PendingMode = "粘贴链接"; form.ShowPage(1); CommitMode(board);
                Check(p1.Mode == "只写数值", "changing pages cancels the mode editor");
                var planField = typeof(DataMigrationForm).GetField("lastPlan", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                foreach (string expected in new[] { "数值+格式", "只写公式", "公式+格式", "只写格式", "粘贴链接", "只写数值" })
                {
                    planField.SetValue(form, new MigrationPlan());
                    board.BeginModeSelection(0); board.PendingMode = expected; CommitMode(board);
                    Check(p1.Mode == expected && planField.GetValue(form) == null, "dropdown selects " + expected + " and invalidates precheck");
                }
                foreach (string expected in new[] { "数值+格式", "只写公式", "公式+格式", "只写格式", "粘贴链接", "只写数值" })
                    Check(board.CycleMode(0) && p1.Mode == expected, "clicking the mode pill steps to " + expected);
                var unchangedPlan = new MigrationPlan(); planField.SetValue(form, unchangedPlan);
                board.BeginModeSelection(0); CommitMode(board);
                Check(ReferenceEquals(planField.GetValue(form), unchangedPlan), "choosing the existing mode preserves precheck");
                board.BeginModeSelection(0); board.PendingMode = "粘贴链接"; board.ReadOnly = true; CommitMode(board);
                Check(p1.Mode == "只写数值" && !board.BeginModeSelection(0), "running locks and closes the dropdown without applying a pending choice"); board.ReadOnly = false;
                board.BeginModeSelection(0); board.PendingMode = "粘贴链接";
                board.UpdateItem(0, new MigrationMappingBoard.Entry { ValuesOnly = true }); CommitMode(board);
                Check(!p1.PasteLinks, "a replaced mapping entry rejects a stale dropdown commit"); form.RefreshResults();
                p1.Source.Trigger = new MigrationTrigger();
                form.RefreshResults(); board.BeginModeSelection(0);
                Check(board.ModeChoices.SequenceEqual(new[] { "只写数值", "数值+格式" }), "recalculation dropdown only offers two values modes");
                board.PendingMode = "数值+格式"; CommitMode(board); Check(p1.ValuesOnly && p1.KeepFormat && !p1.PasteLinks, "recalculation selects values with formats");
                board.BeginModeSelection(0); board.PendingMode = "只写数值"; CommitMode(board); Check(p1.ValuesOnly && !p1.KeepFormat, "recalculation selects values without formats");
                Check(board.CycleMode(0) && p1.KeepFormat && board.CycleMode(0) && !p1.KeepFormat && p1.ValuesOnly, "clicking a recalculation pill steps between its two values modes");
                Reject(delegate { p1.SelectMode(MigrationPair.ModeLinks); }, "recalculation cannot select links through the model");
                p1.Source.Trigger = null;
                Check(p3 == p1 && form.PairItems().Count == 2 && form.Sources.Rows.Count == 2 && form.Targets.Rows.Count == 1 && form.Mappings.Count == 2, "a new target fills the first row without one instead of shifting pairs");
                form.SwapSides(0, 1, true);
                Check(p2.Target == t1 && !p1.Target.Used && p1.Source == s1 && p2.Source == s2, "dragging a side in 映射 swaps only that side between two rows");
                Check(!form.Mappings[0].TargetUsed && form.Mappings[1].TargetUsed && form.Mappings.Current == 1 && form.Mappings.CurrentTarget, "a row without a target shows as unpaired and the moved card stays selected");
                form.ShowPage(2); form.SwapSides(1, 0, false);
                Check(p1.Source == s2 && p2.Source == s1 && p2.Target == t1, "source cards swap independently of targets");
            }
            using (var form = new DataMigrationForm())
            {
                form.AddBook(false, @"D:\x1.xlsx", new[] { "A", "B" }); form.AddBook(false, @"D:\x2.xlsx", new[] { "A" }); form.AddBook(false, @"D:\x3.xlsx", new[] { "B" });
                var side = new MigrationSide { Sheet = "B", Range = "A1:B2" }; side.Books.AddRange(new[] { @"D:\x1.xlsx", @"D:\x2.xlsx", @"D:\x3.xlsx" });
                var pair = form.AddSide(false, side);
                Check(form.ApplySheet(pair, false, "A") == 1, "a sheet only some workbooks have applies to those that have it");
                var rows = form.PairItems();
                Check(rows.Count == 2 && rows[0] == pair && pair.Source.Sheet == "A" && pair.Source.Books.Count == 2
                    && rows[1].Source.Books.Count == 1 && rows[1].Source.Books[0].EndsWith("x3.xlsx") && rows[1].Source.Sheet == "B" && rows[1].Source.Range == "A1:B2" && !rows[1].Target.Used,
                    "workbooks lacking the sheet are split into the next row and keep their old sheet");
            }
            using (var form = new DataMigrationForm())
            {
                form.AddBook(false, @"D:\s1.xlsx", new[] { "甲", "乙" }); form.AddBook(false, @"D:\s2.xlsx", new[] { "乙" }); form.AddBook(false, @"D:\s3.xlsx", new[] { "丙" });
                Check(form.AddSideRows(false, new[] { @"D:\s1.xlsx", @"D:\s2.xlsx", @"D:\s3.xlsx" }) == 3 && form.Sources.Rows.Count == 3, "each added source workbook gets its own source range");
                var rows = form.PairItems();
                Check(rows.All(p => p.Source.Books.Count == 1 && p.Source.Range.Length == 0) && rows[0].Source.Sheet == "甲" && rows[2].Source.Sheet == "丙", "a new per-workbook range starts on that workbook's first sheet without guessing an area");
                Check(form.AddSideRows(false, new[] { @"D:\s2.xlsx" }) == 0, "adding a workbook again reuses its empty range");
                form.AddBook(true, @"D:\t1.xlsx", new[] { "目标" }); form.AddBook(true, @"D:\t2.xlsx", new[] { "目标" });
                Check(form.AddSideRows(true, new[] { @"D:\t1.xlsx", @"D:\t2.xlsx" }) == 2 && form.Targets.Rows.Count == 2 && rows[0].Target.Used && rows[1].Target.Used && !rows[2].Target.Used
                    && rows[0].Target.Books[0].EndsWith("t1.xlsx") && rows[0].Target.Sheet == "目标", "target workbooks also get one range each, filling mapping rows in order");
            }
            using (var form = new DataMigrationForm())
            {
                var source = new MigrationSide { Sheet = "表", Range = "A1:B2" }; source.Books.Add(@"D:\one.xlsx");
                var first = form.AddSide(false, source);
                var second = form.AddMappingRow();
                Check(second != null && second.Source == first.Source && form.PairItems().Count == 2 && form.Sources.Rows.Count == 1, "新增映射 shares the selected source range instead of copying it");
                var t1 = new MigrationSide { Sheet = "表", Range = "C1" }; t1.Books.Add(@"D:\t1.xlsx");
                var t2 = new MigrationSide { Sheet = "表", Range = "E1" }; t2.Books.Add(@"D:\t2.xlsx");
                form.AddSide(true, t1); form.AddSide(true, t2);
                Check(first.Target == t1 && second.Target == t2 && form.Targets.Rows.Count == 2, "new target ranges fill the shared rows, so one source maps to several targets");
                source.Range = "A1:B3"; Check(second.Source.Range == "A1:B3", "editing a shared source range changes every mapping that uses it");
                form.ShowPage(2); Check(!form.SidebarAvailable, "the range panel cannot be opened from 映射");
                form.ShowPage(0); Check(form.SidebarAvailable, "the range panel toggle returns on the range pages");
            }
        }
        // Presets keep the setup only: shared sources, 逐行重算, write modes and
        // relocated paths survive a save/apply round trip; newer formats are refused.
        private static void RecalculationTargetChecks(string root)
        {
            string source = Path.Combine(root, "trigger-targets.xlsx");
            using (var file = File.Create(source)) using (var zip = new ZipArchive(file, ZipArchiveMode.Create))
            {
                var parts = new Dictionary<string, string> {
                    { "xl/workbook.xml", "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"表\" sheetId=\"1\" r:id=\"r1\"/></sheets></workbook>" },
                    { "xl/_rels/workbook.xml.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"r1\" Target=\"worksheets/sheet1.xml\"/></Relationships>" },
                    { "xl/worksheets/sheet1.xml", "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData><row r=\"2\"><c r=\"G2\" t=\"inlineStr\"><is><t>触发值</t></is></c><c r=\"H2\" t=\"inlineStr\"><is><t>目标甲</t></is></c></row></sheetData></worksheet>" }
                };
                foreach (var part in parts) using (var writer = new StreamWriter(zip.CreateEntry(part.Key).Open())) writer.Write(part.Value);
            }
            using (var form = new DataMigrationForm())
            {
                string first = Path.Combine(root, "目标甲.xlsx"); string other = Path.Combine(root, "目标乙.xlsx"); string unused = Path.Combine(root, "未使用.xlsx");
                foreach (string path in new[] { first, other, unused }) form.AddBook(true, path, new[] { "表" });
                var src = new MigrationSide { Sheet = "表", Range = "A1", Trigger = new MigrationTrigger { Values = "G2:G2" } }; src.Books.Add(source);
                var recalc = form.AddPair(new MigrationPair { Source = src, Target = new MigrationSide { Sheet = "表", Range = "D1" }, ValuesOnly = true });
                var plain = src.Copy(); plain.Trigger = null;
                var redundant = new MigrationSide { Sheet = "表", Range = "E1" }; redundant.Books.Add(first);
                var duplicate = form.AddPair(new MigrationPair { Source = plain, Target = redundant });
                var kept = new MigrationSide { Sheet = "表", Range = "F1" }; kept.Books.Add(other);
                var independent = form.AddPair(new MigrationPair { Source = plain.Copy(), Target = kept });
                var gap = new MigrationSide { Sheet = "表", Range = "G1" }; gap.Books.Add(unused); form.AddSide(true, gap);
                form.RefreshResults();
                Check(src.Trigger.Error.Length == 0 && src.Trigger.Rows.Count == 1 && src.Trigger.Rows[0].Path == first, "recalculation source names resolve to the added target workbook");
                var flags = System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance;
                string prepared = src.Trigger.Signature; src.Trigger.Signature = "presentation-only";
                object[] summaryArguments = { src, null };
                string text = (string)typeof(DataMigrationForm).GetMethod("TriggerSummary", flags).Invoke(form, summaryArguments);
                typeof(DataMigrationForm).GetMethod("MatchedFiles", flags).Invoke(form, new object[] { recalc.Target });
                Check(src.Trigger.Signature == "presentation-only" && text.Contains("已匹配1行（缺失0行）"),
                    "summary and matched-file presentation reuse prepared results without reading files or rematching the target pool");
                src.Trigger.Signature = prepared;
                typeof(DataMigrationForm).GetMethod("DropUnmatchedTargets", flags).Invoke(form, new object[0]);
                typeof(DataMigrationForm).GetMethod("RebuildViews", flags).Invoke(form, new object[0]);
                Check(!duplicate.Target.Used && !form.PairItems().Any(p => p.Target == gap) && form.Targets.Rows.Count == 2, "applying recalculation removes redundant and unpaired ordinary target rows");
                Check(duplicate.Target.Sheet == "表" && duplicate.Target.Range == "E1" && duplicate.Target.Books.Count == 0, "a cleared target row remembers its sheet and range");
                Check(independent.Target == kept && kept.Used && form.TargetBooks.Select(b => b.Path).OrderBy(p => p).SequenceEqual(new[] { first, other }.OrderBy(p => p)), "target cleanup preserves an independent ordinary mapping and its workbook");
                src.Trigger.Chosen[2] = other; src.Trigger.Signature = ""; form.RefreshResults();
                Check(src.Trigger.Rows[0].Path == other && src.Trigger.Rows[0].Kind == "已确认" && form.Targets.Rows.Cast<System.Windows.Forms.DataGridViewRow>().Count(r => r.Tag == recalc) == 1
                    && form.Targets.Rows.Cast<System.Windows.Forms.DataGridViewRow>().Any(r => r.Tag == recalc && Convert.ToString(r.Cells["book"].Value) == "【重算】" + Path.GetFileName(other)), "manual trigger target override shows the chosen workbook in the by-name row, tagged 【重算】");
                typeof(DataMigrationForm).GetMethod("SplitByFile", flags).Invoke(form, new object[] { src });
                Check(recalc.Target.Books.Count == 1 && recalc.Target.Books[0] == other && recalc.Target.Range == "D1" && form.PairItems().Count(p => p.Source == src) == 1,
                    "applying recalculation turns the by-name row into an ordinary row per matched workbook, keeping its sheet and range");
                Check(recalc.Target.Recalc && !recalc.Target.Copy().Recalc, "the applied recalculation row is tagged; a copy of it is ordinary");
                src.Trigger.Chosen[2] = ""; src.Trigger.Signature = ""; form.RefreshResults();
                Check(src.Trigger.Rows[0].Kind == "跳过" && src.Trigger.Rows[0].Path.Length == 0 && independent.Target == kept, "skipping a trigger target leaves the independent ordinary mapping intact");
                // Deleting the 【重算】 row, then 应用 again, brings it back with its sheet and range.
                recalc.Target = new MigrationSide { Sheet = "表", Range = "D1" };
                src.Trigger.Chosen[2] = other; src.Trigger.Signature = ""; form.RefreshResults();
                typeof(DataMigrationForm).GetMethod("SplitByFile", flags).Invoke(form, new object[] { src });
                Check(recalc.Target.Used && recalc.Target.Recalc && recalc.Target.Books.SequenceEqual(new[] { other }) && recalc.Target.Sheet == "表" && recalc.Target.Range == "D1",
                    "applying recalculation again recreates a deleted recalculation row with its sheet and range");
            }
        }
        private static void ArchitectureChecks()
        {
            var source = new MigrationSide { Used = true, Sheet = "Source", Range = "A1" };
            source.Books.Add(@"C:\fixtures\source.xlsx");
            var target = new MigrationSide { Used = true, Sheet = "Target", Range = "B2" };
            target.Books.Add(@"C:\fixtures\target.xlsx");
            var first = new MigrationPair { Source = source, Target = target };
            var second = new MigrationPair { Source = source, Target = target.Copy(), Enabled = false };
            var book = new MigrationBook { Path = source.Books[0] }; book.Sheets.Add("Source");
            var snapshot = new MigrationConfiguration(new[] { first, second }, new[] { book }, new MigrationBook[0], false, true);
            Check(snapshot.Pairs[0].Source == snapshot.Pairs[1].Source && snapshot.Pairs[0].Source != source,
                "configuration preserves shared side identities without borrowing mutable UI data");
            source.Range = "Z9"; source.Books.Clear(); book.Sheets.Clear();
            var plan = new MigrationPlanner(snapshot).Build();
            Check(plan.Passes.Count == 1 && plan.Passes[0].Request.SourceRange == "A1" && plan.Passes[0].Request.AllowEmptySource
                && snapshot.SourceBooks[0].Sheets.SequenceEqual(new[] { "Source" }), "pure planning uses the captured configuration despite subsequent edits");
            Check(first.Instances.Count == 0, "planning alone does not modify the original row");
            snapshot.Attach(plan);
            Check(first.Instances.Count == 1 && first.Instances[0] == plan.Mappings[0] && plan.Mappings[0].Pair == first && second.Instances.Count == 0,
                "attaching results restores original UI row ownership without losing the plan mapping");

            source.Trigger = new MigrationTrigger { Values = "G2:G2" };
            source.Trigger.Chosen[2] = target.Books[0];
            source.Trigger.Extra.Add(new MigrationCondition { Cell = "C2", Value = "kept" });
            var row = new MigrationTriggerRow { Row = 2, Name = "Target", Path = target.Books[0], Kind = "已确认" };
            row.Candidates.Add(new KeyValuePair<string, double>(target.Books[0], 1)); source.Trigger.Rows.Add(row);
            var detached = new MigrationConfiguration(new[] { first }, new MigrationBook[0], new MigrationBook[0], true, false);
            source.Trigger.Extra[0].Value = "edited"; source.Trigger.Chosen.Clear(); row.Candidates.Clear(); row.Path = "";
            var captured = detached.Pairs[0].Source.Trigger;
            Check(captured.Extra[0].Value == "kept" && captured.Chosen[2] == target.Books[0] && captured.Rows[0].Ready && captured.Rows[0].Candidates.Count == 1,
                "trigger conditions, choices and resolved candidate rows are independently captured");

            var a = new MigrationSourceItem { Path = "source-a" }; var b = new MigrationSourceItem { Path = "source-b" };
            var sequence = new[] { new MigrationMapping { Source = a, TargetPath = "X" }, new MigrationMapping { Source = b, TargetPath = "Z" },
                new MigrationMapping { Source = a, TargetPath = "Y" }, new MigrationMapping { Source = a, TargetPath = "x" },
                new MigrationMapping { Source = a, TargetPath = "W" }, new MigrationMapping { Source = a, TargetPath = "Y" } };
            foreach (var item in sequence) item.Result = new MigrationTarget { Path = item.TargetPath };
            var grouped = new MigrationPlan(); MigrationPlanner.Group(grouped, sequence.ToList(), true);
            Check(grouped.Passes.Count == 3 && grouped.Passes[0].Mappings.SequenceEqual(new[] { sequence[0], sequence[2], sequence[4] })
                && grouped.Passes[1].Mappings.SequenceEqual(new[] { sequence[3], sequence[5] })
                && grouped.Passes[2].Mappings.SequenceEqual(new[] { sequence[1] }), "grouping retains source order, per-pass target order and case-insensitive uniqueness");
            Check(grouped.Passes.All(pass => pass.Request.ReplaceTargets && pass.Request.Targets.SequenceEqual(pass.Mappings.Select(m => m.Result))),
                "grouping retains the exact target result objects and save policy");

            var index = new MigrationNameIndex(new[] { @"C:\b\01-同名.xlsx", @"C:\a\同名.xlsx", @"C:\c\别名.xlsx" });
            Check(index.Find("同名").SequenceEqual(new[] { @"C:\b\01-同名.xlsx", @"C:\a\同名.xlsx" }) && index.Find("").Count == 0,
                "name index preserves pool order for manual candidates and never matches an empty name");
        }
        private static void CommitMode(MigrationMappingBoard board)
        { board.CommitModeSelection(); }
        private static void LinkConfigurationChecks()
        {
            Check(DataMigrationEngine.SourceLink(@"C:\fixtures\！01-示例集团.xlsx", "Sheet1", 3, 1) == "='C:\\fixtures\\[！01-示例集团.xlsx]Sheet1'!$A$3", "links quote Chinese workbook names and lock both coordinates");
            Check(DataMigrationEngine.SourceLink(@"C:\O'Brien\a'b.xlsx", "O'Brien", 12, 28) == "='C:\\O''Brien\\[a''b.xlsx]O''Brien'!$AB$12", "links escape apostrophes in directory, workbook and sheet");
            Check(DataMigrationEngine.SourceLink(@"\\server\shared folder\source.xlsx", "Data", 2, 2) == "='\\\\server\\shared folder\\[source.xlsx]Data'!$B$2", "links preserve UNC paths and spaces");
            var compare = new MigrationFormulaContext { SourcePath = @"D:\targets\target.xlsx", SourceSheet = "Target", TargetSheet = "Target" };
            compare.ExternalFiles["source.xlsx"] = @"C:\source\source.xlsx";
            Check(DataMigrationEngine.ComparisonFormula("='[source.xlsx]Data'!$A$3", compare) == DataMigrationEngine.ComparisonFormula("='C:\\source\\[source.xlsx]Data'!$A$3", compare), "open and closed source formula spellings compare equally");
            Check(DataMigrationEngine.ComparisonFormula("='[source.xlsx]Data'!A3", compare) != DataMigrationEngine.ComparisonFormula("='[source.xlsx]Data'!$A$3", compare), "link verification retains absolute reference locks");
            using (var form = new DataMigrationForm())
            {
                var source = new MigrationSide { Sheet = "Source", Range = "A3:B4" }; source.Books.Add(@"C:\source\source.xlsx");
                var target = new MigrationSide { Sheet = "Target", Range = "D8" }; target.Books.Add(@"D:\target\target.xlsx");
                var pair = form.AddSide(false, source); form.AddSide(true, target);
                foreach (string mode in MigrationPair.Modes)
                {
                    pair.SelectMode(mode);
                    var data = form.CapturePreset("six modes"); var mapping = MigrationJson.Maps(data, "mappings").Single();
                    Check(MigrationJson.Num(data, "version") == 2 && MigrationJson.Flag(mapping, "pasteLinks", false) == (mode == MigrationPair.ModeLinks), "preset v2 captures " + mode);
                    var restored = MigrationPresetCodec.Read((Dictionary<string, object>)MigrationJson.Parse(MigrationJson.Write(data)), null);
                    Check(restored.Pairs.Single().Mode == mode, "preset roundtrip restores " + mode);
                }
                pair.SelectMode(MigrationPair.ModeLinks);
                var snapshot = new MigrationConfiguration(form.PairItems(), new MigrationBook[0], new MigrationBook[0], false, false); pair.SelectMode(MigrationPair.ModeValues);
                var plan = new MigrationPlanner(snapshot).Build();
                Check(snapshot.Pairs.Single().PasteLinks && plan.Passes.Single().Request.PasteLinks && !plan.Passes.Single().Request.KeepFormat, "link flag survives an independent snapshot and request planning");
                var legacy = form.CapturePreset("legacy"); legacy["version"] = 1L;
                var old = MigrationJson.Maps(legacy, "mappings").Single(); old.Remove("pasteLinks"); old.Remove("keepFormat"); old["valuesOnly"] = false;
                Check(MigrationPresetCodec.Read(legacy, null).Pairs.Single().Mode == MigrationPair.ModeFormulaFormat, "v1 preset without format flag retains its historical default");
                old["formatOnly"] = true;
                Check(MigrationPresetCodec.Read(legacy, null).Pairs.Single().Mode == "只写格式", "legacy format-only preset uses the renamed label");
                pair.SelectMode(MigrationPair.ModeLinks); form.AddMappingRow();
                Check(form.PairItems().Last().PasteLinks, "copying a mapping preserves link mode");
                var secondTarget = new MigrationSide { Sheet = "Target", Range = "H8" }; secondTarget.Books.Add(@"D:\target\target.xlsx");
                form.AddSide(true, secondTarget); form.PairItems().Last().SelectMode(MigrationPair.ModeFormula);
                var grouped = new MigrationPlanner(new MigrationConfiguration(form.PairItems(), new MigrationBook[0], new MigrationBook[0], false, false)).Build();
                Check(grouped.Passes.Count == 2 && grouped.Passes.Count(p => p.Request.PasteLinks) == 1 && grouped.Passes.Count(p => !p.Request.PasteLinks && !p.Request.ValuesOnly) == 1, "same source with link and formula modes remains in separate planned passes");
                form.SetTrigger(source, new MigrationTrigger());
                Check(form.PairItems().All(p => !p.PasteLinks && p.ValuesOnly), "applying recalculation converts shared linked rows to values");
            }
        }
        private static void LinkChecks(string root)
        {
            string folder = Path.Combine(root, "links O'Brien"); Directory.CreateDirectory(folder);
            var app = new FakeApp(); var source = app.Workbooks.Create(Path.Combine(folder, "！01-示例集团.xlsx"), "O'Brien");
            source.Sheets[0].Cell("A3").Value2 = 12.5; source.Sheets[0].Cell("B3").Formula = "=UnsupportedName()";
            source.Sheets[0].Cell("A4").Value2 = "plain"; source.Sheets[0].Cell("C3").Formula = "=A3";
            source.Names.Add("Unneeded", "=UnsupportedName()", true); source.Links.Add(Path.Combine(folder, "upstream.xlsx")); source.Persist();
            string sourceHash = DataMigrationHost.Digest(source.FullName);
            app.ActiveWorkbook = source; app.ActiveSheet = source.Sheets[0];
            foreach (string ext in new[] { ".xlsx", ".xlsm", ".xls", ".et" })
            {
                var original = app.Workbooks.Create(Path.Combine(folder, "target" + ext), "Target");
                foreach (string cell in new[] { "D8", "E8", "D9", "E9" }) { original.Sheets[0].Cell(cell).Value2 = 99.0; original.Sheets[0].Cell(cell).NumberFormat = "0.000"; original.Sheets[0].Cell(cell).Font.Bold = true; }
                original.Sheets[0].Cell("D8").NumberFormat = "@"; original.Names.Add("Unneeded", "=7", true); original.Persist(); original.Close(false);
                string hash = DataMigrationHost.Digest(original.FullName);
                var target = new MigrationTarget { Path = original.FullName, Sheet = "Target", Start = "D8" };
                var request = new MigrationRequest { SourcePath = source.FullName, SourceSheet = "O'Brien", SourceRange = "A3:B4", PasteLinks = true, ReplaceTargets = ext == ".xlsm", Targets = new List<MigrationTarget> { target } };
                using (var host = new DataMigrationHost(app))
                {
                    Drain(host.Preview(request)); Check(target.Ready && target.ValueFallbacks == 0 && request.NonEmptyCells == 3 && request.FormulaCells == 1, "link preview ignores source dependencies and counts actual source content " + ext);
                    Drain(host.Execute()); Check(target.Status == "成功", "link save/reopen succeeds in managed fixture " + ext);
                    var output = app.Workbooks.Open(target.Output, 0, true); var sheet = output.Sheets[0];
                    string prefix = "='" + folder.Replace("'", "''") + "\\[！01-示例集团.xlsx]O''Brien'!";
                    Check(sheet.Cell("D8").Formula == prefix + "$A$3" && sheet.Cell("E8").Formula == prefix + "$B$3" && sheet.Cell("D9").Formula == prefix + "$A$4" && sheet.Cell("E9").Formula == prefix + "$B$4", "every rectangle cell including blanks links to the original absolute source coordinates " + ext);
                    Check(sheet.Cell("D8").NumberFormat == "General" && sheet.Cell("E8").NumberFormat == "0.000" && sheet.Cell("D8").Font.Bold && sheet.Cell("D8").RowHeight == 15 && sheet.Cell("D8").ColumnWidth == 8.5, "links retain target style and dimensions with text-format correction " + ext);
                    Check(output.Links.SequenceEqual(new[] { source.FullName }) && output.Names.Items.Single().RefersTo == "=7", "only the intended original source link is added " + ext);
                    output.Close(false);
                    Check(request.ReplaceTargets ? target.Output == original.FullName && DataMigrationHost.Digest(target.Backup) == hash : DataMigrationHost.Digest(original.FullName) == hash, "link output honors copy or backed-up replacement policy " + ext);
                    Check(File.ReadAllText(Path.Combine(host.DiagnosticDirectory, "migration.log")).Contains("粘贴链接"), "link mode appears in the migration log " + ext);
                }
                source.Close(false);
            }
            Check(DataMigrationHost.Digest(source.FullName) == sourceHash, "linked source bytes are unchanged with open and closed originals");
            var mergedSource = app.Workbooks.Create(Path.Combine(folder, "merged-source.xlsx"), "Source");
            mergedSource.Sheets[0].Cell("A3").Value2 = "merged"; mergedSource.Sheets[0].Cell("B3").Value2 = "hidden source data";
            mergedSource.Sheets[0].Merges.Add("A3:B4"); mergedSource.Persist(); mergedSource.Close(false);
            var mergedBook = app.Workbooks.Create(Path.Combine(folder, "merged-target.xlsx"), "Target"); mergedBook.Persist(); mergedBook.Close(false);
            var mergedTarget = new MigrationTarget { Path = mergedBook.FullName, Sheet = "Target", Start = "D8" };
            using (var host = new DataMigrationHost(app))
            {
                Drain(host.Preview(new MigrationRequest { SourcePath = mergedSource.FullName, SourceSheet = "Source", SourceRange = "A3:B4", PasteLinks = true, Targets = new List<MigrationTarget> { mergedTarget } }));
                Drain(host.Execute()); Check(mergedTarget.Status == "成功", "complete merged source links survive save and reopen");
                var output = app.Workbooks.Open(mergedTarget.Output, 0, true); var sheet = output.Sheets[0];
                Check(sheet.Merges.SequenceEqual(new[] { "D8:E9" }) && sheet.Cell("D8").Formula.EndsWith("!$A$3") && !sheet.Cell("E8").HasFormula && sheet.Cell("E8").Value2 == null, "only the merge anchor receives a link; followers receive no source hidden data"); output.Close(false);
            }
            var combined = app.Workbooks.Create(Path.Combine(folder, "combined.xlsx"), "Target"); combined.Sheets[0].Cell("G18").Value2 = "keep"; combined.Persist(); combined.Close(false);
            using (var host = new DataMigrationHost(app) { HoldTargets = true })
            {
                var links = new MigrationTarget { Path = combined.FullName, Sheet = "Target", Start = "D8" };
                var values = new MigrationTarget { Path = combined.FullName, Sheet = "Target", Start = "E12" };
                var formulas = new MigrationTarget { Path = combined.FullName, Sheet = "Target", Start = "F15" };
                var formats = new MigrationTarget { Path = combined.FullName, Sheet = "Target", Start = "G18" };
                foreach (var request in new[] {
                    new MigrationRequest { SourcePath = source.FullName, SourceSheet = "O'Brien", SourceRange = "A3:B4", PasteLinks = true, Targets = new List<MigrationTarget> { links } },
                    new MigrationRequest { SourcePath = source.FullName, SourceSheet = "O'Brien", SourceRange = "A3", ValuesOnly = true, Targets = new List<MigrationTarget> { values } },
                    new MigrationRequest { SourcePath = source.FullName, SourceSheet = "O'Brien", SourceRange = "C3", Targets = new List<MigrationTarget> { formulas } },
                    new MigrationRequest { SourcePath = source.FullName, SourceSheet = "O'Brien", SourceRange = "A3", ValuesOnly = true, KeepFormat = true, FormatOnly = true, Targets = new List<MigrationTarget> { formats } } })
                { Drain(host.Preview(request)); host.Keep(); }
                Drain(host.ExecuteKept());
                Check(new[] { links, values, formulas, formats }.All(t => t.Status == "成功" && t.Output == links.Output), "mixed link/value/formula/format passes share one successful output");
                var output = app.Workbooks.Open(links.Output, 0, true); var sheet = output.Sheets[0];
                Check(sheet.Cell("D8").Formula.EndsWith("!$A$3") && Convert.ToDouble(sheet.Cell("E12").Value2) == 12.5 && sheet.Cell("F15").Formula == "=D15" && (string)sheet.Cell("G18").Value2 == "keep", "mixed batch preserves all four independent write behaviors"); output.Close(false);
            }
            foreach (string fault in new[] { "save", "unexpected", "missing", "cancel" })
            {
                var original = app.Workbooks.Create(Path.Combine(folder, fault + ".xlsx"), "Target"); original.FailSave = fault == "save"; original.Persist(); original.Close(false);
                string hash = DataMigrationHost.Digest(original.FullName);
                var target = new MigrationTarget { Path = original.FullName, Sheet = "Target", Start = "D8" };
                var request = new MigrationRequest { SourcePath = source.FullName, SourceSheet = "O'Brien", SourceRange = "A3", PasteLinks = true, ReplaceTargets = true, Targets = new List<MigrationTarget> { target } };
                using (var host = new DataMigrationHost(app))
                {
                    Drain(host.Preview(request));
                    if (fault == "unexpected" || fault == "missing") app.Workbooks.Opened = delegate(string path) {
                        if (!Path.GetFileName(path).StartsWith(".xy-migration-") || Path.GetFileName(path).StartsWith(".xy-migration-read-")) return;
                        if (fault == "unexpected") app.Workbooks.Find(path).Links.Add(Path.Combine(folder, "unplanned.xlsx")); else app.Workbooks.Find(path).Links.Clear(); };
                    if (fault == "cancel")
                    {
                        using (var steps = host.Execute().GetEnumerator()) { Check(steps.MoveNext(), "link execution yields before writing"); host.Cancelled = true; while (steps.MoveNext()) { } }
                    }
                    else Drain(host.Execute());
                    Check(target.Status == (fault == "cancel" ? "已取消" : "失败") && target.Output.Length == 0 && target.Backup.Length == 0 && DataMigrationHost.Digest(original.FullName) == hash, "link " + fault + " never commits an unverified output");
                }
                app.Workbooks.Opened = null;
            }
            var empty = new MigrationRequest { SourcePath = source.FullName, SourceSheet = "O'Brien", SourceRange = "Z99", PasteLinks = true };
            using (var host = new DataMigrationHost(app)) Reject(delegate { Drain(host.Preview(empty)); }, "link mode preserves explicit all-blank opt-in");
            empty.AllowEmptySource = true;
            using (var host = new DataMigrationHost(app)) { Drain(host.Preview(empty)); Check(empty.NonEmptyCells == 0, "explicit all-blank link source is permitted"); }
            using (var host = new DataMigrationHost(app)) Reject(delegate { Drain(host.Preview(new MigrationRequest { PasteLinks = true, TriggerCell = "B2" })); }, "direct link request cannot use recalculation");
        }
        private static void PresetChecks()
        {
            var parsed = (Dictionary<string, object>)MigrationJson.Parse(MigrationJson.Write(MigrationJson.Obj("text", "引号\"反斜杠\\换行\n", "list", new List<object> { 1L, 2.5, true, null })));
            var items = (List<object>)parsed["list"];
            Check((string)parsed["text"] == "引号\"反斜杠\\换行\n" && (long)items[0] == 1 && (double)items[1] == 2.5 && (bool)items[2] && items[3] == null, "preset JSON round-trips text, numbers, flags and null");
            using (var form = new DataMigrationForm())
            {
                var source = new MigrationSide { Sheet = "表", Range = "C5:C29" }; source.Books.Add(@"D:\预设\源.xlsx");
                form.AddSide(false, source); form.AddMappingRow();
                var t1 = new MigrationSide { Sheet = "Notes", Range = "H944" }; t1.Books.Add(@"D:\预设\目标一.xlsx");
                var t2 = new MigrationSide { Sheet = "Notes", Range = "H970" }; t2.Books.Add(@"D:\预设\子目录\目标二.xlsx");
                form.AddSide(true, t1); form.AddSide(true, t2);
                form.SetTrigger(source, new MigrationTrigger { Cell = "B2", Values = "G2:G30", NameColumn = "H" });
                source.Trigger.Chosen[9] = @"D:\预设\子目录\目标二.xlsx";
                source.Trigger.Extra.Add(new MigrationCondition { Cell = "C2", Value = "2025" }); source.Trigger.Extra.Add(new MigrationCondition { Cell = "D2", Column = "I" });
                form.PairItems()[1].ValuesOnly = true; form.PairItems()[1].Enabled = false;
                var data = form.CapturePreset("测试预设");
                var preset = MigrationPreset.From((Dictionary<string, object>)MigrationJson.Parse(MigrationJson.Write(data)), null);
                Check(preset.Name == "测试预设" && preset.Sources == 1 && preset.Targets == 2 && preset.Mappings == 2 && MigrationJson.Str(data, "root") == @"D:\预设", "a preset lists a shared source once and records the common folder");
                Check(preset.Relative(@"D:\预设\子目录\目标二.xlsx") == @"子目录\目标二.xlsx", "workbooks keep their path relative to the common folder for relocation");
                using (var other = new DataMigrationForm())
                {
                    other.ApplyPresetData(preset.Data, new Dictionary<string, string> { { @"D:\预设\子目录\目标二.xlsx", @"E:\别处\目标二.xlsx" } });
                    var rows = other.PairItems();
                    Check(rows.Count == 2 && rows[0].Source == rows[1].Source && rows[0].Source.Trigger != null && rows[0].Source.Trigger.Cell == "B2" && rows[0].Source.Trigger.Values == "G2:G30"
                        && rows[0].Source.Trigger.Chosen[9] == @"E:\别处\目标二.xlsx", "applying restores the shared source, 逐行重算 and relocated choices");
                    Check(rows[0].Source.Trigger.Extra.Count == 2 && rows[0].Source.Trigger.Extra[0].Cell == "C2" && rows[0].Source.Trigger.Extra[0].Value == "2025" && rows[0].Source.Trigger.Extra[1].Column == "I",
                        "a preset keeps the extra 逐行重算 conditions (fixed value and per-row column)");
                    Check(rows[1].Target.Books[0] == @"E:\别处\目标二.xlsx" && rows[1].Target.Range == "H970" && !rows[1].Enabled && rows[1].ValuesOnly && other.Mappings.Count == 2, "applying restores targets, modes and enabled state with relocated paths");
                }
                data["version"] = 99L;
                Reject(delegate { MigrationPreset.From(data, null); }, "a preset from a newer format version is refused with a clear reason");
                data["version"] = 1L; data["format"] = "something-else";
                Reject(delegate { MigrationPreset.From(data, null); }, "a JSON file that is not a migration preset is refused");
            }
        }
        private static void MappingSummaryChecks()
        {
            var summary = new MigrationMappingBoard.MatchSummary { Rows = 13, Matched = 10, Missing = 3 };
            Check(summary.Text == "目标文件区域：13行；已匹配10行（缺失3行）。", "matched and missing wording uses a single physical summary line");
            summary.Additional = " 同名待选 2 行，跳过 1 行，另有 1 个目标工作簿未用到";
            Check(summary.SecondLine.Contains("同名待选 2 行") && !summary.Text.Contains("\n"), "additional match diagnostics remain available without shifting detail links");
            summary.Additional = "";
            var flags = System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance;
            foreach (float factor in new[] { 1f, 1.25f, 1.5f, 2f })
            using (var board = new MigrationMappingBoard())
            using (var bitmap = new Bitmap((int)(1600 * factor), (int)(180 * factor)))
            {
                foreach (string name in new[] { "body", "head" })
                {
                    var field = typeof(MigrationMappingBoard).GetField(name, flags); var old = (Font)field.GetValue(board);
                    field.SetValue(board, new Font(old.FontFamily, old.Size * factor, old.Style)); old.Dispose();
                }
                board.Size = bitmap.Size;
                board.SetItems(new[] { new MigrationMappingBoard.Entry { SourceUsed = true, TargetUsed = true, SourceTitle = "源区域", TargetTitle = "目标区域", Tone = 4, Status = summary.Text, StatusTip = summary.Text, Match = summary } }, 0, false);
                bitmap.SetResolution(96 * factor, 96 * factor);
                using (var graphics = Graphics.FromImage(bitmap)) typeof(MigrationMappingBoard).GetMethod("OnPaint", flags).Invoke(board, new object[] { new System.Windows.Forms.PaintEventArgs(graphics, new Rectangle(Point.Empty, bitmap.Size)) });
                int green = 0, red = 0;
                for (int y = 0; y < bitmap.Height; y++) for (int x = 0; x < bitmap.Width; x++)
                { int color = bitmap.GetPixel(x, y).ToArgb(); if (color == Ui.Green.ToArgb()) green++; if (color == Ui.Danger.ToArgb()) red++; }
                Check(green > 0 && red > 0, "production mapping paints green matches and red missing counts at scale " + factor);
            }
        }
        private static void Render(string root, string path)
        {
            string source = Path.Combine(root, "render-source.xlsx");
            var rows = new System.Text.StringBuilder();
            for (int r = 2; r <= 17; r++) rows.Append("<row r=\"" + r + "\"><c r=\"G" + r + "\" t=\"inlineStr\"><is><t>触发" + r
                + "</t></is></c><c r=\"H" + r + "\" t=\"inlineStr\"><is><t>目标" + (r - 1) + "</t></is></c></row>");
            using (var file = File.Create(source)) using (var zip = new ZipArchive(file, ZipArchiveMode.Create))
            {
                var parts = new Dictionary<string, string> {
                    { "xl/workbook.xml", "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"表\" sheetId=\"1\" r:id=\"r1\"/></sheets></workbook>" },
                    { "xl/_rels/workbook.xml.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"r1\" Target=\"worksheets/sheet1.xml\"/></Relationships>" },
                    { "xl/worksheets/sheet1.xml", "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>" + rows + "</sheetData></worksheet>" }
                };
                foreach (var part in parts) using (var writer = new StreamWriter(zip.CreateEntry(part.Key).Open())) writer.Write(part.Value);
            }
            bool previous = DialogPreviewFactory.IsConstructing;
            DialogPreviewFactory.IsConstructing = true;
            try
            {
                using (var form = new DataMigrationForm())
                using (var parking = new System.Windows.Forms.Panel { Size = new Size(2100, 1600) })
                {
                    form.TopLevel = false; form.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None; parking.Controls.Add(form);
                    form.AddBook(false, source, new[] { "表" });
                    var pair = form.AddPair(new MigrationPair());
                    pair.Source.Books.Add(source); pair.Source.Sheet = "表"; pair.Source.Range = "C5:C24";
                    pair.Source.Trigger = new MigrationTrigger { Cell = "B2", Values = "G2:G17", NameColumn = "H2:H17" };
                    pair.Target.Sheet = "表"; pair.Target.Range = "H944";
                    for (int i = 1; i <= 3; i++)
                    {
                        string book = Path.Combine(root, "目标" + i + ".xlsx");
                        form.AddBook(true, book, new[] { "表" }); pair.Target.Books.Add(book);
                    }
                    form.RefreshResults(); form.ShowPage(2); form.ClientSize = new Size(1600, 760); form.Show(); form.PerformLayout();
                    var flags = System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance;
                    var status = (System.Windows.Forms.Label)typeof(DataMigrationForm).GetField("status", flags).GetValue(form);
                    Check(form.Mappings[0].Status == "目标文件区域：16行；已匹配3行（缺失13行）。", "render fixture exercises real filename matching for 16 rows");
                    Check(form.Mappings[0].Match != null && form.Mappings[0].Match.Matched == 3 && form.Mappings[0].Match.Missing == 13, "mapping drawing and plain text share actual match statistics");
                    var details = (System.Windows.Forms.RichTextBox)typeof(DataMigrationForm).GetField("details", flags).GetValue(form);
                    var links = (Dictionary<int, KeyValuePair<MigrationSide, MigrationTriggerRow>>)typeof(DataMigrationForm).GetField("detailLinks", flags).GetValue(form);
                    Check(links.Count == 16 && links.All(link => details.Lines[link.Key].StartsWith("第 " + link.Value.Value.Row + " 行", StringComparison.Ordinal)),
                        "each clickable detail index still points at its own trigger row after the display split");
                    form.Message("已应用逐行重算：触发值 G2:G17 → B2，目标文件区域 H2:H17；目标区域已清理，只保留逐行重算匹配到的工作簿。目标区域页这一行的工作簿即匹配到的文件。");
                    Check(status.Font.Bold && status.ForeColor == Ui.Blue, "production explanation uses blue bold text");
                    using (var bitmap = new Bitmap(form.Width, form.Height)) { form.DrawToBitmap(bitmap, new Rectangle(Point.Empty, bitmap.Size)); bitmap.Save(path, ImageFormat.Png); }
                    form.ClientSize = new Size(1000, 640); form.PerformLayout();
                    using (var bitmap = new Bitmap(form.Width, form.Height)) { form.DrawToBitmap(bitmap, new Rectangle(Point.Empty, bitmap.Size)); bitmap.Save(Path.ChangeExtension(path, ".narrow.png"), ImageFormat.Png); }
                    Check(status.GetPreferredSize(new Size(status.Width, 0)).Height <= status.Height + 2, "long bold explanation is measured within the narrower window");
                    form.ShowPage(0);
                    typeof(DataMigrationForm).GetMethod("ShowTriggerPopup", flags).Invoke(form, new object[] { form.Sources, pair, null });
                    var drop = (System.Windows.Forms.ToolStripDropDown)typeof(DataMigrationForm).GetField("triggerPopup", flags).GetValue(form);
                    var panel = ((System.Windows.Forms.ToolStripControlHost)drop.Items[0]).Control;
                    Check(!drop.Visible && DialogFormatting.Walk(panel).Any(c => c is System.Windows.Forms.Label && c.Text == "目标文件区域"), "actual recalculation panel constructs its new label without displaying a popup");
                    panel.CreateControl(); panel.PerformLayout(); panel.Size = panel.PreferredSize;
                    using (var bitmap = new Bitmap(panel.Width, panel.Height)) { panel.DrawToBitmap(bitmap, new Rectangle(Point.Empty, bitmap.Size)); bitmap.Save(Path.ChangeExtension(path, ".popup.png"), ImageFormat.Png); }
                    form.ShowPage(2); form.Mappings.Size = new Size(1400, 250);
                    using (var bitmap = new Bitmap(1400, 250))
                    {
                        bitmap.SetResolution(144, 144);
                        using (var graphics = Graphics.FromImage(bitmap))
                            typeof(MigrationMappingBoard).GetMethod("OnPaint", flags).Invoke(form.Mappings,
                                new object[] { new System.Windows.Forms.PaintEventArgs(graphics, new Rectangle(Point.Empty, bitmap.Size)) });
                        bitmap.Save(Path.ChangeExtension(path, ".board150.png"), ImageFormat.Png);
                    }
                    using (var board = new MigrationMappingBoard { Size = new Size(1500, 530) })
                    using (var bitmap = new Bitmap(1500, 530))
                    {
                        parking.Controls.Add(board);
                        var entries = MigrationPair.Modes.Select(mode => {
                            var row = new MigrationPair(); row.SelectMode(mode);
                            return new MigrationMappingBoard.Entry { SourceUsed = true, TargetUsed = true, SourceTitle = "源工作簿 · Sheet1", SourceDetail = "A3:B4", TargetTitle = "目标工作簿 · 工作表", TargetDetail = "D8:E9", ValuesOnly = row.ValuesOnly, KeepFormat = row.KeepFormat, FormatOnly = row.FormatOnly, PasteLinks = row.PasteLinks, Status = "就绪（待预检）" }; }).ToList();
                        board.SetItems(entries, 5, false); board.CreateControl();
                        board.DrawToBitmap(bitmap, new Rectangle(Point.Empty, bitmap.Size)); bitmap.Save(Path.ChangeExtension(path, ".modes.png"), ImageFormat.Png);
                        Check(board.BeginModeSelection(5) && !board.ModeMenuOpen, "mode editor prepares without showing its menu");
                        board.DrawToBitmap(bitmap, new Rectangle(Point.Empty, bitmap.Size)); bitmap.Save(Path.ChangeExtension(path, ".editor.png"), ImageFormat.Png);
                    }
                }
            }
            finally { DialogPreviewFactory.IsConstructing = previous; }
        }

    }

    public sealed class FakeApp
    {
        public FakeBooks Workbooks { get; private set; }
        public object ActiveWorkbook { get; set; } public object ActiveSheet { get; set; }
        public bool DisplayAlerts { get; set; } public bool AskToUpdateLinks { get; set; } public bool EnableEvents { get; set; } public bool ScreenUpdating { get; set; }
        public int AutomationSecurity { get; set; } public int Calculation { get; set; }
        public FakeApp() { Workbooks = new FakeBooks(); DisplayAlerts = AskToUpdateLinks = EnableEvents = ScreenUpdating = true; AutomationSecurity = 1; Calculation = -4105; }
    }
    public sealed class FakeBooks
    {
        public bool DirtyReadCopies;
        public Action<string> Opened;
        public int OpenCalls;
        public List<FakeBook> Items = new List<FakeBook>();
        public Dictionary<string, FakeBook> Storage = new Dictionary<string, FakeBook>();
        public int Count { get { return Items.Count; } }
        public FakeBook Item(int i) { return Items[i - 1]; }
        public FakeBook Find(string path) { return Items.FirstOrDefault(b => b.FullName == path); }
        public FakeBook Create(string path, string sheet)
        { var b = new FakeBook(this, path); b.Sheets.Add(new FakeSheet(b, sheet)); Items.Add(b); b.Persist(); return b; }
        public FakeBook Open(string path, int update, bool readOnly)
        { OpenCalls++; var original = Storage[File.ReadAllText(path)]; var copy = original.Clone(path); if (DirtyReadCopies && Path.GetFileName(path).StartsWith(".xy-migration-read-")) copy.Saved = false; Items.Add(copy); if (Opened != null) Opened(path); return copy; }
    }
    public sealed class FakeBook
    {
        public FakeBooks Owner; public string FullName { get; set; } public bool Saved { get; set; } public bool ProtectStructure { get; set; }
        public bool FailSave { get; set; }
        public List<FakeSheet> Sheets = new List<FakeSheet>(); public List<string> Links = new List<string>(); public FakeNames Names = new FakeNames();
        public FakeSheets Worksheets { get { return new FakeSheets(Sheets); } }
        private readonly FakeWindows windows = new FakeWindows();
        public FakeWindows Windows { get { return windows; } }
        public FakeBook(FakeBooks owner, string path) { Owner = owner; FullName = path; Saved = true; }
        public object LinkSources(int type) { return Links.Count == 0 ? null : Links.ToArray(); }
        public void Activate() { }
        public void Close(bool save) { Owner.Items.Remove(this); }
        public void Save() { if (FailSave) throw new IOException("fixture save failure"); Persist(); }
        public void Persist() { string id = Guid.NewGuid().ToString("N"); Owner.Storage[id] = Clone(FullName); File.WriteAllText(FullName, id); Saved = true; }
        public FakeBook Clone(string path)
        {
            var b = new FakeBook(Owner, path) { FailSave = FailSave }; b.Links.AddRange(Links); b.Windows.Item(1).Visible = Windows.Item(1).Visible;
            foreach (var name in Names.Items) b.Names.Add(name.Name, name.RefersTo, name.Visible);
            foreach (var s in Sheets)
            { var copy = new FakeSheet(b, s.Name); copy.Merges.AddRange(s.Merges); foreach (var p in s.Cells) copy.Cells[p.Key] = p.Value.Clone(copy); b.Sheets.Add(copy); } return b;
        }
    }
    public sealed class FakeSheets
    {
        private readonly List<FakeSheet> sheets; public FakeSheets(List<FakeSheet> value) { sheets = value; }
        public int Count { get { return sheets.Count; } }
        public FakeSheet Item(object key) { return key is int ? sheets[(int)key - 1] : sheets.First(s => s.Name == (string)key); }
    }
    public sealed class FakeWindow { public FakeWindow() { Visible = true; } public bool Visible { get; set; } public int ScrollRow { get; set; } public int ScrollColumn { get; set; } }
    public sealed class FakeWindows { private readonly FakeWindow window = new FakeWindow(); public int Count { get { return 1; } } public FakeWindow Item(int i) { return window; } }
    public sealed class FakeName { public string Name { get; set; } public string RefersTo { get; set; } public bool Visible { get; set; } }
    public sealed class FakeNames
    {
        public List<FakeName> Items = new List<FakeName>(); public int Count { get { return Items.Count; } }
        public FakeName Item(int i) { return Items[i - 1]; }
        public FakeName Add(string name, string formula, bool visible) { var n = new FakeName { Name = name, RefersTo = formula, Visible = visible }; Items.Add(n); return n; }
    }
    public sealed class FakeCount { public int Count { get; set; } }
    public sealed class PickerSelection
    {
        public object Parent { get; set; }
        public int Row { get; set; } public int Column { get; set; }
        public FakeCount Areas { get; set; } public FakeCount Rows { get; set; } public FakeCount Columns { get; set; }
        public PickerSelection() { Areas = new FakeCount { Count = 1 }; }
    }
    public sealed class FakeSheet
    {
        public object Parent { get { return Book; } }
        public FakeBook Book; public string Name { get; set; } public int Visible { get { return -1; } } public bool ProtectContents { get; set; }
        public FakeCount Rows { get { return new FakeCount { Count = 1048576 }; } } public FakeCount Columns { get { return new FakeCount { Count = 16384 }; } }
        public FakeCount ListObjects { get { return new FakeCount(); } } public FakeRanges Range { get { return new FakeRanges(this); } }
        public Dictionary<string, FakeCell> Cells = new Dictionary<string, FakeCell>();
        public List<string> Merges = new List<string>();
        public FakeSheet(FakeBook book, string name) { Book = book; Name = name; }
        public FakeCell Cell(string address) { FakeCell c; if (!Cells.TryGetValue(address, out c)) Cells[address] = c = new FakeCell(this, address); return c; }
        public void Activate() { }
    }
    public sealed class FakeRanges
    {
        private readonly FakeSheet sheet; public FakeRanges(FakeSheet s) { sheet = s; }
        public FakeCell this[string address] { get { return sheet.Cell(address); } }
    }
    public sealed class FakeCell
    {
        private FakeSheet sheet; private string formula; private object value;
        public string Address; public FakeCell(FakeSheet s, string a) { sheet = s; Address = a; RowHeight = 15; ColumnWidth = 8.5; NumberFormat = "General"; Font = new FakeFont(); Interior = new FakeFill(); Borders = new FakeBorders(); }
        public bool HasArray { get { return false; } } public bool HasSpill { get { return false; } }
        private string MergedAddress { get { return sheet.Merges.FirstOrDefault(m => MigrationArea.Parse(m).Contains(MigrationArea.Parse(Address))); } }
        public bool MergeCells { get { return MergedAddress != null; } }
        public FakeCell MergeArea { get { return sheet.Cell(MergedAddress); } }
        public int Row { get { return MigrationArea.Parse(Address).Row; } }
        public int Column { get { return MigrationArea.Parse(Address).Column; } }
        public FakeCount Rows { get { return new FakeCount { Count = MigrationArea.Parse(Address).Rows }; } }
        public FakeCount Columns { get { return new FakeCount { Count = MigrationArea.Parse(Address).Columns }; } }
        public bool HasFormula { get { return formula != null; } }
        public string Formula
        {
            get { return formula; }
            set
            {
                formula = value; this.value = null; sheet.Book.Saved = false;
                foreach (Match m in Regex.Matches(value, @"'((?:[^']|'')*)\[([^\]]+)\](?:[^']|'')*'!"))
                { string p = Path.Combine(m.Groups[1].Value.Replace("''", "'"), m.Groups[2].Value.Replace("''", "'")); if (Path.IsPathRooted(p) && !sheet.Book.Links.Contains(p)) sheet.Book.Links.Add(p); }
            }
        }
        public string Formula2 { get { return Formula; } set { Formula = value; } }
        public object Value2 { get { return value; } set { this.value = value; formula = null; sheet.Book.Saved = false; } }
        // Model Range.ClearContents used by the host for a genuinely blank
        // source. Formatting is intentionally retained, as it is in WPS.
        public void ClearContents() { value = null; formula = null; sheet.Book.Saved = false; }
        public string NumberFormat { get; set; } public int HorizontalAlignment { get; set; } public int VerticalAlignment { get; set; } public bool WrapText { get; set; }
        public int Orientation { get; set; } public int IndentLevel { get; set; } public bool ShrinkToFit { get; set; } public bool Locked { get; set; } public bool FormulaHidden { get; set; }
        public double RowHeight { get; set; } public double ColumnWidth { get; set; }
        public FakeFont Font { get; set; } public FakeFill Interior { get; set; } public FakeBorders Borders { get; set; }
        public void UnMerge() { var area = MigrationArea.Parse(Address); sheet.Merges.RemoveAll(m => area.Contains(MigrationArea.Parse(m))); }
        public void Merge() { if (!sheet.Merges.Contains(Address)) sheet.Merges.Add(Address); }
        public void Select() { }
        public FakeCell Clone(FakeSheet destination) { var c = (FakeCell)MemberwiseClone(); c.sheet = destination; c.Font = Font.Clone(); c.Interior = Interior.Clone(); c.Borders = Borders.Clone(); return c; }
    }
    public sealed class FakeFont
    {
        public string Name { get; set; } public double Size { get; set; } public bool Bold { get; set; } public bool Italic { get; set; } public int Underline { get; set; } public bool Strikethrough { get; set; } public int Color { get; set; }
        public FakeFont() { Name = "Arial"; Size = 11; } public FakeFont Clone() { return (FakeFont)MemberwiseClone(); }
    }
    public sealed class FakeFill
    {
        private int color, patternColor;
        public int Pattern { get; set; }
        public int Color { get { return color; } set { color = value; Pattern = 1; } }
        public int PatternColor { get { return patternColor; } set { patternColor = value; Pattern = 1; } }
        public FakeFill() { Pattern = -4142; }
        public FakeFill Clone() { return (FakeFill)MemberwiseClone(); }
    }
    public sealed class FakeBorder
    { public int LineStyle { get; set; } public int Color { get; set; } public int Weight { get; set; } public FakeBorder() { LineStyle = -4142; } public FakeBorder Clone() { return (FakeBorder)MemberwiseClone(); } }
    public sealed class FakeBorders
    {
        private Dictionary<int, FakeBorder> items = new Dictionary<int, FakeBorder>();
        public FakeBorder Item(int i) { FakeBorder b; if (!items.TryGetValue(i, out b)) items[i] = b = new FakeBorder(); return b; }
        public FakeBorders Clone() { var b = new FakeBorders(); foreach (var p in items) b.items[p.Key] = p.Value.Clone(); return b; }
    }
}
