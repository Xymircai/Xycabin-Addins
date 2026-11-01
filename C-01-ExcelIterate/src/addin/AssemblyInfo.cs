using System;
using System.Reflection;
using System.Runtime.InteropServices;

[assembly: AssemblyTitle("WPS表格增强追踪加载项")]
[assembly: AssemblyDescription("WPS表格：追踪引用/从属——解析公式逐项展开被引用单元格详情，反向找出引用当前单元格的全部单元格，点击跳转并高亮")]
[assembly: AssemblyCompany("Xy")]
[assembly: AssemblyProduct("Excel-iterate")]
[assembly: AssemblyVersion("0.9.9.3")]
[assembly: AssemblyFileVersion("0.9.9.3")]
[assembly: ComVisible(false)]
[assembly: Guid("7A4B2E15-9D3F-4A8C-B6E1-2F5D7C9E1A43")]

[assembly: System.Runtime.CompilerServices.InternalsVisibleTo("FormulaProbe")]
[assembly: System.Runtime.CompilerServices.InternalsVisibleTo("FormulaInteractionProbe")]

[assembly: System.Runtime.CompilerServices.InternalsVisibleTo("WorkbookToolsProbe")]

[assembly: System.Runtime.CompilerServices.InternalsVisibleTo("FileOperationsProbe")]

[assembly: System.Runtime.CompilerServices.InternalsVisibleTo("Feature099Probe")]

[assembly: System.Runtime.CompilerServices.InternalsVisibleTo("HostDialogProbe")]

[assembly: System.Runtime.CompilerServices.InternalsVisibleTo("MergeHeadersProbe")]
[assembly: System.Runtime.CompilerServices.InternalsVisibleTo("SingleSummaryProbe")]
[assembly: System.Runtime.CompilerServices.InternalsVisibleTo("CompanyWorkbookProbe")]
[assembly: System.Runtime.CompilerServices.InternalsVisibleTo("SingleSummaryLiveProbe")]
[assembly: System.Runtime.CompilerServices.InternalsVisibleTo("MergeLiveProbe")]
[assembly: System.Runtime.CompilerServices.InternalsVisibleTo("TrackerProbe")]
[assembly: System.Runtime.CompilerServices.InternalsVisibleTo("EliminationProbe")]

[assembly: System.Runtime.CompilerServices.InternalsVisibleTo("RuntimeQualityProbe")]
[assembly: System.Runtime.CompilerServices.InternalsVisibleTo("DataMigrationProbe")]
[assembly: System.Runtime.CompilerServices.InternalsVisibleTo("DataMigrationLiveProbe")]
