using System.Threading;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

#if WPS_WORD
namespace WordUpdater
#else
namespace Excel_Iterate
#endif
{
    // The actual form classes are reused. Construction mode prevents host loads;
    // disabled business controls and null sessions prevent business execution.
    internal static class DialogPreviewFactory
    {
        [ThreadStatic]internal static bool IsConstructing;
        [ThreadStatic]internal static string RequestedId;
        [ThreadStatic]internal static Size RequestedSize;
        [ThreadStatic]internal static bool IsDefaultMeasurement;
        [ThreadStatic]internal static bool IsFormatConstruction;
        private static readonly ThreadLocal<HashSet<Form>> previewsState=new ThreadLocal<HashSet<Form>>(()=>new HashSet<Form>());
        private static HashSet<Form> previews {get {return previewsState.Value;}}
        private static readonly ThreadLocal<HashSet<Form>> formatPreviewsState=new ThreadLocal<HashSet<Form>>(()=>new HashSet<Form>());
        private static HashSet<Form> formatPreviews {get {return formatPreviewsState.Value;}}
        private static readonly ThreadLocal<Dictionary<Form,string>> titlesState=new ThreadLocal<Dictionary<Form,string>>(()=>new Dictionary<Form,string>());
        private static Dictionary<Form,string> titles {get {return titlesState.Value;}}
        internal static bool IsPreview(Form form){return previews.Contains(form);}
        internal static bool IsFormatRendering(Form form){return IsFormatConstruction||formatPreviews.Contains(form);}
        internal static string OriginalTitle(Form form){string title;return titles.TryGetValue(form,out title)?title:form.Text;}
        internal static Form Create(string id,Size desired)
        {return Construct(id,desired,false,false,false);}
        internal static Form CreateFormatRender(string id,Size desired)
        {return Construct(id,desired,false,false,true);}
        internal static Form CreateDefaultMeasure(string id,Size desired)
        {return Construct(id,desired,true,false,false);}
        internal static Form CreateProgramDefault(string id)
        {return Construct(id,DialogCatalog.Find(id).Default,false,true,false);}
        private static Form Construct(string id,Size desired,bool measure,bool normalSizing,bool format)
        {
            Ui.PreviewActivationGuard=IsFormatRendering;
            if(DialogCatalog.Find(id)==null||!DialogCatalog.Valid(id,desired))throw new ArgumentException("无效预览尺寸");
            bool oldConstructing=IsConstructing,oldMeasure=IsDefaultMeasurement,oldFormat=IsFormatConstruction;string oldId=RequestedId;Size oldSize=RequestedSize;
            Form form=null;IsConstructing=true;IsFormatConstruction=format;RequestedId=normalSizing?null:id;RequestedSize=desired;IsDefaultMeasurement=measure;
            try
            {
                switch(id)
                {
#if WPS_WORD
                    case "word.links":form=new LinkViewForm(null);break;
#else
                    case "trace.precedents":form=new TrackerForm(null,TrackerMode.Precedents);break;
                    case "trace.dependents":form=new TrackerForm(null,TrackerMode.Dependents);break;
                    case "trace.errorsources":form=new ErrorSourceForm(null,null,null);break;
                    case "columns.filter":form=new ColumnFilterForm(new List<ColumnChoice>{
                        new ColumnChoice {Number=1,Letter="A",Header="凭证号",Visible=true},new ColumnChoice {Number=2,Letter="B",Header="月份",Visible=true},
                        new ColumnChoice {Number=3,Letter="C",Header="借方数",Visible=true},new ColumnChoice {Number=4,Letter="D",Header="贷方数",Visible=true},new ColumnChoice {Number=5,Letter="E",Header="科目",Visible=true}},"示例表格：A1:E10");break;
                    case "formula.iferror":form=new FormulaOptionsForm(FormulaOperation.IFError);break;
                    case "formula.round":form=new FormulaOptionsForm(FormulaOperation.Round);break;
                    case "formula.add":form=new FormulaOptionsForm(FormulaOperation.Add);break;
                    case "formula.subtract":form=new FormulaOptionsForm(FormulaOperation.Subtract);break;
                    case "formula.multiply":form=new FormulaOptionsForm(FormulaOperation.Multiply);break;
                    case "formula.divide":form=new FormulaOptionsForm(FormulaOperation.Divide);break;
                    case "formula.custom":form=new CustomFunctionForm();break;
                    case "conversion.text":form=new CellConversionForm(CellConversionKind.TextFilter,"工作簿：示例.xlsx\r\n工作表：数据","J1:J5",new ConversionPreview("J1:J5",new ConversionCellState { Value="MIRCAI文化旅游公司2025!!",Display="MIRCAI文化旅游公司2025!!",NumberFormat="General" }));break;
                    case "conversion.batchInsertTrimDelete":form=new CellConversionForm(CellConversionKind.BatchInsertTrimDelete,"工作簿：示例.xlsx\r\n工作表：数据","A1:A5",new ConversionPreview("A1:A5",new ConversionCellState { Value="甲乙丙丁戊123",Display="甲乙丙丁戊123",NumberFormat="General" }));((TextBox)form.Controls.Find("batchInsertText",true)[0]).Text="示例";break;
                    case "formula.references":form=new CellConversionForm(CellConversionKind.ReferenceStyle,"工作簿：示例.xlsx\r\n工作表：数据","G11:G14,G16:G18",new ConversionPreview("G11:G14,G16:G18",new ConversionCellState { IsFormula=true,Formula="=H23+H24+H26",Value=100.0,Display="100",NumberFormat="General" }));break;
                    case "links.assistant":form=new LinkAssistantForm(null);break;
                    case "file.organize":form=new OrganizeFilesForm();break;
                    case "file.names":form=new FileNamesForm(null);break;
                    case "file.ocrprefixes":form=new OcrPrefixesForm(OcrCombined.ExampleRules(),null);break;
                    case "file.ocrfields":form=new OcrFieldsForm(new List<OcrFieldResult>{new OcrFieldResult {Prefix="磅单号：",Chosen="00123",Candidates=new List<string>{"00123","00456"}},new OcrFieldResult {Prefix="毛重：",Chosen="10kg",Candidates=new List<string>{"10kg"}}});break;
                    case "file.rename":form=new RenameFilesForm(null);break;
                    case "merge.headers":form=new MergeHeadersForm();break;
                    case "audit.batch.progress":form=new BatchProgressForm(new BatchProgressState("批量任务"));break;
                    case "batch.preset.name":form=new BatchPresetNameForm("company", "保存为预设", "示例年度预设");break;
                    case "batch.preset.manager":form=new BatchPresetManagerForm("company");break;
                    case "company.workbooks":form=new CompanyWorkbookForm();break;
                    case "summary.single":form=new SingleSummaryForm();break;
                    case "migration.data":form=new DataMigrationForm();break;
                    case "migration.preset.name":form=new MigrationNameDialog("保存为预设", "预设名称：", "示例预设");break;
                    case "migration.preset.manager":form=new MigrationPresetManager(null);break;
                    case "migration.preset.relocate":form=new MigrationRelocateDialog(new MigrationPreset(),new List<string>());break;
                    case "elimination.receivable":form=new EliminationForm(EliminationKind.Receivable);break;
                    case "elimination.revenuecost":form=new EliminationForm(EliminationKind.RevenueCost);break;
#endif
                    case "file.pdf.current":form=new PdfToolsForm("current");break;
                    case "file.pdf.convert":form=new PdfToolsForm("convert");break;
                    case "file.pdf.watermark":form=new PdfToolsForm("watermark");break;
                    case "file.pdf.split":form=new PdfToolsForm("split");break;
                    case "file.pdf.merge":form=new PdfToolsForm("merge");break;
                    case "agent.developer":form=new AgentDocumentForm(false);break;
                    case "agent.developerpassword":form=new DeveloperPasswordForm();break;
                    case "agent.user":form=new AgentDocumentForm(true);break;
                    case "agent.sort":form=new CommandCustomizationForm("sort",true);break;
                    case "agent.name":form=new CommandCustomizationForm("name",true);break;
                    case "agent.format":form=new DialogFormatEditor(true);break;
                    case "agent.notes":form=new RequirementsForm(true);break;
                    case "agent.dialogsizes":form=new DialogSizesForm(true);break;
                    case "message.ok":case "message.yesno":case "message.okcancel":
                        form=new UiMessageForm("这里显示操作结果或确认说明。\r\n尺寸预览不会执行实际业务。","提示与确认",id=="message.ok"?MessageBoxButtons.OK:id=="message.yesno"?MessageBoxButtons.YesNo:MessageBoxButtons.OKCancel,MessageBoxIcon.Information,MessageBoxDefaultButton.Button2);break;
                    default:throw new ArgumentException("未提供预览："+id);
                }
            }
            finally{IsConstructing=oldConstructing;RequestedId=oldId;RequestedSize=oldSize;IsDefaultMeasurement=oldMeasure;IsFormatConstruction=oldFormat;}
            previews.Add(form);if(format)formatPreviews.Add(form);titles[form]=form.Text;form.Disposed+=delegate {previews.Remove(form);formatPreviews.Remove(form);titles.Remove(form);};
            try
            {
                form.Text="尺寸预览 · "+form.Text;form.FormBorderStyle=FormBorderStyle.Sizable;form.MinimizeBox=false;form.MaximizeBox=false;
                if(!measure)Populate(form);MakeReadOnly(form);form.Shown+=delegate {MakeReadOnly(form);};return form;
            }
            catch{form.Dispose();throw;}
        }
        private static IEnumerable<Control> Descendants(Control root)
        {foreach(Control c in root.Controls){yield return c;foreach(Control child in Descendants(c))yield return child;}}
        private static void Populate(Form form)
        {
            foreach(Control c in Descendants(form))
            {
                var grid=c as DataGridView;
                if(grid!=null&&grid.Rows.Count==0&&grid.Columns.Count>0)
                {for(int i=1;i<=3;i++){var cells=new object[grid.Columns.Count];for(int j=0;j<cells.Length;j++)cells[j]=grid.Columns[j] is DataGridViewCheckBoxColumn?(object)true:"示例 "+i;grid.Rows.Add(cells);}}
                var label=c as Label;if(label!=null&&label.Text.StartsWith("正在",StringComparison.Ordinal))label.Text="尺寸预览 · 以下为示例内容";
            }
#if !WPS_WORD
            var files=form as FileToolForm;if(files!=null)files.FolderPath.Text="示例文件夹（不读取实际目录）";
            var organize=form as OrganizeFilesForm;if(organize!=null)organize.SetPreviewTypes();
#endif
            var pdf=form as PdfToolsForm;if(pdf!=null)pdf.SetPreview();
        }
        private static void MakeReadOnly(Form form)
        {
            foreach(Control c in Descendants(form))
            {
                var button=c as Button;if(button!=null&&button.Text!="关闭"&&button.Text!="取消")button.Enabled=false;
                var text=c as TextBoxBase;if(text!=null)text.ReadOnly=true;
                if(c is ComboBox||c is CheckBox||c is RadioButton)c.Enabled=false;
                var grid=c as DataGridView;if(grid!=null){grid.ReadOnly=true;grid.ContextMenuStrip=null;grid.AllowUserToAddRows=false;grid.AllowUserToDeleteRows=false;}
                if(c.ContextMenuStrip!=null)c.ContextMenuStrip=null;
            }
        }
    }
}

