using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Excel_Iterate
{
    internal sealed class FileNameEntry {
        internal string Name,Extension,Path,Feature="",Status="未识别";
        internal List<string> Candidates=new List<string>();
        internal bool Resolved;
        internal List<OcrSegment> OcrSegments;
        internal List<OcrFieldResult> Fields=new List<OcrFieldResult>();
        internal bool ManualFeature,ExplicitBlank;
    }
    internal sealed class FileNameScan {
        internal readonly List<FileNameEntry> Entries=new List<FileNameEntry>();
        internal readonly List<string> Warnings=new List<string>();
    }
    internal static class FileNameEngine {
        internal static bool IsChinese(int code) {return CellCharacterClassifier.IsChinese(code);}
        internal static string ExtractName(string name,int mode) {
            var result=new StringBuilder();for(int i=0;i<name.Length;i++) {
                int code=char.IsHighSurrogate(name[i])&&i+1<name.Length&&char.IsLowSurrogate(name[i+1])?char.ConvertToUtf32(name[i],name[i+1]):name[i];
                bool chinese=IsChinese(code);
                bool use=mode==0||mode==1&&chinese||mode==2&&!chinese||mode==3&&code>='0'&&code<='9'||mode==4&&(code>='a'&&code<='z'||code>='A'&&code<='Z');
                int count=code>0xffff?2:1;if(use)result.Append(name,i,count);i+=count-1;
            }return result.ToString();
        }
        internal static FileNameScan Scan(string folder,int depth,int mode,CancellationToken token) {
            if(depth<1||depth>256)throw new ArgumentException("扫描深度应为 1～256。");folder=System.IO.Path.GetFullPath(folder);
            if(!Directory.Exists(folder))throw new DirectoryNotFoundException("所选文件夹不存在。");
            if((File.GetAttributes(folder)&FileAttributes.ReparsePoint)!=0)throw new IOException("请选择实际文件夹；不扫描目录联接。");
            var scan=new FileNameScan();ScanDirectory(folder,depth,mode,scan,token);return scan;
        }
        private static void ScanDirectory(string folder,int remaining,int mode,FileNameScan scan,CancellationToken token) {
            token.ThrowIfCancellationRequested();try {
                foreach(string path in Directory.GetFiles(folder).OrderBy(p=>p,StringComparer.OrdinalIgnoreCase)) {token.ThrowIfCancellationRequested();scan.Entries.Add(new FileNameEntry {Path=path,Name=ExtractName(System.IO.Path.GetFileNameWithoutExtension(path),mode),Extension=System.IO.Path.GetExtension(path).TrimStart('.')});}
                if(remaining>1)foreach(string child in Directory.GetDirectories(folder).OrderBy(p=>p,StringComparer.OrdinalIgnoreCase)) {
                    token.ThrowIfCancellationRequested();try {if((File.GetAttributes(child)&FileAttributes.ReparsePoint)==0)ScanDirectory(child,remaining-1,mode,scan,token);else scan.Warnings.Add("跳过目录联接："+child);}catch(IOException){scan.Warnings.Add("无法访问子目录："+child);}catch(UnauthorizedAccessException){scan.Warnings.Add("无权访问子目录："+child);}
                }
            }catch(UnauthorizedAccessException){scan.Warnings.Add("无权访问目录："+folder);}catch(IOException){scan.Warnings.Add("无法访问目录："+folder);}
        }
        internal static object[,] Values(IList<FileNameEntry> entries,bool features) {
            var values=new object[entries.Count+1,features?4:3];values[0,0]="文件提取名称";values[0,1]="文件后缀";values[0,2]="文件路径";if(features)values[0,3]="文件OCR识别";
            for(int i=0;i<entries.Count;i++){values[i+1,0]=entries[i].Name;values[i+1,1]=entries[i].Extension;values[i+1,2]=entries[i].Path;if(features)values[i+1,3]=entries[i].Feature;}return values;
        }
    }
    internal sealed class FileNameWorksheet:IDisposable {
        internal const string ResultSheetName="OCR提取结果";
        private readonly object application;private object sheet;private WorksheetIdentity identity;
        internal string Description;
        internal FileNameWorksheet(object application) {
            this.application=application;
            try {sheet=((dynamic)application).ActiveSheet;identity=WorksheetIdentity.Capture(sheet);Description=identity.Name;}catch {Dispose();throw;}
        }
        private static void Release(object obj) {if(obj!=null&&Marshal.IsComObject(obj))Marshal.ReleaseComObject(obj);}
        private void Validate(int rows,int column,int width) {
            if(sheet==null||!identity.Matches(sheet))throw new InvalidOperationException("原目标工作表已变化，请关闭窗口后重新打开。");
            dynamic s=sheet;object rowCount=null,colCount=null;
            try {rowCount=s.Rows;colCount=s.Columns;if(rows<1||rows>(int)((dynamic)rowCount).Count||column<1||column+width-1>(int)((dynamic)colCount).Count)throw new InvalidOperationException("清单超出工作表行列范围。");if((bool)s.ProtectContents)throw new InvalidOperationException("目标工作表已保护，请先取消保护。");}finally {Release(rowCount);Release(colCount);}
        }
        internal void Preflight(int rows,int column,int width) {Validate(rows,column,width);}
        // Redirects later writes to a dedicated results sheet; the captured entry sheet only anchors the workbook.
        internal string EnsureResultSheet(string preferred) {
            if(sheet==null)throw new InvalidOperationException("原目标工作表已关闭，请关闭窗口后重新打开。");
            object workbook=null,sheets=null,target=null;
            try {
                dynamic original=sheet;try {workbook=original.Parent;}catch {workbook=null;}
                if(workbook==null)throw new InvalidOperationException("无法定位目标工作簿，请关闭窗口后重新打开。");
                try {sheets=((dynamic)workbook).Worksheets;}catch {sheets=null;}
                if(sheets==null)throw new InvalidOperationException("当前工作簿不支持新建工作表，请改用可编辑的工作簿。");
                int count=(int)((dynamic)sheets).Count;
                for(int index=1;index<=count;index++) {
                    object candidate=((dynamic)sheets).Item(index);
                    try {if(target==null&&string.Equals(Convert.ToString(((dynamic)candidate).Name),preferred,StringComparison.OrdinalIgnoreCase))target=candidate;}finally {if(target!=candidate)Release(candidate);}
                }
                if(target==null) {
                    object after=count>0?((dynamic)sheets).Item(count):null;
                    try {target=((dynamic)sheets).Add(Type.Missing,after);}
                    catch(Exception failure) {OcrDiagnostics.Record("file.names-sheet",failure);throw new InvalidOperationException("无法新建工作表，请检查工作簿是否处于保护、共享或只读状态。");}
                    finally {Release(after);}
                    ((dynamic)target).Name=preferred;
                } else ClearSheet(target);
                try {((dynamic)target).Activate();}catch {}
                Release(sheet);sheet=target;target=null;
                if(identity!=null)identity.Dispose();identity=WorksheetIdentity.Capture(sheet);Description=identity.Name;return Description;
            } finally {Release(target);Release(sheets);Release(workbook);}
        }
        private static void ClearSheet(object target) {
            try {((dynamic)target).AutoFilterMode=false;}catch {}
            object cells=null;try {cells=((dynamic)target).Cells;((dynamic)cells).Clear();}finally {Release(cells);}
        }
        internal bool Write(IList<FileNameEntry> entries,int column,bool features,Func<bool> confirm) {
            object[,] values=FileNameEngine.Values(entries,features);int rows=values.GetLength(0),width=values.GetLength(1);Validate(rows,column,width);
            object cells=null,first=null,last=null,range=null,header=null;bool started=false,addedFilter=false,hadFilter=false;
            object formulas=null,oldValues=null;var hasFormula=new bool[rows,width];var hiddenRows=new bool[rows];var formats=new object[rows,width];var wrap=new object[rows,width];var columnWidths=new object[width];var rowHeights=new object[rows];var bold=new object[width];var align=new object[rows,width];var vAlign=new object[rows,width];var headColor=new object[width];var headFill=new object[width];
            try {
                dynamic s=sheet;hadFilter=(bool)s.AutoFilterMode||(bool)s.FilterMode;
                object tables=null;try {tables=s.ListObjects;int count=(int)((dynamic)tables).Count;for(int t=1;t<=count;t++){object table=((dynamic)tables).Item(t);try {hadFilter|=(bool)((dynamic)table).ShowAutoFilter;}finally {Release(table);}}}catch(Microsoft.CSharp.RuntimeBinder.RuntimeBinderException){}finally {Release(tables);}
                cells=s.Cells;first=((dynamic)cells)[1,column];last=((dynamic)cells)[rows,column+width-1];range=s.Range[first,last];formulas=((dynamic)range).Formula;oldValues=((dynamic)range).Value2;
                for(int r=0;r<rows;r++){object cell=((dynamic)cells)[r+1,column],entireRow=null;try {entireRow=((dynamic)cell).EntireRow;hiddenRows[r]=Convert.ToBoolean(((dynamic)entireRow).Hidden);}finally {Release(entireRow);Release(cell);}}
                for(int r=0;r<rows;r++)for(int c=0;c<width;c++) {
                    object cell=((dynamic)cells)[r+1,column+c],font=null;
                    try {if(Convert.ToBoolean(((dynamic)cell).MergeCells)||Convert.ToBoolean(((dynamic)cell).HasArray))throw new InvalidOperationException("目标区域包含合并单元格或数组公式，请选择其他列。");formats[r,c]=((dynamic)cell).NumberFormat;hasFormula[r,c]=Convert.ToBoolean(((dynamic)cell).HasFormula);
                        wrap[r,c]=((dynamic)cell).WrapText;
                        if(c==0){object entireRow=((dynamic)cell).EntireRow;try {rowHeights[r]=WorkbookUi.ReadRowHeight(entireRow,hiddenRows[r]);}finally {Release(entireRow);}}
                        align[r,c]=((dynamic)cell).HorizontalAlignment;vAlign[r,c]=((dynamic)cell).VerticalAlignment;
                        if(r==0){object entireColumn=((dynamic)cell).EntireColumn;try {columnWidths[c]=((dynamic)entireColumn).ColumnWidth;}finally {Release(entireColumn);}font=((dynamic)cell).Font;bold[c]=((dynamic)font).Bold;headColor[c]=((dynamic)font).Color;object fill=((dynamic)cell).Interior;try {headFill[c]=((dynamic)fill).Color;}finally {Release(fill);}}
                    }finally {Release(font);Release(cell);}
                }
                bool occupied=false;var array=formulas as Array;if(array!=null){foreach(var v in array)if(v!=null&&Convert.ToString(v)!=""){occupied=true;break;}}else occupied=formulas!=null&&Convert.ToString(formulas)!="";
                if(occupied&&!confirm())return false;Validate(rows,column,width);
                started=true;((dynamic)range).NumberFormat="@";((dynamic)range).Value2=values;
                // WPS bulk setters skip filtered-out rows. Single-cell setters preserve
                // those rows without clearing the existing filter or its criteria.
                for(int r=0;r<rows;r++)if(hiddenRows[r])for(int c=0;c<width;c++){object cell=((dynamic)cells)[r+1,column+c];try{((dynamic)cell).NumberFormat="@";((dynamic)cell).Value2=values[r,c];}finally {Release(cell);}}
                object written=((dynamic)range).Value2;for(int r=0;r<rows;r++)for(int c=0;c<width;c++)if(Convert.ToString(At(written,r,c))!=Convert.ToString(values[r,c]))throw new InvalidOperationException("工作表未完整接收清单，已尝试还原，请核对目标区域。");
                WorkbookUi.StyleHeader(sheet, column, width);
                WorkbookUi.Center(range);
                if(!hadFilter){addedFilter=true;((dynamic)range).AutoFilter();}
                WorkbookUi.FitOutput(sheet, column, width, rows);
            }catch(Exception failure) {
                var errors=new List<string>();
                if(started) {
                    if(addedFilter)RestoreAttempt(()=>((dynamic)sheet).AutoFilterMode=false,"筛选状态",errors);
                    for(int r=0;r<rows;r++)for(int c=0;c<width;c++) {
                        int rr=r,cc=c;string address=CellAddress.Name(r+1,column+c);
                        RestoreAttempt(delegate {
                            object cell=((dynamic)cells)[rr+1,column+cc],font=null,fill=null;
                            try {
                                RestoreAttempt(()=>((dynamic)cell).NumberFormat=formats[rr,cc],address+"数字格式",errors);
                                RestoreAttempt(delegate {if(hasFormula[rr,cc])((dynamic)cell).Formula=At(formulas,rr,cc);else ((dynamic)cell).Value2=At(oldValues,rr,cc);},address+"内容",errors);
                                RestoreAttempt(()=>((dynamic)cell).HorizontalAlignment=align[rr,cc],address+"水平对齐",errors);
                                RestoreAttempt(()=>((dynamic)cell).VerticalAlignment=vAlign[rr,cc],address+"垂直对齐",errors);
                                RestoreAttempt(()=>((dynamic)cell).WrapText=wrap[rr,cc],address+"换行",errors);
                                if(rr==0){font=((dynamic)cell).Font;fill=((dynamic)cell).Interior;
                                    RestoreAttempt(()=>((dynamic)font).Bold=bold[cc],address+"粗体",errors);
                                    RestoreAttempt(()=>((dynamic)font).Color=headColor[cc],address+"字色",errors);
                                    RestoreAttempt(()=>((dynamic)fill).Color=headFill[cc],address+"底色",errors);
                                }
                            }finally {Release(fill);Release(font);Release(cell);}
                        },address+"引用",errors);
                    }
                    for(int c=0;c<width;c++){int cc=c;RestoreAttempt(delegate {object cell=((dynamic)cells)[1,column+cc],entireColumn=null;try {entireColumn=((dynamic)cell).EntireColumn;((dynamic)entireColumn).ColumnWidth=columnWidths[cc];}finally {Release(entireColumn);Release(cell);}},"第"+(column+c)+"列宽",errors);}
                    for(int r=0;r<rows;r++){int rr=r;RestoreAttempt(delegate {object cell=((dynamic)cells)[rr+1,column],entireRow=null;try {entireRow=((dynamic)cell).EntireRow;
                        RestoreAttempt(()=>((dynamic)entireRow).RowHeight=rowHeights[rr],"第"+(rr+1)+"行高",errors);
                        RestoreAttempt(()=>((dynamic)entireRow).Hidden=hiddenRows[rr],"第"+(rr+1)+"行隐藏状态",errors);
                    }finally {Release(entireRow);Release(cell);}},"第"+(r+1)+"行引用",errors);}
                }
                if(errors.Count!=0)throw new InvalidOperationException("清单写入失败，自动还原未能完成。请立即核对目标区域："+string.Join("；",errors)+"；原始失败："+failure.Message,failure);
                throw;
            }finally {Release(header);Release(range);Release(last);Release(first);Release(cells);}return true;
        }
        private static void RestoreAttempt(Action action,string label,List<string> errors) {try {action();}catch(Exception ex){errors.Add(label+"："+ex.Message);}}
        private static object At(object values,int row,int column) {var array=values as Array;return array==null?values:array.GetValue(row+array.GetLowerBound(0),column+array.GetLowerBound(1));}
        public void Dispose() {if(identity!=null){identity.Dispose();identity=null;}Release(sheet);sheet=null;}
    }
    internal sealed partial class FileNamesForm:Form,IDialogDefaultLayout {
        private readonly FileNameWorksheet target;
        private readonly TextBox folder=new TextBox {Dock=DockStyle.Fill,Name="folderInput"},prefix=new TextBox {Width=120,Name="ocrPrefix"},api=new TextBox {Width=150,Name="apiKey"},secret=new TextBox {Width=150,UseSystemPasswordChar=true,Name="secretKey"},pdfPage=new TextBox {Width=92,Name="pdfPageInput"};
        private readonly ComboBox depth=new ComboBox {DropDownStyle=ComboBoxStyle.DropDown,Width=75,Name="scanDepth"},column=new ComboBox {DropDownStyle=ComboBoxStyle.DropDown,Width=90,Name="startColumn"};
        private readonly ComboBox extraction=new ComboBox {DropDownStyle=ComboBoxStyle.DropDownList,Width=110,Name="nameExtraction"},provider=new ComboBox {DropDownStyle=ComboBoxStyle.DropDownList,Width=130,Name="ocrProvider"},precision=new ComboBox {DropDownStyle=ComboBoxStyle.DropDownList,Width=170,Name="cloudPrecision"};
        private readonly CheckBox features=new CheckBox {Text="是否开启文件OCR识别",AutoSize=true,Name="useOcr",Margin=new Padding(12,8,6,6)},remember=new CheckBox {Text="记住凭据（本机加密）",AutoSize=true,Name="rememberCredentials"};
        private readonly ResultGrid grid=new ResultGrid {Dock=DockStyle.Fill,AllowUserToAddRows=false,AllowUserToDeleteRows=false,RowHeadersVisible=false,AutoSizeColumnsMode=DataGridViewAutoSizeColumnsMode.None};
        private readonly Label status=new Label {AutoSize=true,Name="operationStatus",Text="选择文件夹并扫描，预览后再写入。"};
        private readonly TableLayoutPanel options=new TableLayoutPanel {Dock=DockStyle.Top,AutoSize=true,AutoSizeMode=AutoSizeMode.GrowAndShrink,ColumnCount=1,RowCount=6,Name="fileOptions"};
        private readonly List<Font> layoutFonts=new List<Font>();
        private readonly Label apiLabel=LabelFor("API Key"),secretLabel=LabelFor("Secret Key");
        private readonly string[][] credentialDrafts={new[]{"",""},new[]{"",""}};
        private readonly bool[] rememberDrafts=new bool[2];
        private int credentialProvider=-1;
        private Label explanation;
        private readonly RoundedButton scan=Ui.Button("扫描并预览",120),retry=Ui.Button("重试未完成",120),write=Ui.Button("写入工作表",120),cancel=Ui.Button("取消识别",110),jump=Ui.Button("跳转到API平台",120);
        private CancellationTokenSource cancellation;private IFileOcr ocr;private List<FileNameEntry> entries=new List<FileNameEntry>();private bool busy,closeRequested,scannedFeatures;private readonly bool preview;
        private readonly int uiThread=Thread.CurrentThread.ManagedThreadId;
        private readonly CancellationTokenSource lifetime=new CancellationTokenSource();
        internal Func<string,string,OcrMode,IFileOcr> OcrFactory;
        internal Func<string,bool> ConfirmUpload;
        private OcrPrefixSnapshot prefixRules;
        private readonly RoundedButton setPrefixes=Ui.Button("设置前缀",108);
        internal FileNamesForm(FileNameWorksheet target) {
            this.target=target;preview=target==null||DialogPreviewFactory.IsConstructing;Text="文件提取名称";DialogSizing.Prepare(this,"file.names","document",Ui.Blue);
            for(int i=1;i<=256;i++)depth.Items.Add(i.ToString());for(int i=1;i<=32;i++)column.Items.Add(i.ToString());depth.Text="1";column.Text="1";pdfPage.Text="1";
            extraction.Items.AddRange(new object[]{"全部字符","仅中文","非中文","仅数字","仅字母"});extraction.SelectedIndex=0;provider.Items.AddRange(new object[]{"百度云OCR","阿里云OCR"});provider.SelectedIndex=0;
            var root=new TableLayoutPanel {Dock=DockStyle.Fill,ColumnCount=1,RowCount=4,Padding=new Padding(14)};root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));root.RowStyles.Add(new RowStyle(SizeType.AutoSize));root.RowStyles.Add(new RowStyle(SizeType.Percent,100));root.RowStyles.Add(new RowStyle(SizeType.AutoSize));root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            options.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));
            for(int i=0;i<6;i++)options.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var pathRow=new TableLayoutPanel {Dock=DockStyle.Top,AutoSize=true,AutoSizeMode=AutoSizeMode.GrowAndShrink,ColumnCount=3,Name="folderRow"};pathRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));pathRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));pathRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));pathRow.Controls.Add(LabelFor("文件夹"),0,0);pathRow.Controls.Add(folder,1,0);CommandCustomizationForm.AddButton(pathRow,"浏览",()=>{string selected=HostDialogService.Folder(this,"选择需要扫描的文件夹",true);if(selected!=null)folder.Text=selected;});options.Controls.Add(pathRow,0,0);
            var heading=LabelFor("自定义配置：");heading.Name="configurationHeading";heading.ForeColor=Ui.Blue;heading.Font=BoldFont();options.Controls.Add(heading,0,1);
            var parameters=OptionRow("parametersRow");parameters.Controls.Add(Field("深度",depth));parameters.Controls.Add(Field("起始列",column));parameters.Controls.Add(Field("保留字符数",extraction));parameters.Controls.Add(features);options.Controls.Add(parameters,0,2);
            prefix.ReadOnly=true;setPrefixes.Name="setOcrPrefixes";
            var featureRow=OptionRow("featureRow");featureRow.Controls.Add(Field("OCR前缀组合（按启用顺序）",prefix));featureRow.Controls.Add(setPrefixes);featureRow.Controls.Add(Field("PDF文件OCR统一页码设置",pdfPage));jump.AutoSize=true;jump.Name="apiPlatformJump";featureRow.Controls.Add(jump);options.Controls.Add(featureRow,0,3);
            var recognition=OptionRow("recognitionRow");recognition.Controls.Add(Gripped(provider,null));recognition.Controls.Add(Gripped(precision,null));apiLabel.Name="credentialKeyLabel";secretLabel.Name="credentialSecretLabel";recognition.Controls.Add(Field(apiLabel,api));recognition.Controls.Add(Field(secretLabel,secret));recognition.Controls.Add(remember);CommandCustomizationForm.AddButton(recognition,"清除API凭据",()=>{OcrCredentials.Clear(provider.SelectedIndex==1);api.Clear();secret.Clear();remember.Checked=false;});options.Controls.Add(recognition,0,4);
            foreach(Control row in new Control[]{parameters,featureRow,recognition})BoldTree(row);
            explanation=LabelFor("");explanation.Name="targetSummary";explanation.ForeColor=Color.Firebrick;explanation.Font=BoldFont();options.Controls.Add(explanation,0,5);root.Controls.Add(options);
            grid.ColumnFitOverride=FitResultColumns;
            Ui.StyleGrid(grid);grid.Columns.Add("name","文件提取名称");grid.Columns.Add("extension","文件后缀");grid.Columns.Add("path","文件路径");grid.Columns.Add("feature","文件OCR识别");grid.Columns.Add(new DataGridViewButtonColumn {Name="candidates",HeaderText="候选",Text="核对 / 留空",UseColumnTextForButtonValue=false});grid.Columns.Add("status","识别状态");foreach(DataGridViewColumn c in grid.Columns)c.ReadOnly=true;root.Controls.Add(grid);
            root.Controls.Add(status);var bar=new FlowLayoutPanel {Dock=DockStyle.Fill,AutoSize=true};bar.Controls.Add(scan);bar.Controls.Add(retry);bar.Controls.Add(write);bar.Controls.Add(cancel);CommandCustomizationForm.AddButton(bar,"关闭",()=>Close());root.Controls.Add(bar);Controls.Add(root);
            CommandCustomizationForm.AddButton(bar,"复制诊断",()=>{Clipboard.SetText(string.IsNullOrEmpty(UiDiagnostics.Text)?"当前窗口没有记录界面异常。":UiDiagnostics.Text);});
            ConfigureFileNamesLayout(root);
            grid.PaintFailure=delegate {status.Text="结果显示失败，识别结果已保留；请复制诊断后重试。";};
            scan.Click+=async delegate {await Run(false);};retry.Click+=async delegate {await Run(true);};write.Click+=delegate {Write();};cancel.Click+=delegate {if(cancellation!=null)cancellation.Cancel();};jump.Click+=delegate {ShowApiPlatforms();};
            features.CheckedChanged+=delegate {Toggle();};provider.SelectedIndexChanged+=delegate {SwitchProvider();Toggle();};
            provider.FontChanged+=delegate {MeasureChoice(provider);};precision.FontChanged+=delegate {MeasureChoice(precision);};extraction.FontChanged+=delegate {MeasureChoice(extraction);};Load+=delegate {MeasureChoice(provider);MeasureChoice(precision);MeasureChoice(extraction);};
            setPrefixes.Click+=delegate {EditPrefixes();};
            grid.CellClick+=delegate(object sender,DataGridViewCellEventArgs e){if(e.RowIndex<0||e.ColumnIndex<0||grid.Columns[e.ColumnIndex].Name!="candidates"||busy||!scannedFeatures)return;ReviewFields((FileNameEntry)grid.Rows[e.RowIndex].Tag);};
            grid.CellEndEdit+=delegate(object sender,DataGridViewCellEventArgs e){if(e.RowIndex<0||grid.Columns[e.ColumnIndex].Name!="feature")return;var entry=(FileNameEntry)grid.Rows[e.RowIndex].Tag;entry.Feature=Convert.ToString(grid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value);entry.Resolved=true;entry.ManualFeature=true;entry.ExplicitBlank=string.IsNullOrWhiteSpace(entry.Feature);entry.Status=entry.ExplicitBlank?"已明确留空":"手工结果";};
            FormClosing+=delegate(object sender,FormClosingEventArgs e){if(busy){closeRequested=true;e.Cancel=true;cancellation.Cancel();status.Text="正在取消，请稍候…";}};
            FormClosed+=delegate {if(cancellation!=null)cancellation.Dispose();if(ocr!=null)ocr.Dispose();if(target!=null)target.Dispose();};
            Disposed+=delegate {lifetime.Cancel();foreach(var font in layoutFonts)font.Dispose();};
            prefixRules=preview?new OcrPrefixSnapshot {Rules=OcrCombined.ExampleRules()}:OcrPrefixStore.Read();SetPrefixRules(prefixRules.Rules);
            for(int service=0;service<2;service++){if(!preview)credentialDrafts[service]=OcrCredentials.Read(service==1);rememberDrafts[service]=credentialDrafts[service][0]!=""&&credentialDrafts[service][1]!="";}SwitchProvider();Toggle();
            if(prefixRules.Warning.Length>0)status.Text=prefixRules.Warning;
        }
        private void SwitchProvider() {
            if(credentialProvider>=0){credentialDrafts[credentialProvider]=new[]{api.Text,secret.Text};rememberDrafts[credentialProvider]=remember.Checked;}
            credentialProvider=Math.Max(0,provider.SelectedIndex);bool aliyun=credentialProvider==1;
            api.Text=credentialDrafts[credentialProvider][0];secret.Text=credentialDrafts[credentialProvider][1];remember.Checked=rememberDrafts[credentialProvider];
            apiLabel.Text=aliyun?"AccessKey ID":"API Key";secretLabel.Text=aliyun?"AccessKey Secret":"Secret Key";
            precision.Items.Clear();precision.Items.AddRange(aliyun?new object[]{"全文识别高精版"}:OcrModes.BaiduNames.Cast<object>().ToArray());precision.SelectedIndex=0;
            explanation.Text="云端识别会上传图片或 PDF 指定页并使用你的账号配额，可能产生费用。"+(aliyun?"  阿里云不支持 PDF 文件，PDF 行将跳过。":"  PDF 页码可填 0（不扫描）、ALL（全部页）、单数字或 起始-结束 范围。")+"  结果写入工作表「"+FileNameWorksheet.ResultSheetName+"」（同名覆盖）；不自动保存。";
            MeasureChoice(provider);MeasureChoice(precision);
        }
        private static void MeasureChoice(ComboBox box) {
            int width=0;foreach(var item in box.Items)width=Math.Max(width,TextRenderer.MeasureText(Convert.ToString(item),box.Font,Size.Empty,TextFormatFlags.NoPadding).Width);
            float scale;using(var graphics=box.CreateGraphics())scale=graphics.DpiX/96f;
            int measured=Math.Max((int)(110*scale),width+(int)(42*scale));if(!DialogControlSizing.HasWidth(box))box.Width=measured;box.DropDownWidth=Math.Max(box.Width,measured);
        }
        private Font BoldFont() {var font=new Font(Font,FontStyle.Bold);layoutFonts.Add(font);return font;}
        private static Label LabelFor(string text) {return new Label {Text=text,AutoSize=true,Margin=new Padding(3,8,6,6)};}
        private static FlowLayoutPanel OptionRow(string name) {return new FlowLayoutPanel {Dock=DockStyle.Top,AutoSize=true,AutoSizeMode=AutoSizeMode.GrowAndShrink,WrapContents=true,Name=name,Margin=new Padding(0)};}
        // Inputs get a trailing grip so users can widen them and reveal long values.
        private static FlowLayoutPanel Gripped(Control input,Control leading) {return Ui.Grip(input,leading);}
        private static Control Field(string label,Control input){return Field(LabelFor(label),input);}
        private static Control Field(Label label,Control input) {return Gripped(input,label);}
        private void BoldTree(Control root) {root.Font=BoldFont();foreach(Control child in root.Controls)BoldTree(child);}
        private static int Number(ComboBox box,int maximum,string label) {int value;if(!int.TryParse(box.Text,out value)||value<1||value>maximum){box.Focus();throw new InvalidOperationException(label+"必须为 1 至 "+maximum+" 的整数。");}return value;}
        private static int Number(TextBox box,int maximum,string label) {int value;if(!int.TryParse(box.Text,out value)||value<1||value>maximum){box.Focus();throw new InvalidOperationException(label+"必须为 1 至 "+maximum+" 的整数。");}return value;}
        // 上传统计：PDF 按规格求解实际待识别页数；图片仍按其帧数。
        private static int PdfPageUploadCount(string path,PdfPageSpec spec) {if(spec==null||spec.None)return 0;return spec.PagesFor(PdfPages.Count(path)).Count;}
        public Size MeasureDefaultClient() {return new Size(1200,760);}
        internal void CancelForShutdown() {closeRequested=true;if(cancellation!=null)cancellation.Cancel();Close();}
        private void Toggle() {bool cloud=features.Checked;prefix.Enabled=setPrefixes.Enabled=features.Checked;pdfPage.Enabled=features.Checked;provider.Enabled=features.Checked;precision.Enabled=cloud;api.Enabled=secret.Enabled=remember.Enabled=cloud;retry.Enabled=!busy&&features.Checked&&entries.Count>0;write.Enabled=!busy&&entries.Count>0;cancel.Enabled=busy;}
        internal void SetPrefixRules(IList<OcrPrefixRule> rules)
        {
            OcrPrefixStore.Validate(rules,false);
            prefixRules.Rules=rules.Select(r=>r.Copy()).ToList();prefix.Text=string.Join(" → ",prefixRules.Rules.Where(r=>r.Enabled).Select(r=>r.Text));
            foreach(var entry in entries)
            {
                if(prefixRules.Rules.Any(r=>r.Enabled))OcrCombined.Apply(entry,prefixRules.Rules);
                else
                {
                    entry.Fields.Clear();entry.Candidates.Clear();
                    if(entry.ExplicitBlank){entry.Feature="";entry.Resolved=true;entry.Status="已明确留空";}
                    else if(entry.ManualFeature){entry.Resolved=true;entry.Status="手工结果，请核对组合规则";}
                    else {entry.Feature="";entry.Resolved=false;entry.Status="尚未启用前缀";}
                }
            }
            ShowEntries();
        }
        private void EditPrefixes()
        {
            if(preview||busy)return;
            var previous=prefixRules;OcrPrefixSnapshot saved=null;
            var answer=HostDialogService.ShowIsolatedResult(this,()=>new OcrPrefixesForm(previous.Rules,delegate(List<OcrPrefixRule> rules){
                try {saved=OcrPrefixStore.Save(previous,rules);return saved.Writable?"":saved.Warning;}catch(Exception error){return error.Message;}
            }));
            if(answer==DialogResult.OK&&saved!=null){prefixRules=saved;SetPrefixRules(saved.Rules);status.Text="前缀规则已保存；已识别文件按缓存重新组合，未额外上传。";}
        }
        private void ReviewFields(FileNameEntry entry)
        {
            if(preview||busy)return;
            if(!prefixRules.Rules.Any(r=>r.Enabled)){status.Text="请先设置并启用字符前缀。";return;}
            if(entry.Fields.Count==0)entry.Fields=prefixRules.Rules.Where(r=>r.Enabled).Select(r=>new OcrFieldResult {RuleId=r.Id,Prefix=r.Text}).ToList();
            OcrFieldsForm review=null;
            var answer=HostDialogService.ShowIsolatedResult(this,()=>review=new OcrFieldsForm(entry.Fields));
            if(answer!=DialogResult.OK||review==null)return;
            entry.ManualFeature=false;entry.ExplicitBlank=review.Blank;
            if(review.Blank){entry.Feature="";entry.Resolved=true;entry.Status="已明确留空";}
            else {entry.Fields=review.Result;OcrCombined.Compose(entry);}
            ShowEntries();
        }
        internal bool IsWorking {get{return busy;}}
        // Widths dragged at runtime persist into the 自定配置 store; preview windows never save.
        internal void PersistWidth(Control control,int width96) {if(!preview)DialogFormatting.SaveControlWidth(this,control,width96);}
        private void ShowApiPlatforms() {HostDialogService.ShowIsolated(this,()=>new ApiPlatformDialog());}
        private void SetBusy(bool value) {busy=value;options.Enabled=!value;scan.Enabled=!value;grid.ReadOnly=value;if(!value)foreach(DataGridViewColumn c in grid.Columns)c.ReadOnly=c.Name!="feature"||!scannedFeatures;Toggle();}
        private void ShowEntries() {
            grid.Rows.Clear();foreach(var e in entries){int i=grid.Rows.Add(e.Name,e.Extension,e.Path,scannedFeatures?e.Feature:"",scannedFeatures?"核对 / 留空":"—",scannedFeatures?e.Status:"未启用OCR");grid.Rows[i].Tag=e;
                if(!scannedFeatures)grid.Rows[i].Cells["candidates"].Style.ForeColor=SystemColors.GrayText;}
            foreach(DataGridViewColumn c in grid.Columns)c.Visible=true;grid.Columns["feature"].ReadOnly=busy||!scannedFeatures;Toggle();
        }
        private Task<T> OnUi<T>(Func<T> action) {
            var completion=new TaskCompletionSource<T>();
            var registration=lifetime.Token.Register(()=>completion.TrySetCanceled());
            Action invoke=delegate {if(IsDisposed||Disposing){completion.TrySetCanceled();return;}try{completion.TrySetResult(action());}catch(Exception error){completion.TrySetException(error);}finally{registration.Dispose();}};
            if(Thread.CurrentThread.ManagedThreadId==uiThread)invoke();
            else try{BeginInvoke((MethodInvoker)delegate{invoke();});}catch(InvalidOperationException){registration.Dispose();completion.TrySetCanceled();}
            return completion.Task;
        }
        private Task OnUi(Action action){return OnUi(()=>{action();return true;});}
        private async Task Run(bool retryOnly) {
            if(preview||busy)return;
            bool useOcr=features.Checked,aliyun=provider.SelectedIndex==1;
            OcrMode mode=aliyun?OcrMode.AliyunAdvanced:(OcrMode)precision.SelectedIndex;
            string service=aliyun?"阿里云":"百度云",variant=precision.Text,root=folder.Text,key=api.Text.Trim(),password=secret.Text.Trim();
            var rules=prefixRules.Rules.Select(r=>r.Copy()).ToList();
            int targetColumn,level,nameMode=extraction.SelectedIndex;bool keep=remember.Checked;
            PdfPageSpec pdfPageSpec=null;
            try {
                if(useOcr)OcrPrefixStore.Validate(rules,true);
                if(useOcr&&(key.Length==0||password.Length==0)){(key.Length==0?api:secret).Focus();throw new InvalidOperationException(aliyun?"请填写阿里云的 AccessKey ID 和 AccessKey Secret。":"请填写百度云应用的 API Key 和 Secret Key。它们用于验证身份，并使用你账号下的 OCR 权限与配额；未填写无法调用百度云识别。");}
                targetColumn=Number(column,16384,"起始列");level=Number(depth,256,"深度");if(useOcr)pdfPageSpec=PdfPageSpec.Parse(pdfPage.Text);
            }catch(Exception error){status.Text=error.Message;return;}
            if(!retryOnly)BeginColumnFitCycle();
            if(cancellation!=null)cancellation.Dispose();cancellation=new CancellationTokenSource();var token=cancellation.Token;
            SetBusy(true);status.Text="正在扫描…";Action ending=null;
            try {
                if(!retryOnly){
                    var result=await Task.Run(()=>FileNameEngine.Scan(root,level,nameMode,token),token).ConfigureAwait(false);
                    await OnUi(()=>{entries=result.Entries;scannedFeatures=useOcr;ShowEntries();FitResultColumns();status.Text="扫描到 "+entries.Count+" 个文件";
                        if(result.Warnings.Count>0)UiDialog.Show(this,string.Join(Environment.NewLine,result.Warnings.Take(20)),"扫描提示",MessageBoxButtons.OK,MessageBoxIcon.Warning);
                    }).ConfigureAwait(false);
                }
                await OnUi(()=>target.Preflight(entries.Count+1,targetColumn,useOcr?4:3)).ConfigureAwait(false);
                if(useOcr){
                    if(keep)OcrCredentials.Save(aliyun,key,password);else OcrCredentials.Clear(aliyun);
                    ocr=OcrFactory!=null?OcrFactory(key,password,mode):aliyun?(IFileOcr)new AliyunFileOcr(key,password):new BaiduFileOcr(key,password,mode);
                    var work=entries.Where(e=>!retryOnly||(!e.Resolved&&e.Feature=="")).ToList();
                    int pdfCount=await Task.Run(()=>work.Count(e=>OcrImages.IsPdf(e.Path)),token).ConfigureAwait(false);
                    if(aliyun&&pdfCount>0)await OnUi(()=>UiDialog.Show(this,"扫描到 "+pdfCount+" 个 PDF 文件。阿里云全文识别高精版不支持 PDF，这些行将跳过；如需识别 PDF，请改用百度云OCR。","文件提取名称",MessageBoxButtons.OK,MessageBoxIcon.Warning)).ConfigureAwait(false);
                    int count=await Task.Run(()=>work.Sum(e=>{
                        if(OcrImages.IsPdf(e.Path))return aliyun?0:PdfPageUploadCount(e.Path,pdfPageSpec);
                        if(!OcrImages.IsImage(e.Path))return 0;
                        try{return OcrImages.Pages(e.Path);}catch{return 0;}
                    }),token).ConfigureAwait(false);
                    string message="将上传 "+count+" 张图片或页面至"+service+"（"+variant+"），可能使用配额或产生费用。";
                    if(pdfCount>0&&!aliyun&&!pdfPageSpec.None)message+="  PDF 文件统一识别第 "+pdfPageSpec.Description+" 页。";
                    message+="是否开始？";
                    bool approved=await OnUi(()=>ConfirmUpload!=null?ConfirmUpload(message):UiDialog.Show(this,message,service+"OCR",MessageBoxButtons.YesNo,MessageBoxIcon.Question)==DialogResult.Yes).ConfigureAwait(false);
                    if(!approved)throw new OperationCanceledException();
                    string temporary=System.IO.Path.Combine(System.IO.Path.GetTempPath(),"ExcelIterate-Ocr-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(temporary);
                    bool paused=false;
                    try{for(int i=0;i<work.Count;i++){
                        token.ThrowIfCancellationRequested();var entry=work[i];
                        if(!OcrImages.Supported(entry.Path)){entry.Status="不支持识别，保留文件行";continue;}
                        bool pdf=OcrImages.IsPdf(entry.Path);
                        if(pdf&&aliyun){entry.Status="阿里云不支持 PDF，请改用百度云";continue;}
                        int index=i+1;await OnUi(()=>status.Text="正在识别 "+index+" / "+work.Count+"："+System.IO.Path.GetFileName(entry.Path)).ConfigureAwait(false);
                        try{
                            var all=new List<OcrSegment>();
                            if(pdf){
                                int pdfTotal=await Task.Run(()=>PdfPages.Count(entry.Path),token).ConfigureAwait(false);
                                foreach(int pageNum in pdfPageSpec.PagesFor(pdfTotal)){
                                    token.ThrowIfCancellationRequested();
                                    all.AddRange(await ocr.RecognizeAsync(entry.Path,pageNum,token).ConfigureAwait(false));
                                }
                            }
                            else {int pages=await Task.Run(()=>OcrImages.Pages(entry.Path),token).ConfigureAwait(false);
                                for(int page=1;page<=pages;page++){
                                    int currentPage=page;token.ThrowIfCancellationRequested();
                                    string image=await Task.Run(()=>OcrImages.Prepare(entry.Path,currentPage,mode,temporary),token).ConfigureAwait(false);
                                    try{all.AddRange(await ocr.RecognizeAsync(image,page,token).ConfigureAwait(false));}finally{if(File.Exists(image))File.Delete(image);}
                                }
                            }
                            entry.OcrSegments=all;OcrCombined.Apply(entry,rules);
                        }catch(OperationCanceledException){throw;}
                        catch(OcrFailure error){entry.Status=error.Message;if(error.Pause)paused=true;}
                        catch(Exception error){OcrDiagnostics.Record("cloud-file",error);entry.Status="文件处理失败（"+error.GetType().Name+"），请检查文件及本机诊断记录后重试。";}
                        if(paused){await OnUi(()=>status.Text=entry.Status+" 已暂停，可修改设置后重试未完成项。").ConfigureAwait(false);break;}
                    }}finally{if(Directory.Exists(temporary))Directory.Delete(temporary,true);}
                    await OnUi(()=>{scannedFeatures=true;ShowEntries();FitResultColumns();if(!paused&&!grid.PaintFailed)status.Text="识别结束。核对OCR识别结果后再写入；未完成项可切换方式重试。";}).ConfigureAwait(false);
                }
            }catch(OperationCanceledException){ending=()=>{status.Text="已取消，未写入工作表。";ShowEntries();FitResultColumns();};}
            catch(Exception error){UiDiagnostics.Record("file.names-workflow",error);ending=()=>{status.Text=error.Message;UiDialog.Show(this,error.Message,"文件提取名称",MessageBoxButtons.OK,MessageBoxIcon.Warning);};}
            finally{if(ocr!=null){ocr.Dispose();ocr=null;}}
            if(!lifetime.IsCancellationRequested){
                try{if(ending!=null)await OnUi(ending).ConfigureAwait(false);await OnUi(()=>{SetBusy(false);if(closeRequested)BeginInvoke((MethodInvoker)delegate{if(!IsDisposed)Close();});}).ConfigureAwait(false);}
                catch(OperationCanceledException){}
            }
        }
        private void Write() {
            if(preview||busy||entries.Count==0)return;grid.EndEdit();
            try {if(scannedFeatures&&entries.Any(e=>e.Candidates.Count>1&&!e.Resolved))throw new InvalidOperationException("请先选择所有多候选结果，或在候选菜单中明确留空。");
                string destination=target.EnsureResultSheet(FileNameWorksheet.ResultSheetName);
                bool completed=target.Write(entries,Number(column,16384,"起始列"),scannedFeatures,()=>UiDialog.Show("目标区域已有内容，是否覆盖？","文件提取名称",MessageBoxButtons.YesNo,MessageBoxIcon.Warning)==DialogResult.Yes);status.Text=completed?"已写入工作表「"+destination+"」，请核对；工作簿未自动保存。":"已取消写入，"+destination+" 保持原样。";
            }catch(Exception ex){UiDialog.Show(ex.Message,"文件提取名称",MessageBoxButtons.OK,MessageBoxIcon.Warning);}
        }
    }
    // Chooser window behind 跳转到API平台. Options and URLs come from ApiPlatforms
    // (配置器脚本 -> ApiPlatforms.g.cs); an empty URL renders a disabled row.
    internal sealed class ApiPlatformDialog:Form {
        internal ApiPlatformDialog() {
            Text="跳转到API平台";FormBorderStyle=FormBorderStyle.FixedDialog;MinimizeBox=MaximizeBox=false;ShowInTaskbar=false;StartPosition=FormStartPosition.CenterParent;AutoSize=true;AutoSizeMode=AutoSizeMode.GrowAndShrink;
            var layout=new FlowLayoutPanel {Dock=DockStyle.Fill,FlowDirection=FlowDirection.TopDown,WrapContents=false,AutoSize=true,Padding=new Padding(14)};
            string[] names=ApiPlatforms.Names,urls=ApiPlatforms.Urls;
            for(int i=0;i<names.Length;i++) {
                string url=i<urls.Length?urls[i]:"";
                var button=Ui.Button(names[i],320);
                if(string.IsNullOrWhiteSpace(url))button.Enabled=false;
                button.Click+=delegate {OpenPlatform(url);Close();};
                layout.Controls.Add(button);
            }
            var close=Ui.Button("关闭",320);close.Click+=delegate {Close();};layout.Controls.Add(close);
            Controls.Add(layout);
        }
        internal static void OpenPlatform(string url) {
            if(string.IsNullOrWhiteSpace(url))return;
            try {System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(url){UseShellExecute=true});}
            catch(Exception error){OcrDiagnostics.Record("api-platform",error);UiDialog.Show("无法打开浏览器："+error.Message,"跳转到API平台",MessageBoxButtons.OK,MessageBoxIcon.Warning);}
        }
    }
    public partial class Connect {
        private static FileNamesForm fileNames;

        internal static void CloseFileNames() {if(fileNames!=null&&!fileNames.IsDisposed)fileNames.CancelForShutdown();}
    }
}
