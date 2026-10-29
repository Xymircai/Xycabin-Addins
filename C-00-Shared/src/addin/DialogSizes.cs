using System.Threading;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;
using System.Xml;

#if WPS_WORD
namespace WordUpdater
#else
namespace Excel_Iterate
#endif
{
    internal sealed class DialogDefinition
    {
        internal readonly string Id, Name, Entry;
        internal Size Default;
        internal readonly Size BaseDefault, Minimum;
        internal DialogDefinition(string id, string name, string entry, int w, int h, int mw, int mh)
        { Id=id; Name=name; Entry=entry; BaseDefault=Default=new Size(w,h); Minimum=new Size(mw,mh); }
    }
    // Stable identities are independent of translated captions and add-in versions.
    internal static class DialogCatalog
    {
        internal static readonly DialogDefinition[] All = new DialogDefinition[] {
#if WPS_WORD
            new DialogDefinition("word.links", "查看链接", "查看链接", 1180,720,820,560),
#else
            new DialogDefinition("trace.precedents", "追踪引用", "数据 / 追踪引用",1240,680,760,500),
            new DialogDefinition("trace.dependents", "追踪从属", "数据 / 追踪从属",1240,680,760,500),
            new DialogDefinition("trace.errorsources", "错误来源", "数据 / 追踪引用 / 显示错误来源",1100,620,760,460),
            new DialogDefinition("columns.filter", "按列筛选", "数据 / 筛选增强",760,560,540,380),
            new DialogDefinition("formula.iferror", "IFERROR 参数", "操作 / 一键公式",620,420,560,350),
            new DialogDefinition("formula.round", "ROUND 参数", "操作 / 一键公式",620,420,560,350),
            new DialogDefinition("formula.add", "加：自定义参数", "操作 / 一键公式",620,420,560,350),
            new DialogDefinition("formula.subtract", "减：自定义参数", "操作 / 一键公式",620,420,560,350),
            new DialogDefinition("formula.multiply", "乘：自定义参数", "操作 / 一键公式",620,420,560,350),
            new DialogDefinition("formula.divide", "除：自定义参数", "操作 / 一键公式",620,420,560,350),
            new DialogDefinition("formula.custom", "自定义函数", "操作 / 一键公式",620,420,560,350),
            new DialogDefinition("links.assistant", "链接助手", "操作 / 链接助手",1340,650,640,480),
            new DialogDefinition("file.organize", "文件批量整理", "操作 / 文件操作 / 批量处理",840,480,550,310),
            new DialogDefinition("file.names", "文件提取名称", "操作 / 文件操作 / 批量处理",1200,760,900,620),
            new DialogDefinition("file.ocrprefixes", "设置OCR前缀", "操作 / 文件操作 / 批量处理 / 文件提取名称",680,440,540,340),
            new DialogDefinition("file.ocrfields", "核对OCR字段", "操作 / 文件操作 / 批量处理 / 文件提取名称",820,500,640,360),
            new DialogDefinition("file.rename", "文件批量命名", "操作 / 文件操作 / 批量处理",840,480,550,310),
            new DialogDefinition("merge.headers", "跨工作簿表头匹配合并", "数据 / 表头匹配合并",1180,780,900,620),
            new DialogDefinition("elimination.receivable", "往来抵消", "数据 / 合并抵消 / 往来抵消",1180,700,900,580),
            new DialogDefinition("elimination.revenuecost", "收入成本抵消", "数据 / 合并抵消 / 收入成本抵消",1180,700,900,580),
#endif
            new DialogDefinition("file.pdf.current", "当前文档转 PDF", "操作 / 文件操作 / PDF 工具",960,700,800,600),
            new DialogDefinition("file.pdf.convert", "批量文件转 PDF", "操作 / 文件操作 / PDF 工具",960,700,800,600),
            new DialogDefinition("file.pdf.watermark", "PDF 添加水印", "操作 / 文件操作 / PDF 工具",960,700,800,600),
            new DialogDefinition("file.pdf.split", "PDF 拆分", "操作 / 文件操作 / PDF 工具",960,700,800,600),
            new DialogDefinition("file.pdf.merge", "PDF 合并", "操作 / 文件操作 / PDF 工具",960,700,800,600),
            new DialogDefinition("agent.developer", "开发预览", "Agent-Evo / 开发流程",920,680,720,480),
            new DialogDefinition("agent.developerpassword", "开发全文验证", "Agent-Evo / 开发流程 / 开发预览 / 全文",420,230,380,210),
            new DialogDefinition("agent.user", "用户预览", "Agent-Evo / 开发流程",920,680,720,480),
            new DialogDefinition("agent.sort", "自定排序", "Agent-Evo / 自定配置",760,600,600,460),
            new DialogDefinition("agent.name", "自定名称", "Agent-Evo / 自定配置",760,600,600,460),
            new DialogDefinition("agent.format", "自定格式", "Agent-Evo / 自定配置",1380,800,880,560),
            new DialogDefinition("agent.notes", "需求记录", "Agent-Evo / 实验计划",800,600,580,400),
            new DialogDefinition("agent.dialogsizes", "交互弹窗设置", "Agent-Evo / 自定配置",1520,640,720,480),
            new DialogDefinition("message.ok", "提示 / 结果", "各功能通用提示",560,360,440,280),
            new DialogDefinition("message.yesno", "是 / 否确认", "各功能通用确认",560,360,440,280),
            new DialogDefinition("message.okcancel", "确定 / 取消确认", "各功能通用确认",560,360,440,280)
        };
        internal static DialogDefinition Find(string id)
        { foreach(var item in All) if(item.Id==id)return item; return null; }
        internal static bool Valid(string id, Size size)
        { var item=Find(id); return size.Width>0 && size.Height>0 && (item==null || size.Width>=item.Minimum.Width && size.Height>=item.Minimum.Height); }
    }
    internal sealed class DialogSizeSnapshot
    {
        internal string Fingerprint="missing", Warning="";
        internal bool Writable=true;
        internal readonly Dictionary<string,Size> Values=new Dictionary<string,Size>(StringComparer.Ordinal);
    }
    internal static class DialogSizeStore
    {
        // Offline probes replace the root; production never writes project directories.
        internal static string RootOverride;
        internal static string PathForUser { get {return Path.Combine(RootOverride ?? Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),typeof(Connect).Assembly.GetName().Name,"DialogSizes.xml");} }
        private static string Digest(byte[] bytes)
        { using(var sha=SHA256.Create())return BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-",""); }
        private static string Fingerprint(string path)
        { try{return Digest(File.ReadAllBytes(path));}catch(FileNotFoundException){return "missing";}catch(DirectoryNotFoundException){return "missing";} }
        internal static DialogSizeSnapshot Read()
        {
            var result=new DialogSizeSnapshot();
            try
            {
                string path=PathForUser;byte[] bytes;
                try{bytes=File.ReadAllBytes(path);}catch(FileNotFoundException){return result;}catch(DirectoryNotFoundException){return result;}
                result.Fingerprint=Digest(bytes);
                if(bytes.Length>262144)throw new InvalidDataException("尺寸设置文件过大");
                var xml=new XmlDocument();
                using(var stream=new MemoryStream(bytes))
                using(var reader=XmlReader.Create(stream,new XmlReaderSettings { DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null,MaxCharactersInDocument=262144 }))xml.Load(reader);
                if(xml.DocumentElement==null || xml.DocumentElement.Name!="DialogSizes")throw new InvalidDataException("尺寸设置格式不正确");
                if(xml.DocumentElement.GetAttribute("version")!="1") {result.Writable=false;throw new InvalidDataException("尺寸设置版本不受支持，请保留原文件");}
                foreach(XmlElement node in xml.DocumentElement.SelectNodes("Window"))
                {
                    string id=node.GetAttribute("id");int w,h;
                    if(id.Length==0 || id.Length>100 || result.Values.ContainsKey(id) || !int.TryParse(node.GetAttribute("width"),out w) || !int.TryParse(node.GetAttribute("height"),out h) || !DialogCatalog.Valid(id,new Size(w,h)))
                    { result.Warning="部分无效尺寸已采用程序默认。";continue; }
                    result.Values.Add(id,new Size(w,h));
                }
            }
            catch(Exception ex)
            {
                result.Values.Clear(); result.Warning="采用程序默认："+ex.Message;
                if(result.Fingerprint=="missing")result.Writable=false;
            }
            return result;
        }
        internal static void Save(DialogSizeSnapshot original, Dictionary<string,Size> values)
        {
            if(!original.Writable)throw new InvalidOperationException(original.Warning);
            foreach(var pair in values)if(!DialogCatalog.Valid(pair.Key,pair.Value))throw new InvalidDataException("无效尺寸："+pair.Key);
            string path=PathForUser;Directory.CreateDirectory(Path.GetDirectoryName(path));
            using(var guard=new FileStream(path+".lock",FileMode.OpenOrCreate,FileAccess.ReadWrite,FileShare.None))
            {
                if(Fingerprint(path)!=original.Fingerprint)throw new InvalidOperationException("其他窗口已修改设置，请点击重新加载后再保存。");
                string temp=path+"."+Guid.NewGuid().ToString("N")+".tmp";
                try
                {
                    using(var writer=XmlWriter.Create(temp,new XmlWriterSettings {Encoding=new UTF8Encoding(false),Indent=true}))
                    {
                        writer.WriteStartElement("DialogSizes");writer.WriteAttributeString("version","1");
                        var keys=new List<string>(values.Keys);keys.Sort(StringComparer.Ordinal);
                        foreach(string key in keys) {writer.WriteStartElement("Window");writer.WriteAttributeString("id",key);writer.WriteAttributeString("width",values[key].Width.ToString(CultureInfo.InvariantCulture));writer.WriteAttributeString("height",values[key].Height.ToString(CultureInfo.InvariantCulture));writer.WriteEndElement();}
                        writer.WriteEndElement();
                    }
                    if(File.Exists(path))File.Replace(temp,path,null);else File.Move(temp,path);
                }
                finally {if(File.Exists(temp))File.Delete(temp);}
            }
        }
    }
    internal interface IDialogDefaultLayout
    {
        Size MeasureDefaultClient();
    }
    internal static class DialogSizing
    {
        private static readonly ThreadLocal<HashSet<Form>> customState=new ThreadLocal<HashSet<Form>>(()=>new HashSet<Form>());
        private static HashSet<Form> custom {get {return customState.Value;}}
        private static readonly ThreadLocal<Dictionary<Form,Size>> requestedState=new ThreadLocal<Dictionary<Form,Size>>(()=>new Dictionary<Form,Size>());
        private static Dictionary<Form,Size> requested {get {return requestedState.Value;}}
        [ThreadStatic] private static bool resolving;
        private static readonly object initialization=new object();
        private static readonly Dictionary<string,Size> measuredDefaults=new Dictionary<string,Size>();
        internal static void Initialize()
        {
            Ui.DialogSizeApplier=Apply;Ui.DialogFormatApplier=DialogFormatting.Attach;Ui.SaveControlWidth=DialogFormatting.SaveControlWidth;Ui.SaveRowHeight=DialogFormatting.SaveRowHeight;
        }
        private static Size DefaultFor(Form context,DialogDefinition definition)
        {
            if(resolving || !(definition.Id.StartsWith("formula.",StringComparison.Ordinal)||definition.Id.StartsWith("file.",StringComparison.Ordinal))) return definition.BaseDefault;
            string fontName;using(var font=Ui.MakeFont(10f,FontStyle.Regular,null))fontName=font.Name;
            string key=definition.Id+"|"+Scale(context).ToString("R",CultureInfo.InvariantCulture)+"|"+fontName;
            lock(initialization)
            {
                Size value;if(measuredDefaults.TryGetValue(key,out value)){definition.Default=value;return value;}
                resolving=true;
                try
                {
                    using(var form=DialogPreviewFactory.CreateDefaultMeasure(definition.Id,definition.BaseDefault))
                    {
                        form.PerformAutoScale();
                        float scale=Scale(form);
                        form.ClientSize=new Size((int)Math.Round(definition.BaseDefault.Width*scale),(int)Math.Round(definition.BaseDefault.Height*scale));
                        form.PerformLayout();
                        var layout=form as IDialogDefaultLayout;
                        if(layout==null)throw new InvalidOperationException("缺少默认布局测量："+definition.Id);
                        Size measured=layout.MeasureDefaultClient();
                        value=new Size(definition.BaseDefault.Width,Math.Max(definition.Minimum.Height,measured.Height));
                        measuredDefaults.Add(key,value);definition.Default=value;return value;
                    }
                }
                finally {resolving=false;}
            }
        }
        internal static void EnsureAllDefaults(Form context) {Initialize();foreach(var definition in DialogCatalog.All)DefaultFor(context,definition);}
        internal static bool HasOverride(Form form) {return custom.Contains(form);}
        internal static float Scale(Form form) {using(var g=form.CreateGraphics())return g.DpiX/96f;}
        internal static Size LogicalClient(Form form)
        {float scale=Scale(form);return new Size((int)Math.Round(form.ClientSize.Width/scale),(int)Math.Round(form.ClientSize.Height/scale));}
        internal static bool WasClamped(Form form)
        {Size desired;if(!requested.TryGetValue(form,out desired))return false;Size actual=LogicalClient(form);return actual.Width+1<desired.Width||actual.Height+1<desired.Height;}
        internal static void Prepare(Form form,string id,string icon,Color accent,FormWindowState initialState=FormWindowState.Normal)
        {
            Initialize();var d=DialogCatalog.Find(id);if(d==null)throw new ArgumentException("未登记的弹窗："+id);
            Ui.PrepareForm(form,DefaultFor(form,d),d.Minimum,icon,accent,id);
            if(initialState!=FormWindowState.Normal)form.Load+=delegate {
                if(form.TopLevel&&!Ui.IsInertPreview(form)&&!DialogPreviewFactory.IsPreview(form)&&!DialogPreviewFactory.IsConstructing)form.WindowState=initialState;
            };
        }
        private static bool Apply(Form form,string id)
        {
            var d=DialogCatalog.Find(id);if(d==null)return false;
            bool measure=DialogPreviewFactory.IsDefaultMeasurement;
            Size saved=d.Default;bool overridden=false;
            if(DialogPreviewFactory.IsConstructing&&DialogPreviewFactory.RequestedId==id)
            {saved=DialogPreviewFactory.RequestedSize;overridden=!measure;}
            else {
int[] user;var values=FormatSession.Capture(form).WindowSizes;if(values!=null&&values.TryGetValue(id,out user)&&user.Length==2){saved=new Size(user[0],user[1]);overridden=true;}
            }
            if(measure)saved=d.BaseDefault;
            requested[form]=saved;
            if(overridden)custom.Add(form);
            // Seed in logical units; apply the same physical conversion after autoscaling
            // for defaults and overrides. The host owns process DPI awareness.
            Size frame=new Size(form.Width-form.ClientSize.Width,form.Height-form.ClientSize.Height);
            Rectangle area=Screen.FromControl(form).WorkingArea;float scale=Scale(form);
            form.MinimumSize=new Size(d.Minimum.Width+frame.Width,d.Minimum.Height+frame.Height);
            form.ClientSize=measure?saved:new Size(Math.Min(saved.Width,Math.Max(1,(int)((area.Width-24-frame.Width)/scale))),Math.Min(saved.Height,Math.Max(1,(int)((area.Height-24-frame.Height)/scale))));
            if(!measure)
            {
                form.Load+=delegate {ApplyFinal(form,d,saved);};
            }
            form.Disposed+=delegate {custom.Remove(form);requested.Remove(form);};return true;
        }
        private static void ApplyFinal(Form form,DialogDefinition d,Size saved)
        {
            if(form.TopLevel&&form.WindowState!=FormWindowState.Normal)return;
            Rectangle visible=Screen.FromControl(form).WorkingArea;float scale=Scale(form);
            Size frame=new Size(form.Width-form.ClientSize.Width,form.Height-form.ClientSize.Height);
            if(!form.TopLevel) {
                form.MinimumSize=new Size((int)Math.Round(d.Minimum.Width*scale)+frame.Width,(int)Math.Round(d.Minimum.Height*scale)+frame.Height);
                form.ClientSize=new Size((int)Math.Round(saved.Width*scale),(int)Math.Round(saved.Height*scale));form.Location=Point.Empty;return;
            }
            int margin=(int)Math.Round(24*scale),width=Math.Max(1,visible.Width-margin-frame.Width),height=Math.Max(1,visible.Height-margin-frame.Height);
            form.MinimumSize=new Size(Math.Min((int)Math.Round(d.Minimum.Width*scale)+frame.Width,Math.Max(1,visible.Width-margin)),Math.Min((int)Math.Round(d.Minimum.Height*scale)+frame.Height,Math.Max(1,visible.Height-margin)));
            form.ClientSize=new Size((int)Math.Min(Math.Round((double)saved.Width*scale),width),(int)Math.Min(Math.Round((double)saved.Height*scale),height));
            form.Location=new Point(visible.Left+(visible.Width-form.Width)/2,visible.Top+(visible.Height-form.Height)/2);
        }
    }
}
