using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace Excel_Iterate
{
    // =====================================================================
    // 合并抵消（往来抵消 / 收入成本抵消）—— 功能区接线与业务编排
    //
    // 结构同 Connect.MergeHeaders.cs：窗体只做参数收集与结果呈现，一切 COM
    // 都在 Ribbon 宿主线程执行（窗体事件与回调同线程，不开后台 STA）。
    // 非模态显示取宿主主窗口作 owner——TrackerForm 锁死事故后全仓库统一
    // 写法，禁止 TopMost 无主窗。
    //
    // 两个功能区按钮共用一个窗体槽位：
    //   · 同 kind 已开 → 激活；
    //   · 异 kind 且旧窗空闲（不忙、无未确认批次）→ 换窗（旧窗资源先释放）；
    //   · 旧窗忙或本批未确认 → 提示先处理，不换窗。
    //
    // 写入目标 = 来源工作簿内的指定工作表或新建「抵消分录」表；写入前备份、
    // 快照，「取消操作」还原（新建表整表删除 / 已存在表快照回写）。
    // =====================================================================

    public partial class Connect
    {
        private static EliminationForm _eliminationForm;
        private static EliminationSession _eliminationSession;











        /// <summary>懒清理列映射临时记忆：清除已关闭工作簿的记忆条目（打开弹窗时执行）。</summary>
        private static void CleanupStaleMemory()
        {
            try
            {
                if (_app == null) return;
                dynamic app = _app;
                HashSet<string> open = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                using(var owned=new ComBag())
                {
                    dynamic books=owned.Own(app.Workbooks);
                    int count=Convert.ToInt32(books.Count);
                    for(int i=1;i<=count;i++)using(var item=new ComBag())
                    {
                        dynamic book=item.Own(MergeWorkbookGateway.GetItem(books,i));
                        string path=Convert.ToString(book.FullName);
                        if(!string.IsNullOrEmpty(path))open.Add(path);
                    }
                }
                EliminationForm.EliminationMapStore.CleanupExcept(open);
            }
            catch { }
        }

        /// <summary>
        /// 打开/切换抵消窗体。两个按钮共用一个窗体槽位（同 kind 激活、异 kind
        /// 空闲换窗、忙/未确认则提示）。
        /// </summary>
        private static void OpenElimination(EliminationKind kind)
        {
            if (_busy) return;
            EliminationForm current = _eliminationForm;
            if (current != null && !current.IsDisposed)
            {
                if (current.Kind == kind) { try { current.Activate(); } catch { } return; }
                if (current.IsBusy || current.CanUndo)
                {
                    UiDialog.Show("「" + current.Text + "」正在处理，或本批写入尚未「确认操作 / 取消操作」。\r\n请先处理完再打开另一种抵消。",
                        "合并抵消", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                // 旧窗空闲：先释放它持有的工作簿引用与写入态，再摘槽位关窗。
                // FormClosed 里的令牌判断此时不命中（槽位已空），只会 Dispose 旧窗，
                // 不会碰随后新建窗体的网关。
                ReleaseEliminationState();
                _eliminationForm = null;
                try { current.Close(); } catch { }
            }
            try
            {
                if (!EnsureApplication()) throw new InvalidOperationException("请先打开 WPS 表格。");
                CleanupStaleMemory();
                EliminationForm form = new EliminationForm(kind);
                _eliminationSession = new EliminationSession(_app);
                form.PreviewRequested += delegate { RunEliminationPreview(form); };
                form.ExecuteRequested += delegate { RunEliminationExecute(form); };
                form.UndoRequested += delegate { RunEliminationUndo(form); };
                form.ConfirmRequested += delegate { RunEliminationConfirm(form); };
                form.SheetsRequested += delegate(string path) { RunLoadEliminationSheets(form, path); };
                ApplyActiveSource(form);
                form.FormClosed += delegate
                {
                    bool mine = ReferenceEquals(_eliminationForm, form);
                    if (mine)
                    {
                        _eliminationForm = null;
                        ReleaseEliminationState();
                    }
                    try { form.PersistMapping(); } catch { }
                    form.Dispose();
                };
                _eliminationForm = form;
                // 与链接助手/表头匹配合并等非模态弹窗完全一致：取宿主主窗口作 owner 显示。
                IntPtr owner = GetHostWindowHandle();
                Log("合并抵消：打开前 enabled=" + HostDialogService.HostEnabled()
                    + " / 前台=" + NativeOle.GetForegroundWindow() + " / owner=" + owner);
                if (owner == IntPtr.Zero) form.Show(); else form.Show(new WindowWrapper(owner));
                Log("合并抵消：Show 后 enabled=" + HostDialogService.HostEnabled()
                    + " / 前台=" + NativeOle.GetForegroundWindow()
                    + " / 弹窗=" + (form.IsHandleCreated ? form.Handle : IntPtr.Zero));
            }
            catch (Exception ex)
            {
                Log("合并抵消：" + ex);
                UiDialog.Show(ex.Message, "合并抵消", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        /// <summary>卸载/关文档时关掉还开着的抵消窗，避免挂在已死的宿主上。</summary>
        internal static void CloseEliminationForm()
        {
            EliminationForm form = _eliminationForm;
            _eliminationForm = null;
            if (form != null)
            {
                try
                {
                    if (!form.IsDisposed && form.IsHandleCreated)
                        form.BeginInvoke((MethodInvoker)delegate { try { form.Close(); } catch { } });
                }
                catch { try { form.Close(); } catch { } }
            }
            // 宿主即将卸载：BeginInvoke 未必还会执行，同步释放 COM 引用（幂等）。
            ReleaseEliminationState();
        }

        /// <summary>释放抵消功能持有的网关与写入态（关窗/换窗/卸载共用；幂等）。</summary>
        private static void ReleaseEliminationState()
        {
            EliminationSession state=_eliminationSession;_eliminationSession=null;
            if(state==null)return;
            try {state.Dispose();}
            catch(Exception error){Log("合并抵消释放会话："+error.Message);}
            finally {if(state.OwnsBusy)_busy=false;}
        }

        private static EliminationGateway Gateway()
        {
            if(_eliminationSession==null)throw new InvalidOperationException("抵消会话已结束，请重新打开窗口。");
            return _eliminationSession.EnsureGateway();
        }

        // ---------------------------- 来源 ----------------------------

        /// <summary>
        /// 把当前活动工作簿与活动工作表填进来源区。COM 只能在 Ribbon 宿主线程做，
        /// 因此放在这里而不是弹窗构造函数里。不满足条件时留空并说明理由。
        /// </summary>
        private static void ApplyActiveSource(EliminationForm form)
        {
            try
            {
                using (ComBag bag = new ComBag())
                {
                    object book = null;
                    try { book = bag.Own(((dynamic)_app).ActiveWorkbook); } catch { }
                    if (book == null) return;
                    string path = Convert.ToString(((dynamic)book).FullName);
                    string sheet = "";
                    try
                    {
                        object active = bag.Own(((dynamic)_app).ActiveSheet);
                        sheet = Convert.ToString(((dynamic)active).Name);
                    }
                    catch { }
                    if (string.IsNullOrEmpty(path))
                    {
                        form.AppendLog("当前工作簿尚未保存到磁盘，请保存后重开本窗口，或点「选择来源…」指定一个 .xlsx 工作簿。");
                        return;
                    }
                    if (!string.Equals(Path.GetExtension(path), ".xlsx", StringComparison.OrdinalIgnoreCase))
                    {
                        form.AppendLog("当前工作簿不是 .xlsx（合并抵消只处理 .xlsx）：" + path + "。请点「选择来源…」另选。");
                        return;
                    }
                    form.SetSource(path, sheet);
                }
            }
            catch (Exception ex) { Log("合并抵消读取活动工作簿：" + ex); }
        }

        /// <summary>只为填下拉而读一次工作表名：已打开则复用，否则只读打开。</summary>
        private static void RunLoadEliminationSheets(EliminationForm form, string path)
        {
            if (_busy) return;
            _busy = true;
            try
            {
                if (!EnsureApplication()) throw new InvalidOperationException("请先打开 WPS 表格。");
                if (string.IsNullOrEmpty(path)) throw new InvalidOperationException("路径为空。");
                if (!File.Exists(path)) throw new InvalidOperationException("文件不存在：" + path);
                List<string> names = Gateway().SheetNames(path);
                form.SetSheets(names, null);
            }
            catch (Exception ex)
            {
                Log("合并抵消读取工作表名：" + ex);
                form.SetSheets(null, ex.Message);
            }
            finally { _busy = false; }
        }

        // ---------------------------- 预检 ----------------------------

        private static void RunEliminationPreview(EliminationForm form)
        { StartEliminationPreflight(form,false); }

        private static void StartEliminationPreflight(EliminationForm form,bool execute)
        {
            if(_busy)return;
            var session=_eliminationSession;if(session==null||session.IsDisposed)return;
            _busy=true;session.OwnsBusy=true;session.Phase=EliminationPhase.Reading;form.SetBusy(true);form.ClearPreview();
            bool started=false;
            try
            {
                EliminationRequest request=form.BuildRequest();
                string invalid=ValidateEliminationRequest(request);if(invalid!=null){form.SetLog(invalid);return;}
                var gateway=session.EnsureGateway();form.SetStatus("正在读取来源表头…");
                var headers=gateway.ReadHeaders(request.Path,request.SheetName,request.HeaderRow);
                string headerError=headers.Validate("来源工作簿");if(headerError!=null){form.SetHeaders(null,0,headerError);return;}
                string key=request.Kind+"|"+request.Path+"|"+request.SheetName+"|"+request.HeaderRow.ToString(CultureInfo.InvariantCulture);
                if(form.MapRowCount==0||key!=session.HeadersKey)
                {form.SetHeaders(BuildColumns(headers,request.Kind),Math.Max(0,headers.LastUsedRow-request.HeaderRow),null);session.HeadersKey=key;request=form.BuildRequest();}
                invalid=EliminationEngine.ValidateMapping(request.Kind,request.Columns);
                if(invalid==null&&request.Kind==EliminationKind.RevenueCost)invalid=EliminationEngine.ValidateMatrixMapping(request.Kind,request.Columns);
                if(invalid!=null){form.SetLog(invalid+"\r\n请调整列映射后重新预检。");return;}
                request=request.Copy();
                var options=BuildMatrixOptions(request,headers);
                if(execute)
                {
                    int row,column;if(!EliminationGateway.TryParseCell(request.StartCell,out row,out column)){form.SetLog("起始单元格无法解析。");return;}
                    options.StartRow=row;options.StartColumn=column;
                }
                var reader=gateway.BeginReadRows(request,headers);reader.KeepMonitor=true;session.SourceWatch=reader.Monitor;
                var computation=new EliminationComputation(request,headers,options);
                var task=new HostBatchRunner<EliminationRowsBlock,EliminationPrepared>(reader,computation.Accept,computation.Finish,
                    delegate(EliminationPrepared prepared,Exception failure,bool cancelled)
                    {
                        if(!ReferenceEquals(_eliminationSession,session)||session.IsDisposed||form.IsDisposed)return;
                        session.PreviewTask=null;form.CancelPreview=null;
                        try
                        {
                            if(failure!=null)throw failure;
                            if(cancelled){form.SetLog("已取消，本次未写入任何内容。");session.Phase=EliminationPhase.Idle;return;}
                            session.SourceWatch.Validate();
                            if(execute){session.Phase=EliminationPhase.Writing;CommitElimination(form,prepared);}
                            else
                            {
                                session.Diff=prepared.Diff;form.SetPreviewMatrix(prepared.Matrix);
                                if(prepared.Diff!=null)form.AppendLog("往来抵消差异表："+prepared.Diff.Rows.Count+" 行 / 主体对 "+prepared.Diff.GroupCount+" 组。");
                                session.Phase=prepared.Matrix.Errors.Count==0&&prepared.Matrix.Rows.Count>0?EliminationPhase.PreviewReady:EliminationPhase.Idle;
                                if(session.Phase==EliminationPhase.PreviewReady)EliminationForm.EliminationMapStore.Save(EliminationForm.EliminationMapStore.Key(prepared.Request.Kind,prepared.Request.Path),prepared.Request.Columns);
                            }
                        }
                        catch(Exception error){session.Phase=EliminationPhase.Idle;form.ClearPreview();form.SetLog("预检失败："+error.Message);Log("抵消预检："+error);}
                        finally
                        {session.StopWatch();session.OwnsBusy=false;form.SetBusy(false);_busy=false;}
                    },delegate(EliminationRowsBlock block){if(!form.IsDisposed)form.SetStatus("正在读取来源："+block.Completed+" / "+block.Total+" 行…");});
                task.Captured=delegate{session.Phase=EliminationPhase.Calculating;if(!form.IsDisposed)form.SetStatus("正在计算抵消结果…");};
                session.PreviewTask=task;form.CancelPreview=session.CancelPreview;started=true;
                if(!reader.CanYield)form.AppendLog("宿主不支持来源变化监听，采用同步读取、后台计算的兼容路径。");
                task.Start();
            }
            catch(Exception error)
            {
                if(session.PreviewTask!=null){session.PreviewTask.Dispose();session.PreviewTask=null;}
                started=false;form.SetLog("预检失败："+error.Message);session.Phase=EliminationPhase.Idle;
            }
            finally
            {if(!started){session.StopWatch();session.OwnsBusy=false;session.Phase=EliminationPhase.Idle;form.CancelPreview=null;form.SetBusy(false);_busy=false;}}
        }

        /// <summary>表头快照 → 列映射（角色按表头文本猜，用户可在网格覆盖）。</summary>
        private static List<EliminationColumn> BuildColumns(MergeHeaderSet headers, EliminationKind kind)
        {
            List<EliminationColumn> columns = new List<EliminationColumn>();
            foreach (MergeColumn source in headers.Columns)
            {
                columns.Add(new EliminationColumn
                {
                    Number = source.Number,
                    Letter = source.Letter,
                    Header = source.Text,
                    Role = EliminationEngine.GuessRole(source.Text, kind)
                });
            }
            return columns;
        }

        private static bool HasRole(List<EliminationColumn> columns, EliminationRole role)
        {
            foreach (EliminationColumn column in columns)
                if (column.Role == role) return true;
            return false;
        }

        /// <summary>界面层之外的请求校验（路径/工作表/表头行）；null 表示通过。</summary>
        private static string ValidateEliminationRequest(EliminationRequest request)
        {
            if (request.Error != null) return request.Error;
            if (request.Path.Length == 0) return "请先选择来源工作簿。";
            if (request.SheetName.Length == 0) return "请选择来源工作表。";
            if (!File.Exists(request.Path)) return "来源文件不存在：" + request.Path;
            if (!string.Equals(Path.GetExtension(request.Path), ".xlsx", StringComparison.OrdinalIgnoreCase))
                return "来源工作簿需为 .xlsx：" + request.Path;
            if(request.Kind==EliminationKind.Receivable&&request.DiffSheetItem!=EliminationGateway.NewSheetItem&&request.DiffSheetItem==request.TargetSheetItem)
                return "差异表和结果表必须使用不同的工作表，避免互相覆盖。";
            return null;
        }

        // ---------------------------- 执行 ----------------------------

        private static void RunEliminationExecute(EliminationForm form)
        { StartEliminationPreflight(form,true); }

        private static void CommitElimination(EliminationForm form,EliminationPrepared prepared)
        {
            try
            {
                EliminationRequest request=prepared.Request;
                MergeHeaderSet headers=prepared.Headers;
                EliminationGateway gateway=Gateway();
                int startRow=prepared.Options.StartRow,startColumn=prepared.Options.StartColumn;
                bool receivable=request.Kind==EliminationKind.Receivable;
                EliminationMatrix built=prepared.Matrix;
                MatrixOptions wlOptions=prepared.Options;
                List<EliminationRow> wlRows=prepared.FormulaRows;
                _eliminationSession.Diff=prepared.Diff;
                if(built==null||built.Rows.Count==0){form.SetLog("没有可写入的内容。");return;}
                if(built.Errors.Count>0){form.SetLog(string.Join("\r\n",built.Errors.ToArray()));return;}
                MergeWritePlan plan=EliminationGateway.BuildMatrixPlan(built,wlOptions);
                string guard = MergePlanner.Guard(plan);
                if (guard != null) { form.SetLog(guard); return; }
                // Validate both outputs before creating sheets or clearing any region.
                MergeWritePlan diffPlan=receivable?EliminationGateway.BuildDiffPlan(prepared.Diff,wlOptions):null;
                if(receivable)
                {
                    string diffGuard=MergePlanner.Guard(diffPlan);
                    if(diffGuard!=null){form.SetLog(diffGuard);return;}
                }

                string defaultName = EliminationGateway.MatrixSheetNameFor(request.Kind);
                string targetDesc = request.TargetSheetItem == EliminationGateway.NewSheetItem
                    ? "新建工作表（默认名「" + defaultName + "」，重名自动编号）"
                    : request.TargetSheetItem;
                bool sameSheet = string.Equals(request.TargetSheetItem, request.SheetName, StringComparison.Ordinal);

                if (UiDialog.Show(form, !receivable
                        ? ConfirmationText(request, built, plan, targetDesc, sameSheet)
                        : ConfirmationText(request, _eliminationSession.Diff, built, plan, targetDesc, sameSheet),
                    "确认执行" + form.Text,
                    MessageBoxButtons.OKCancel, MessageBoxIcon.Warning,
                    MessageBoxDefaultButton.Button2) != DialogResult.OK)
                {
                    form.AppendLog("已取消，未写入任何内容。");
                    return;
                }

                if(_eliminationSession.SourceWatch!=null)_eliminationSession.SourceWatch.Validate();
                _eliminationSession.StopWatch();
                // 确认后才建目标：用户取消不留下空表。
                form.SetStatus("正在定位目标工作表…");
                // 往来抵消：先建「差异表」拿真实表名（重名自动编号），再据此重算两张表
                if (receivable)
                {
                    EliminationTargetHandle diffTarget = gateway.CreateTarget(
                        request.Path, request.DiffSheetItem, EliminationGateway.DiffDefaultSheetName);
                    _eliminationSession.DiffTarget = diffTarget;
                    if (request.DiffSheetItem == EliminationGateway.NewSheetItem
                        && !string.Equals(diffTarget.SheetName, EliminationGateway.DiffDefaultSheetName, StringComparison.Ordinal))
                        form.AppendLog("差异表重名，已自动编号为「" + diffTarget.SheetName + "」。");
                    wlOptions.DiffSheetName = diffTarget.SheetName;
                    if(!request.StaticValues)
                    {
                        EliminationDiffTable diff2=null;EliminationMatrix rebuilt=null;
                        EliminationDiffBuilder.Build(wlRows,request.Columns,wlOptions,out diff2,out rebuilt);
                        _eliminationSession.Diff=diff2;built=rebuilt;plan=EliminationGateway.BuildMatrixPlan(built,wlOptions);
                    }
                }
                EliminationTargetHandle target = gateway.CreateTarget(request.Path, request.TargetSheetItem, defaultName);
                _eliminationSession.Target = target;
                if (sameSheet && !target.IsNewSheet)
                    form.AppendLog("注意：目标工作表与来源相同，请确认写入区域未覆盖源数据。");

                form.SetStatus("正在保存并备份目标工作簿…");
                string saved = gateway.SaveTargetBookPath(target.Book);
                _eliminationSession.BackupPath = gateway.CreateBackup(saved);

                // 快照先取，再清区域：这样「取消操作」能还原到写入前的原样。
                form.SetStatus("正在快照写入区域…");
                _eliminationSession.Snapshot = gateway.Capture(target, plan);
                // 目标若之前写过矩阵，残留的合并单元格会让批量赋值失败，先拆掉。
                gateway.PrepareTargetRegion(target, plan);

                form.SetStatus("正在写入…");
                // 往来抵消：**先写差异表、后写结果表**——结果表的金额是跨表引用差异表，
                // 引用目标先有数据，写入核对也更稳（实机曾报「公式写入核对失败：G4」）。
                int diffCells = 0;
                if (receivable)
                {
                    EliminationTargetHandle diffTarget = _eliminationSession.DiffTarget;
                    _eliminationSession.DiffSnapshot = gateway.Capture(diffTarget, diffPlan);
                    gateway.PrepareTargetRegion(diffTarget, diffPlan);
                    gateway.Write(diffTarget, diffPlan);
                    gateway.StyleDiffHeader(diffTarget, startRow, startColumn, _eliminationSession.Diff.ColumnCount);
                    // 差异表：第 1 行表头、第 2 行起数据；金额列 G-N（索引 6-13）；序号 A/年度 B 居中
                    int diffDataTop = startRow + 1;
                    int diffDataRows = _eliminationSession.Diff.Rows.Count > 0
                        ? (_eliminationSession.Diff.Rows[_eliminationSession.Diff.Rows.Count - 1].RowNumber - diffDataTop + 1) : 0;
                    int diffGap = wlOptions.GroupGap > 0 ? wlOptions.GroupGap : 0;
                    int diffGaps = diffGap > 0 && _eliminationSession.Diff.GroupCount > 1
                        ? (_eliminationSession.Diff.GroupCount - 1) * diffGap : 0;
                    gateway.FinishMatrixOutput(diffTarget, startRow, 1, diffDataTop, startColumn,
                        _eliminationSession.Diff.ColumnCount, diffDataRows + diffGaps,
                        new int[] { 6, 7, 8, 9, 10, 11, 12, 13 }, new int[] { 0, 1 });
                    diffCells = diffPlan.TotalCells;
                    form.AppendLog("已写入 " + diffCells + " 格到「" + diffTarget.SheetName + "」（往来抵消差异）");
                }

                gateway.Write(target, plan);
                // 矩阵样式链两种业务恒走（旧代码按 kind 分流，往来走了分录式样式致结果表无格式）
                {
                    int firstHeaderRow = startRow + (request.TotalRow ? 1 : 0);
                    gateway.MergeMatrixHeader(target, built, startRow, startColumn, request.TotalRow);
                    gateway.StyleMatrixHeader(target, built, firstHeaderRow, startColumn, request.TotalRow);
                    // 接入 (2) 规范：金额列会计格式（千位分隔）、数据区对齐、列宽 AutoFit
                    int dataTop = firstHeaderRow + 2;
                    int dataRows = built.Rows.Count > 0
                        ? (built.Rows[built.Rows.Count - 1].RowNumber - dataTop + 1) : 0;
                    int[] moneyCols = EliminationGateway.MoneyColumnIndices(built);
                    int[] centerCols = new int[] { 3, 4 };   // 调整事项、会计科目
                    // 「格式参考」样张变体：宋体 11、slantDashDot 框线、SUBTOTAL 行也走会计格式
                    gateway.FinishMatrixOutput(target, firstHeaderRow, 2, dataTop, startColumn,
                        built.Columns.Count, dataRows, moneyCols, centerCols, true);
                }
                gateway.SaveTargetBook(target.Book);

                _eliminationSession.Phase=EliminationPhase.UndoAvailable;
                form.MarkWritten();
                form.AppendLog("已写入 " + plan.TotalCells + " 格到「" + target.SheetName + "」"
                    + (target.IsNewSheet ? "（本次新建）" : "") + "；备份 " + _eliminationSession.BackupPath);
                form.SetStatus("写入完成；如需还原请点「取消操作」，确认无误请点「确认操作」。");
                Log("合并抵消写入：" + plan.TotalCells + " 格 → " + request.Path + " [" + target.SheetName + "]");
            }
            catch (Exception ex)
            {
                Log("合并抵消执行：" + ex);
                bool recovered=TryEliminationRollback(form);
                string recovery=recovered?"目标内容未改动或已还原，保存状态见日志。":"回滚尚未完成，请重试取消操作或使用备份恢复。";
                form.AppendLog("执行失败："+ex.Message+"；"+recovery);
                // 日志区不显眼；用模态框把真实失败原因直接端到用户面前，避免"以为执行了"。
                UiDialog.Show(form, "执行失败：" + ex.Message + "\r\n"+recovery,
                    form.Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
                if(_eliminationSession!=null&&_eliminationSession.Phase==EliminationPhase.Writing)_eliminationSession.Phase=EliminationPhase.Idle;
                form.SetBusy(false);
                _busy = false;
            }
        }

        /// <summary>界面选项 → 矩阵底稿的输出选项。</summary>
        private static MatrixOptions BuildMatrixOptions(EliminationRequest request, MergeHeaderSet headers)
        {
            MatrixOptions options = new MatrixOptions();
            options.SourceSheetName = request.SheetName;
            options.SourceBookName = "";
            options.ConditionValue = request.ConditionValue;
            options.HasConditionColumn = HasRole(request.Columns, EliminationRole.Condition);
            options.TotalRow = request.TotalRow;
            options.StaticValues = request.StaticValues;
            options.FullColumnRefs = request.FullColumnRefs;
            // 「引用整列」统一表达「自动扩展」：勾选 → SUMIFS 整列、现算 SUMPRODUCT 算到工作表末尾，
            // 源表新增行自动纳入；取消 → 两者都限定到当前数据区，快，但新增行后需重新执行。
            options.LimitComputedRange = !request.FullColumnRefs;
            options.GroupGap = request.GapRows > 0 ? request.GapRows : 0;
            options.SourceFirstRow = request.HeaderRow + 1;
            options.SourceLastRow = headers == null ? 0 : headers.LastUsedRow;
            options.StartRow = 1;
            options.StartColumn = 1;
            return options;
        }

        /// <summary>往来抵消的确认文案（差异表 + 结果表两张表，各自按下拉选择）。</summary>
        private static string ConfirmationText(EliminationRequest request, EliminationDiffTable diff,
            EliminationMatrix matrix, MergeWritePlan plan, string targetDesc, bool sameSheet)
        {
            StringBuilder text = new StringBuilder();
            text.AppendLine("来源：" + request.Path + "  [" + request.SheetName + "]  表头行 "
                + request.HeaderRow.ToString(CultureInfo.InvariantCulture));
            string diffDesc = request.DiffSheetItem == EliminationGateway.NewSheetItem
                ? "新建工作表（默认名「" + EliminationGateway.DiffDefaultSheetName + "」，重名自动编号）"
                : request.DiffSheetItem;
            text.AppendLine("目标1（差异表）：[" + diffDesc + "]");
            text.AppendLine("目标2（结果表）：[" + targetDesc + "]  起始 " + request.StartCell);
            text.AppendLine();
            if (diff != null)
            {
                text.AppendLine("主体对 " + diff.GroupCount + " 组 × 年度 " + diff.Years.Count + " 个 × 科目对 "
                    + diff.RowsPerYear + " → 差异表 " + diff.Rows.Count + " 行（每主体对后空 "
                    + (request.GapRows > 0 ? request.GapRows : 0) + " 行）");
                text.AppendLine("年度：" + string.Join("、", diff.Years.ToArray()));
                if (diff.FilteredRows > 0) text.AppendLine("条件列筛掉 " + diff.FilteredRows + " 行。");
            }
            if (matrix != null)
                text.AppendLine("结果表 " + matrix.Rows.Count + " 行（仅输出抵消金额不为 0 的科目对）。");
            if (request.StaticValues)
                text.AppendLine("金额以静态数值写入（不写公式）。");
            else
                text.AppendLine("金额写为活公式（SUMPRODUCT 前缀匹配），源数据变动后结果自动重算。");
            if (sameSheet)
                text.AppendLine("⚠ 目标工作表与来源相同，请确认写入区域未覆盖源数据。");
            text.AppendLine();
            text.AppendLine("写入前会自动保存并备份目标工作簿（同目录「ExcelIterate合并备份」），写入后可用「取消操作」还原。是否继续？");
            return text.ToString();
        }

        /// <summary>矩阵式底稿的确认文案。</summary>
        private static string ConfirmationText(EliminationRequest request, EliminationMatrix matrix,
            MergeWritePlan plan, string targetDesc, bool sameSheet)
        {
            StringBuilder text = new StringBuilder();
            text.AppendLine("来源：" + request.Path + "  [" + request.SheetName + "]  表头行 "
                + request.HeaderRow.ToString(CultureInfo.InvariantCulture));
            text.AppendLine("目标：" + request.Path + "  [" + targetDesc + "]  起始 " + request.StartCell);
            text.AppendLine();
            text.AppendLine("交易对 " + matrix.GroupCount + " 组 × 年度 " + matrix.Years.Count + " 个 → "
                + matrix.Rows.Count + " 行、" + matrix.ColumnCount + " 列，共 " + plan.TotalCells + " 个单元格"
                + (matrix.Balanced ? "（借贷平衡）" : "（⚠ 借贷不平，请检查）"));
            text.AppendLine("年度：" + string.Join("、", matrix.Years.ToArray()));
            if (matrix.SkippedRows > 0)
                text.AppendLine("跳过 " + matrix.SkippedRows + " 行源数据（原因见预检报告）。");
            if (matrix.FilteredRows > 0)
                text.AppendLine("条件列筛掉 " + matrix.FilteredRows + " 行。");
            if (request.StaticValues)
                text.AppendLine("金额以静态数值写入（不写公式）。");
            else
                text.AppendLine("金额写为活公式（SUMIFS / SUMPRODUCT），源数据变动后结果自动重算。");
            if (sameSheet)
                text.AppendLine("⚠ 目标工作表与来源相同，请确认写入区域未覆盖源数据。");
            text.AppendLine();
            text.AppendLine("写入前会自动保存并备份目标工作簿（同目录「ExcelIterate合并备份」），写入后可用「取消操作」还原。是否继续？");
            return text.ToString();
        }

        private static string ConfirmationText(EliminationRequest request, EliminationResult result,
            MergeWritePlan plan, string targetDesc, bool sameSheet)
        {
            StringBuilder text = new StringBuilder();
            text.AppendLine("来源：" + request.Path + "  [" + request.SheetName + "]  表头行 "
                + request.HeaderRow.ToString(CultureInfo.InvariantCulture));
            text.AppendLine("目标：" + request.Path + "  [" + targetDesc + "]  起始 " + request.StartCell);
            text.AppendLine();
            text.AppendLine("分录 " + result.Entries.Count + " 条，共 " + plan.TotalCells + " 个单元格；总借 "
                + result.TotalDebit.ToString("0.00", CultureInfo.InvariantCulture)
                + "、总贷 " + result.TotalCredit.ToString("0.00", CultureInfo.InvariantCulture)
                + (result.Balanced ? "（借贷平衡）" : "（⚠ 借贷不平，请检查）"));
            if (result.SkippedRows > 0)
                text.AppendLine("跳过 " + result.SkippedRows + " 行源数据（原因见预检报告）。");
            if (sameSheet)
                text.AppendLine("⚠ 目标工作表与来源相同，请确认写入区域未覆盖源数据。");
            text.AppendLine();
            text.AppendLine("写入前会自动保存并备份目标工作簿（同目录「ExcelIterate合并备份」），写入后可用「取消操作」还原。是否继续？");
            return text.ToString();
        }

        private static bool TryEliminationRollback(EliminationForm form)
        {
            var session=_eliminationSession;if(session==null)return true;
            var target=session.Target??session.DiffTarget;
            if(target==null)return true;
            try
            {
                var gateway=session.EnsureGateway();session.RestoreTargets(gateway.Restore);form.MarkUndone();
                try {gateway.SaveTargetBook(target.Book);}
                catch(Exception save){form.AppendLog("内容已还原，但自动保存失败，请手动保存："+save.Message);}
                form.AppendLog("已删除本次新建表或按快照还原目标内容。");return true;
            }
            catch(Exception error)
            {
                session.Phase=EliminationPhase.UndoAvailable;form.MarkWritten();
                Log("合并抵消还原失败："+error);
                form.AppendLog("还原尚未完成，可点「取消操作」重试；备份："+session.BackupPath+"；"+error.Message);return false;
            }
        }

        // ---------------------------- 取消 / 确认 ----------------------------

        private static void RunEliminationUndo(EliminationForm form)
        {
            if (_busy) return;
            _busy = true;
            form.SetBusy(true);
            try
            {
                EliminationTargetHandle target = _eliminationSession.Target??_eliminationSession.DiffTarget;
                if (target == null || target.Sheet == null)
                {
                    form.AppendLog("没有可还原的批次。");
                    return;
                }
                if (UiDialog.Show(form, target.IsNewSheet
                    ? "将删除本次新建的工作表「" + target.SheetName + "」并自动保存。是否继续？"
                    : "将用写入前的快照还原目标工作表，并自动保存。是否继续？",
                    "取消操作", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning,
                    MessageBoxDefaultButton.Button2) != DialogResult.OK)
                {
                    form.AppendLog("已取消还原。");
                    return;
                }
                EliminationGateway gateway = Gateway();
                _eliminationSession.RestoreTargets(gateway.Restore);
                _eliminationSession.Phase=EliminationPhase.Idle;
                form.MarkUndone();
                try {gateway.SaveTargetBook(target.Book);}
                catch(Exception save){form.AppendLog("内容已还原，但自动保存失败，请手动保存："+save.Message);}
                form.AppendLog(target.IsNewSheet
                    ? "已删除新建工作表「" + target.SheetName + "」。文件备份保留：" + _eliminationSession.BackupPath
                    : "已还原到写入前状态。文件备份保留：" + _eliminationSession.BackupPath);
                Log("合并抵消：已还原");
                _eliminationSession.Target = null;
                _eliminationSession.Snapshot = null;
            }
            catch (Exception ex)
            {
                Log("合并抵消取消操作：" + ex);
                form.AppendLog("还原失败：" + ex.Message + "；请用备份文件手动恢复：" + _eliminationSession.BackupPath);
            }
            finally
            {
                form.SetBusy(false);
                _busy = false;
            }
        }

        private static void RunEliminationConfirm(EliminationForm form)
        {
            _eliminationSession.Snapshot = null;
            _eliminationSession.DiffSnapshot=null;
            _eliminationSession.Target=null;
            _eliminationSession.DiffTarget=null;
            _eliminationSession.Phase=EliminationPhase.Idle;
            Log("合并抵消：已确认，本批结束。备份保留于 " + _eliminationSession.BackupPath);
        }
    }
}
