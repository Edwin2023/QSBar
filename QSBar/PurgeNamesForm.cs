using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace QSBar
{
    public class PurgeNamesForm : Form
    {
        private readonly NameCommands.PurgeResult _result;
        private readonly CheckBox _chkIncludeUsedExternal;

        public List<NameCommands.NameEntry> NamesToDelete { get; private set; }

        public PurgeNamesForm(NameCommands.PurgeResult result)
        {
            _result = result;
            NamesToDelete = new List<NameCommands.NameEntry>();

            bool cn = WpsExcelAddIn.UseChineseRibbon;

            Text = cn ? "净化名称" : "Purge Names";
            FormBorderStyle = FormBorderStyle.Sizable;
            StartPosition = FormStartPosition.CenterScreen;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(620, 520);
            MinimumSize = new Size(480, 360);
            Font = new Font("Microsoft YaHei UI", 9F);

            int deletable = result.Broken.Count + result.Unused.Count + result.UnusedExternal.Count;

            Label summary = new Label();
            summary.Dock = DockStyle.Top;
            summary.Height = 86;
            summary.Padding = new Padding(12, 10, 12, 6);
            summary.Text = string.Format(
                cn ? "将删除 {0} 个名称：错误名称 {1} 个，未被引用 {2} 个，未被引用的外链名称 {3} 个。\n"
                   + "保留在用名称 {4} 个，内置名称 {5} 个（打印区域、打印标题等）。"
                   : "Will delete {0} names: {1} broken, {2} unreferenced, {3} unreferenced external.\n"
                   + "Keeping {4} names in use and {5} built-in names (print area, print titles, etc.).",
                deletable, result.Broken.Count, result.Unused.Count, result.UnusedExternal.Count,
                result.Kept.Count, result.BuiltInSkipped);

            ListBox list = new ListBox();
            list.Dock = DockStyle.Fill;
            list.IntegralHeight = false;
            list.HorizontalScrollbar = true;
            list.SelectionMode = SelectionMode.None;
            FillList(list, cn);

            Panel bottom = new Panel();
            bottom.Dock = DockStyle.Bottom;
            bottom.Height = result.UsedExternal.Count > 0 ? 86 : 52;
            bottom.Padding = new Padding(12, 6, 12, 10);

            if (result.UsedExternal.Count > 0)
            {
                _chkIncludeUsedExternal = new CheckBox();
                _chkIncludeUsedExternal.Dock = DockStyle.Top;
                _chkIncludeUsedExternal.Height = 34;
                _chkIncludeUsedExternal.Checked = true;
                _chkIncludeUsedExternal.Text = string.Format(
                    cn ? "删除仍被引用的 {0} 个外链名称（取消勾选可保留；引用它们的公式会变成 #NAME?）"
                       : "Delete the {0} external names still in use (uncheck to keep; formulas referencing them will become #NAME?)",
                    result.UsedExternal.Count);
                bottom.Controls.Add(_chkIncludeUsedExternal);
            }

            Button ok = new Button();
            ok.Text = cn ? "删除" : "Delete";
            ok.Size = new Size(90, 28);
            ok.Anchor = AnchorStyles.Right | AnchorStyles.Bottom;
            ok.DialogResult = DialogResult.OK;
            ok.Click += OnOk;

            Button cancel = new Button();
            cancel.Text = cn ? "取消" : "Cancel";
            cancel.Size = new Size(90, 28);
            cancel.Anchor = AnchorStyles.Right | AnchorStyles.Bottom;
            cancel.DialogResult = DialogResult.Cancel;

            Panel buttons = new Panel();
            buttons.Dock = DockStyle.Bottom;
            buttons.Height = 42;
            ok.Location = new Point(bottom.ClientSize.Width - 200, 6);
            cancel.Location = new Point(bottom.ClientSize.Width - 102, 6);
            buttons.Controls.Add(ok);
            buttons.Controls.Add(cancel);
            buttons.Resize += delegate
            {
                ok.Location = new Point(buttons.ClientSize.Width - 212, 6);
                cancel.Location = new Point(buttons.ClientSize.Width - 114, 6);
            };
            bottom.Controls.Add(buttons);

            Controls.Add(list);
            Controls.Add(bottom);
            Controls.Add(summary);

            AcceptButton = ok;
            CancelButton = cancel;
        }

        private void FillList(ListBox list, bool cn)
        {
            // 上万条逐项 Add 会触发上万次重排，必须挂起绘制
            list.BeginUpdate();
            try { FillGroups(list, cn); }
            finally { list.EndUpdate(); }
        }

        private void FillGroups(ListBox list, bool cn)
        {
            AddGroup(list, cn ? "错误名称（将删除）" : "Broken (will delete)", _result.Broken);
            AddGroup(list, cn ? "未被引用（将删除）" : "Unreferenced (will delete)", _result.Unused);
            AddGroup(list, cn ? "未被引用的外链名称（将删除）" : "Unreferenced external (will delete)",
                _result.UnusedExternal);
            AddGroup(list, cn ? "仍被引用的外链名称（默认保留）" : "External but still in use (kept by default)",
                _result.UsedExternal);
            AddGroup(list, cn ? "表内引用且在用（保留）" : "In use within workbook (kept)", _result.Kept);
        }

        private static void AddGroup(ListBox list, string header, List<NameCommands.NameEntry> entries)
        {
            if (entries.Count == 0) return;

            list.Items.Add("── " + header + " (" + entries.Count + ") ──");
            foreach (NameCommands.NameEntry e in entries)
            {
                string refers = e.RefersTo ?? "";
                if (refers.Length > 90) refers = refers.Substring(0, 90) + "...";
                string flag = e.IsHidden
                    ? (WpsExcelAddIn.UseChineseRibbon ? " [隐藏]" : " [hidden]")
                    : "";
                list.Items.Add("    " + e.FullName + flag + "   →   " + refers);
            }
            list.Items.Add("");
        }

        private void OnOk(object sender, EventArgs e)
        {
            NamesToDelete.Clear();
            NamesToDelete.AddRange(_result.Broken);
            NamesToDelete.AddRange(_result.Unused);
            NamesToDelete.AddRange(_result.UnusedExternal);

            if (_chkIncludeUsedExternal != null && _chkIncludeUsedExternal.Checked)
            {
                NamesToDelete.AddRange(_result.UsedExternal);
            }
        }
    }
}
