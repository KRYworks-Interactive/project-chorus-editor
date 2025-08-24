﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

using Ched.Core.Notes;
using Ched.Core;
using Ched.Core.Events;
using Ched.Configuration;
using Ched.Localization;
using Ched.Plugins;
using Ched.Properties;
using Ched.UI.Operations;
using Ched.UI.Windows;

namespace Ched.UI
{
    public partial class MainForm : Form
    {
        private const string ApplicationTitle = "KRYworks Project Chorus Editor";
        private readonly string FileExtension = ".pce";
        private string FileTypeFilter => FileFilterStrings.ChedFilter + string.Format("({0})|{1}", "*" + FileExtension, "*" + FileExtension);

        private bool isPreviewMode;

        private ScoreBook ScoreBook { get; set; }
        private OperationManager OperationManager { get; }

        private DarkScrollBar NoteViewScrollBar { get; }
        private NoteView NoteView { get; }

        private ToolStripButton ZoomInButton;
        private ToolStripButton ZoomOutButton;
        private ToolStripMenuItem WidenLaneWidthMenuItem;
        private ToolStripMenuItem NarrowLaneWidthMenuItem;
        private ToolStripMenuItem isAbortAtLastNoteItem;

        private CenteredToolStrip MainToolStrip;
        private CenteredToolStrip NoteToolStrip;

        private SoundPreviewManager PreviewManager { get; }
        private SoundSource CurrentMusicSource;

        private ExportManager ExportManager { get; } = new ExportManager();
        private Plugins.PluginManager PluginManager { get; } = Plugins.PluginManager.GetInstance();

        private bool IsPreviewMode
        {
            get { return isPreviewMode; }
            set
            {
                isPreviewMode = value;
                NoteView.Editable = CanEdit;
                NoteView.LaneBorderLightColor = isPreviewMode ? Color.FromArgb(40, 40, 40) : Color.FromArgb(60, 60, 60);
                NoteView.LaneBorderDarkColor = isPreviewMode ? Color.FromArgb(10, 10, 10) : Color.FromArgb(30, 30, 30);
                NoteView.UnitLaneWidth = isPreviewMode ? 4 : ApplicationSettings.Default.UnitLaneWidth;
                NoteView.ShortNoteHeight = isPreviewMode ? 4 : 5;
                NoteView.UnitBeatHeight = isPreviewMode ? 48 : ApplicationSettings.Default.UnitBeatHeight;
                UpdateThumbHeight();
                ZoomInButton.Enabled = CanZoomIn;
                ZoomOutButton.Enabled = CanZoomOut;
                WidenLaneWidthMenuItem.Enabled = CanWidenLaneWidth;
                NarrowLaneWidthMenuItem.Enabled = CanNarrowLaneWidth;
            }
        }

        private bool CanWidenLaneWidth => !IsPreviewMode && NoteView.UnitLaneWidth < 150;
        private bool CanNarrowLaneWidth => !IsPreviewMode && NoteView.UnitLaneWidth > 30;
        private bool CanZoomIn => !IsPreviewMode && NoteView.UnitBeatHeight < 960;
        private bool CanZoomOut => !IsPreviewMode && NoteView.UnitBeatHeight > 30;
        private bool CanEdit => !IsPreviewMode && !PreviewManager.Playing;

        public MainForm()
        {
            InitializeComponent();
            BackColor = Color.FromArgb(45, 45, 48);
            Size = new Size(1200, 700);
            Icon = Resources.MainIcon;

            // If the setting value is below the new lower limit, update it.
            if (ApplicationSettings.Default.UnitLaneWidth < 20)
            {
                ApplicationSettings.Default.UnitLaneWidth = 100;
            }

            ToolStripManager.Renderer = new ToolStripProfessionalRenderer(new DarkColorTable());

            OperationManager = new OperationManager();
            OperationManager.OperationHistoryChanged += (s, e) =>
            {
                SetText(ScoreBook.Path);
                NoteView.Invalidate();
            };
            OperationManager.ChangesCommitted += (s, e) => SetText(ScoreBook.Path);

            NoteView = new NoteView(OperationManager)
            {
                Dock = DockStyle.Fill,
                UnitBeatHeight = ApplicationSettings.Default.UnitBeatHeight,
                UnitLaneWidth = ApplicationSettings.Default.UnitLaneWidth,
                InsertAirWithAirAction = ApplicationSettings.Default.InsertAirWithAirAction
            };

            PreviewManager = new SoundPreviewManager(this);
            PreviewManager.IsStopAtLastNote = ApplicationSettings.Default.IsPreviewAbortAtLastNote;
            PreviewManager.TickUpdated += (s, e) => NoteView.CurrentTick = e.Tick;
            PreviewManager.ExceptionThrown += (s, e) => MessageBox.Show(this, ErrorStrings.PreviewException, ApplicationTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);

            NoteViewScrollBar = new DarkScrollBar()
            {
                Dock = DockStyle.Right,
                Minimum = -NoteView.UnitBeatTick * 4 * 20,
                SmallChange = NoteView.UnitBeatTick
            };

            void processScrollBarRangeExtension(DarkScrollBar s)
            {
                if (NoteViewScrollBar.Value < NoteViewScrollBar.Minimum * 0.9f)
                {
                    NoteViewScrollBar.Minimum = (int)(NoteViewScrollBar.Minimum * 1.2);
                }
            }

            NoteView.Resize += (s, e) => UpdateThumbHeight();

            NoteView.MouseWheel += (s, e) =>
            {
                int value = NoteViewScrollBar.Value - e.Delta / 120 * NoteViewScrollBar.SmallChange;
                NoteViewScrollBar.Value = Math.Min(Math.Max(value, NoteViewScrollBar.Minimum), NoteViewScrollBar.GetMaximumValue());
                processScrollBarRangeExtension(NoteViewScrollBar);
            };

            NoteView.DragScroll += (s, e) =>
            {
                NoteViewScrollBar.Value = Math.Max(-NoteView.HeadTick, NoteViewScrollBar.Minimum);
                processScrollBarRangeExtension(NoteViewScrollBar);
            };

            NoteViewScrollBar.ValueChanged += (s, e) =>
            {
                NoteView.HeadTick = -NoteViewScrollBar.Value / 60 * 60; // 60の倍数できれいに表示されるので…
                NoteView.Invalidate();
            };

            NoteViewScrollBar.Scroll += (s, e) =>
            {
                if (e.Type == ScrollEventType.EndScroll)
                {
                    processScrollBarRangeExtension(NoteViewScrollBar);
                }
            };

            NoteView.NewNoteTypeChanged += (s, e) => NoteView.EditMode = EditMode.Edit;

            AllowDrop = true;
            DragEnter += (s, e) =>
            {
                e.Effect = DragDropEffects.None;
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    var items = (string[])e.Data.GetData(DataFormats.FileDrop);
                    if (items.Length == 1 && items.All(p => Path.GetExtension(p) == FileExtension && File.Exists(p)))
                        e.Effect = DragDropEffects.Copy;
                }
            };
            DragDrop += (s, e) =>
            {
                string path = ((string[])e.Data.GetData(DataFormats.FileDrop)).Single();
                if (!ConfirmDiscardChanges()) return;
                LoadFile(path);
            };

            FormClosing += (s, e) =>
            {
                if (!ConfirmDiscardChanges())
                {
                    e.Cancel = true;
                    return;
                }

                ApplicationSettings.Default.Save();
            };

            using (var manager = this.WorkWithLayout())
            {
                var menuStrip = CreateMainMenu(NoteView);
                menuStrip.Dock = DockStyle.Top;
                this.MainMenuStrip = menuStrip;
                NoteToolStrip = (CenteredToolStrip)CreateNewNoteTypeToolStrip(NoteView);
                MainToolStrip = (CenteredToolStrip)CreateMainToolStrip(NoteView);

                this.Controls.Add(NoteView);
                this.Controls.Add(NoteViewScrollBar);
                this.Controls.Add(NoteToolStrip);
                this.Controls.Add(MainToolStrip);
                this.Controls.Add(menuStrip);
            }

            this.ClientSizeChanged += (s, e) =>
            {
                // Defer recentering to ensure the layout pass is complete.
                this.BeginInvoke((Action)(() =>
                {
                    if (IsDisposed) return;
                    if (MainMenuStrip != null) MainMenuStrip.ForeColor = Color.White;
                }));
            };

            NoteView.NewNoteType = NoteType.Tap;
            NoteView.EditMode = EditMode.Edit;

            LoadEmptyBook();
            SetText();

            if (PluginManager.FailedFiles.Count > 0)
            {
                MessageBox.Show(this, string.Join("\n", new[] { ErrorStrings.PluginLoadError }.Concat(PluginManager.FailedFiles)), ApplicationTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            if (PluginManager.InvalidFiles.Count > 0)
            {
                MessageBox.Show(this, string.Join("\n", new[] { ErrorStrings.PluginNotSupported }.Concat(PluginManager.InvalidFiles)), ApplicationTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            if (!PreviewManager.IsSupported)
                MessageBox.Show(this, ErrorStrings.PreviewNotSupported, ApplicationTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public MainForm(string filePath) : this()
        {
            LoadFile(filePath);
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            // Defer initial centering as well to ensure all controls are properly initialized.
            this.BeginInvoke((Action)(() => {
                if (IsDisposed) return;
                if (MainMenuStrip != null) MainMenuStrip.ForeColor = Color.White;
            }));
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Space)
            {
                TogglePlay();
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        protected void LoadFile(string filePath)
        {
            try
            {
                if (!ScoreBook.IsCompatible(filePath))
                {
                    MessageBox.Show(this, ErrorStrings.FileNotCompatible, ApplicationTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (ScoreBook.IsUpgradeNeeded(filePath))
                {
                    if (MessageBox.Show(this, ErrorStrings.FileUpgradeNeeded, ApplicationTitle, MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No)
                        return;
                }
                LoadBook(ScoreBook.LoadFile(filePath));
            }
            catch (UnauthorizedAccessException)
            {
                MessageBox.Show(this, ErrorStrings.FileNotAccessible, ApplicationTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
                LoadEmptyBook();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ErrorStrings.FileLoadError, ApplicationTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
                Program.DumpExceptionTo(ex, "file_exception.json");
                LoadEmptyBook();
            }
        }

        protected void LoadBook(ScoreBook book)
        {
            ScoreBook = book;
            OperationManager.Clear();
            ExportManager.Load(book);
            NoteView.Initialize(book.Score);
            NoteViewScrollBar.Value = NoteViewScrollBar.GetMaximumValue();
            NoteViewScrollBar.Minimum = -Math.Max(NoteView.UnitBeatTick * 4 * 20, NoteView.Notes.GetLastTick());
            NoteViewScrollBar.SmallChange = NoteView.UnitBeatTick;
            UpdateThumbHeight();
            SetText(book.Path);
            CurrentMusicSource = new SoundSource();
            if (!string.IsNullOrEmpty(book.Path))
            {
                SoundSettings.Default.ScoreSound.TryGetValue(book.Path, out SoundSource src);
                if (src != null) CurrentMusicSource = src;
            }
        }

        protected void LoadEmptyBook()
        {
            var book = new ScoreBook();
            var events = book.Score.Events;
            events.BpmChangeEvents.Add(new BpmChangeEvent() { Tick = 0, Bpm = 120 });
            events.TimeSignatureChangeEvents.Add(new TimeSignatureChangeEvent() { Tick = 0, Numerator = 4, DenominatorExponent = 2 });
            LoadBook(book);
        }

        protected void OpenFile()
        {
            if (!ConfirmDiscardChanges()) return;
            if (!TrySelectOpeningFile(FileTypeFilter, out string path)) return;
            LoadFile(path);
        }

        protected bool TrySelectOpeningFile(string filter, out string path)
        {
            path = null;

            var dialog = new OpenFileDialog()
            {
                Filter = filter
            };

            if (dialog.ShowDialog(this) == DialogResult.OK)
            {
                path = dialog.FileName;
                return true;
            }
            return false;
        }

        protected void SaveAs()
        {
            var dialog = new SaveFileDialog()
            {
                Filter = FileTypeFilter
            };

            if (dialog.ShowDialog(this) == DialogResult.OK)
            {
                ScoreBook.Path = dialog.FileName;
                SaveFile();
                SetText(ScoreBook.Path);
            }
        }

        protected void SaveFile()
        {
            if (string.IsNullOrEmpty(ScoreBook.Path))
            {
                SaveAs();
                return;
            }
            CommitChanges();
            ScoreBook.Save();
            OperationManager.CommitChanges();

            SoundSettings.Default.ScoreSound[ScoreBook.Path] = CurrentMusicSource;
            SoundSettings.Default.Save();
        }

        protected void ExportAs(IScoreBookExportPlugin exportPlugin)
        {
            var dialog = new SaveFileDialog() { Filter = exportPlugin.FileFilter };
            if (dialog.ShowDialog(this) != DialogResult.OK) return;

            HandleExport(ScoreBook, ExportManager.PrepareExport(exportPlugin, dialog.FileName));
        }

        private void HandleExport(ScoreBook book, ExportContext context)
        {
            CommitChanges();
            string message;
            bool hasError = true;
            try
            {
                context.Export(book);
                message = ErrorStrings.ExportComplete;
                hasError = false;
                ExportManager.CommitExported(context);
            }
            catch (UserCancelledException)
            {
                // Do nothing
                return;
            }
            catch (InvalidTimeSignatureException ex)
            {
                int beatAt = ex.Tick / ScoreBook.Score.TicksPerBeat + 1;
                message = string.Format(ErrorStrings.InvalidTimeSignature, beatAt);
            }
            catch (Exception ex)
            {
                Program.DumpExceptionTo(ex, "export_exception.json");
                message = ErrorStrings.ExportFailed + Environment.NewLine + ex.Message;
            }

            ShowDiagnosticsResult(MainFormStrings.Export, message, hasError, context.Diagnostics);
        }

        protected void HandleImport(IScoreBookImportPlugin plugin, ScoreBookImportPluginArgs args)
        {
            string message;
            bool hasError = true;
            try
            {
                var book = plugin.Import(args);
                LoadBook(book);
                message = ErrorStrings.ImportComplete;
                hasError = false;
            }
            catch (Exception ex)
            {
                Program.DumpExceptionTo(ex, "import_exception.json");
                LoadEmptyBook();
                message = ErrorStrings.ImportFailed + Environment.NewLine + ex.Message;
            }

            ShowDiagnosticsResult(MainFormStrings.Import, message, hasError, args.Diagnostics);
        }

        protected void ShowDiagnosticsResult(string title, string message, bool hasError, IReadOnlyCollection<Diagnostic> diagnostics)
        {
            if (diagnostics.Count > 0)
            {
                var vm = new DiagnosticsWindowViewModel()
                {
                    Title = title,
                    Message = message,
                    Diagnostics = new System.Collections.ObjectModel.ObservableCollection<Diagnostic>(diagnostics)
                };
                var window = new DiagnosticsWindow()
                {
                    DataContext = vm
                };
                window.ShowDialog(this);
            }
            else
            {
                MessageBox.Show(this, message, ApplicationTitle, MessageBoxButtons.OK, hasError ? MessageBoxIcon.Error : MessageBoxIcon.Information);
            }
        }

        protected void CommitChanges()
        {
            ScoreBook.Score.Notes = NoteView.Notes.Reposit();
            // Eventsは参照渡ししてますよん
        }

        protected void ClearFile()
        {
            if (!ConfirmDiscardChanges()) return;
            LoadEmptyBook();
        }

        protected bool ConfirmDiscardChanges()
        {
            if (!OperationManager.IsChanged) return true;
            return MessageBox.Show(this, ErrorStrings.FileDiscardConfirmation, ApplicationTitle, MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) == DialogResult.OK;
        }

        protected void SetText()
        {
            SetText(null);
        }

        protected void SetText(string filePath)
        {
            Text = ApplicationTitle + (string.IsNullOrEmpty(filePath) ? "" : " - " + Path.GetFileName(filePath)) + (OperationManager.IsChanged ? " *" : "");
        }

        private void TogglePlay()
        {
            if (string.IsNullOrEmpty(CurrentMusicSource?.FilePath))
            {
                MessageBox.Show(this, ErrorStrings.MusicSourceNull, ApplicationTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (!File.Exists(CurrentMusicSource.FilePath))
            {
                MessageBox.Show(this, ErrorStrings.SourceFileNotFound, ApplicationTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (PreviewManager.Playing)
            {
                PreviewManager.Stop();
                return;
            }

            int startTick = NoteView.CurrentTick;
            void lambda(object p, EventArgs q)
            {
                isAbortAtLastNoteItem.Enabled = true;
                PreviewManager.Finished -= lambda;
                NoteView.CurrentTick = startTick;
                NoteView.Editable = CanEdit;
            }

            try
            {
                CommitChanges();
                var context = new SoundPreviewContext(ScoreBook.Score, CurrentMusicSource);
                if (!PreviewManager.Start(context, startTick)) return;
                isAbortAtLastNoteItem.Enabled = false;
                PreviewManager.Finished += lambda;
                NoteView.Editable = CanEdit;
            }
            catch (Exception ex)
            {
                Program.DumpExceptionTo(ex, "sound_exception.json");
            }
        }

        private void UpdateThumbHeight()
        {
            NoteViewScrollBar.LargeChange = NoteView.TailTick - NoteView.HeadTick;
            NoteViewScrollBar.Maximum = NoteViewScrollBar.LargeChange + NoteView.PaddingHeadTick;
        }

        private MenuStrip CreateMainMenu(NoteView noteView)
        {
            var importPluginItems = PluginManager.ScoreBookImportPlugins.Select(p => new ToolStripMenuItem(p.DisplayName, null, (s, e) =>
            {
                if (!ConfirmDiscardChanges()) return;
                if (!TrySelectOpeningFile(p.FileFilter, out string path)) return;

                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read))
                {
                    var args = new ScoreBookImportPluginArgs(stream);
                    HandleImport(p, args);
                }
            })).ToArray();

            var exportPluginItems = PluginManager.ScoreBookExportPlugins.Select(p => new ToolStripMenuItem(p.DisplayName, null, (s, e) =>
            {
                ExportAs(p);
            })).ToArray();

            var bookPropertiesMenuItem = new ToolStripMenuItem(MainFormStrings.BookProperty, null, (s, e) =>
            {
                var vm = new BookPropertiesWindowViewModel(ScoreBook, CurrentMusicSource);
                var window = new BookPropertiesWindow()
                {
                    DataContext = vm
                };
                window.ShowDialog(this);
            });

            var fileMenuItems = new ToolStripItem[]
            {
                new ToolStripMenuItem(MainFormStrings.NewFile + "(&N)", null, (s, e) => ClearFile(), Keys.Control | Keys.N),
                new ToolStripMenuItem(MainFormStrings.OpenFile + "(&O)", null, (s, e) => OpenFile(), Keys.Control | Keys.O),
                new ToolStripMenuItem(MainFormStrings.SaveFile + "(&S)", null, (s, e) => SaveFile(), Keys.Control | Keys.S),
                new ToolStripMenuItem(MainFormStrings.SaveAs + "(&A)", null, (s, e) => SaveAs(), Keys.Control | Keys.Shift | Keys.S),
                new ToolStripSeparator(),
                new ToolStripMenuItem(MainFormStrings.Import, null, importPluginItems) { Enabled = importPluginItems.Length > 0 },
                new ToolStripMenuItem(MainFormStrings.Export, null, exportPluginItems) { Enabled = exportPluginItems.Length > 0 },
                new ToolStripSeparator(),
                bookPropertiesMenuItem,
                new ToolStripSeparator(),
                new ToolStripMenuItem(MainFormStrings.Exit + "(&X)", null, (s, e) => this.Close())
            };

            var undoItem = new ToolStripMenuItem(MainFormStrings.Undo, null, (s, e) => OperationManager.Undo(), Keys.Control | Keys.Z)
            {
                Enabled = false
            };
            var redoItem = new ToolStripMenuItem(MainFormStrings.Redo, null, (s, e) => OperationManager.Redo(), Keys.Control | Keys.Y)
            {
                Enabled = false
            };

            var cutItem = new ToolStripMenuItem(MainFormStrings.Cut, null, (s, e) => noteView.CutSelectedNotes(), Keys.Control | Keys.X);
            var copyItem = new ToolStripMenuItem(MainFormStrings.Copy, null, (s, e) => noteView.CopySelectedNotes(), Keys.Control | Keys.C);
            var pasteItem = new ToolStripMenuItem(MainFormStrings.Paste, null, (s, e) => noteView.PasteNotes(), Keys.Control | Keys.V);
            var pasteFlippedItem = new ToolStripMenuItem(MainFormStrings.PasteFlipped, null, (s, e) => noteView.PasteFlippedNotes(), Keys.Control | Keys.Shift | Keys.V);

            var selectAllItem = new ToolStripMenuItem(MainFormStrings.SelectAll, null, (s, e) => noteView.SelectAll(), Keys.Control | Keys.A);
            var selectToEndItem = new ToolStripMenuItem(MainFormStrings.SelectToEnd, null, (s, e) => noteView.SelectToEnd());
            var selectoToBeginningItem = new ToolStripMenuItem(MainFormStrings.SelectToBeginning, null, (s, e) => noteView.SelectToBeginning());

            var flipSelectedNotesItem = new ToolStripMenuItem(MainFormStrings.FlipSelectedNotes, null, (s, e) => noteView.FlipSelectedNotes());
            var removeSelectedNotesItem = new ToolStripMenuItem(MainFormStrings.RemoveSelectedNotes, null, (s, e) => noteView.RemoveSelectedNotes(), Keys.Delete);

            var removeEventsItem = new ToolStripMenuItem(MainFormStrings.RemoveEvents, null, (s, e) =>
            {
                int minTick = noteView.SelectedRange.StartTick + (noteView.SelectedRange.Duration < 0 ? noteView.SelectedRange.Duration : 0);
                int maxTick = noteView.SelectedRange.StartTick + (noteView.SelectedRange.Duration < 0 ? 0 : noteView.SelectedRange.Duration);
                bool isContained(EventBase p) => p.Tick != 0 && minTick <= p.Tick && maxTick >= p.Tick;
                var events = ScoreBook.Score.Events;

                var bpmOp = events.BpmChangeEvents.Where(p => isContained(p)).ToList().Select(p =>
                {
                    ScoreBook.Score.Events.BpmChangeEvents.Remove(p);
                    return new RemoveEventOperation<BpmChangeEvent>(events.BpmChangeEvents, p);
                }).ToList();

                var speedOp = events.HighSpeedChangeEvents.Where(p => isContained(p)).ToList().Select(p =>
                {
                    ScoreBook.Score.Events.HighSpeedChangeEvents.Remove(p);
                    return new RemoveEventOperation<HighSpeedChangeEvent>(events.HighSpeedChangeEvents, p);
                }).ToList();

                var signatureOp = events.TimeSignatureChangeEvents.Where(p => isContained(p)).ToList().Select(p =>
                {
                    ScoreBook.Score.Events.TimeSignatureChangeEvents.Remove(p);
                    return new RemoveEventOperation<TimeSignatureChangeEvent>(events.TimeSignatureChangeEvents, p);
                }).ToList();

                OperationManager.Push(new CompositeOperation("イベント削除", bpmOp.Cast<IOperation>().Concat(speedOp).Concat(signatureOp)));
                noteView.Invalidate();
            });

            var insertAirWithAirActionItem = new ToolStripMenuItem(MainFormStrings.InsertAirWithAirAction, null, (s, e) =>
            {
                var item = s as ToolStripMenuItem;
                item.Checked = !item.Checked;
                NoteView.InsertAirWithAirAction = item.Checked;
                ApplicationSettings.Default.InsertAirWithAirAction = item.Checked;
            })
            {
                Checked = ApplicationSettings.Default.InsertAirWithAirAction
            };

            var pluginItems = PluginManager.ScorePlugins.Select(p => new ToolStripMenuItem(p.DisplayName, null, (s, e) =>
            {
                CommitChanges();
                void updateScore(Score newScore)
                {
                    var op = new UpdateScoreOperation(ScoreBook.Score, newScore, score =>
                    {
                        ScoreBook.Score = score;
                        noteView.UpdateScore(score);
                    });
                    OperationManager.InvokeAndPush(op);
                }

                try
                {
                    p.Run(new ScorePluginArgs(() => ScoreBook.Score.Clone(), noteView.SelectedRange, updateScore));
                }
                catch (Exception ex)
                {
                    Program.DumpExceptionTo(ex, "plugin_exception.json");
                    MessageBox.Show(this, ErrorStrings.PluginException, ApplicationTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            })).ToArray();
            var pluginItem = new ToolStripMenuItem(MainFormStrings.Plugin, null, pluginItems) { Enabled = pluginItems.Length > 0 };

            var editMenuItems = new ToolStripItem[]
            {
                undoItem, redoItem, new ToolStripSeparator(),
                cutItem, copyItem, pasteItem, pasteFlippedItem, new ToolStripSeparator(),
                selectAllItem, selectToEndItem, selectoToBeginningItem, new ToolStripSeparator(),
                flipSelectedNotesItem, removeSelectedNotesItem, removeEventsItem, new ToolStripSeparator(),
                insertAirWithAirActionItem, new ToolStripSeparator(),
                pluginItem
            };

            var viewModeItem = new ToolStripMenuItem(MainFormStrings.ScorePreview, null, (s, e) =>
            {
                IsPreviewMode = !IsPreviewMode;
                ((ToolStripMenuItem)s).Checked = IsPreviewMode;
            }, Keys.Control | Keys.P);

            WidenLaneWidthMenuItem = new ToolStripMenuItem(MainFormStrings.WidenLaneWidth);
            NarrowLaneWidthMenuItem = new ToolStripMenuItem(MainFormStrings.NarrowLaneWidth);

            WidenLaneWidthMenuItem.Click += (s, e) =>
            {
                noteView.UnitLaneWidth += 4;
                ApplicationSettings.Default.UnitLaneWidth = noteView.UnitLaneWidth;
                WidenLaneWidthMenuItem.Enabled = CanWidenLaneWidth;
                NarrowLaneWidthMenuItem.Enabled = CanNarrowLaneWidth;
            };
            NarrowLaneWidthMenuItem.Click += (s, e) =>
            {
                noteView.UnitLaneWidth -= 4;
                ApplicationSettings.Default.UnitLaneWidth = noteView.UnitLaneWidth;
                WidenLaneWidthMenuItem.Enabled = CanWidenLaneWidth;
                NarrowLaneWidthMenuItem.Enabled = CanNarrowLaneWidth;
            };

            var viewMenuItems = new ToolStripItem[] {
                viewModeItem,
                new ToolStripSeparator(),
                WidenLaneWidthMenuItem, NarrowLaneWidthMenuItem
            };

            void UpdateEvent<T>(List<T> list, T item) where T : EventBase
            {
                var prev = list.SingleOrDefault(p => p.Tick == item.Tick);

                var insertOp = new InsertEventOperation<T>(list, item);
                if (prev == null)
                {
                    OperationManager.InvokeAndPush(insertOp);
                }
                else
                {
                    var removeOp = new RemoveEventOperation<T>(list, prev);
                    OperationManager.InvokeAndPush(new CompositeOperation(insertOp.Description, new IOperation[] { removeOp, insertOp }));
                }
                noteView.Invalidate();
            }

            var insertBpmItem = new ToolStripMenuItem("BPM", null, (s, e) =>
            {
                var form = new BpmSelectionForm()
                {
                    Bpm = noteView.ScoreEvents.BpmChangeEvents.OrderBy(p => p.Tick).LastOrDefault(p => p.Tick <= noteView.CurrentTick)?.Bpm ?? 120
                };
                if (form.ShowDialog(this) != DialogResult.OK) return;

                var item = new BpmChangeEvent()
                {
                    Tick = noteView.CurrentTick,
                    Bpm = form.Bpm
                };
                UpdateEvent(noteView.ScoreEvents.BpmChangeEvents, item);
            });

            var insertHighSpeedItem = new ToolStripMenuItem(MainFormStrings.HighSpeed, null, (s, e) =>
            {
                var form = new HighSpeedSelectionForm()
                {
                    SpeedRatio = noteView.ScoreEvents.HighSpeedChangeEvents.OrderBy(p => p.Tick).LastOrDefault(p => p.Tick <= noteView.CurrentTick)?.SpeedRatio ?? 1.0m
                };
                if (form.ShowDialog(this) != DialogResult.OK) return;

                var item = new HighSpeedChangeEvent()
                {
                    Tick = noteView.CurrentTick,
                    SpeedRatio = form.SpeedRatio
                };
                UpdateEvent(noteView.ScoreEvents.HighSpeedChangeEvents, item);
            });

            var insertTimeSignatureItem = new ToolStripMenuItem(MainFormStrings.TimeSignature, null, (s, e) =>
            {
                var form = new TimeSignatureSelectionForm();
                if (form.ShowDialog(this) != DialogResult.OK) return;

                var item = new TimeSignatureChangeEvent()
                {
                    Tick = noteView.CurrentTick,
                    Numerator = form.Numerator,
                    DenominatorExponent = form.DenominatorExponent
                };
                UpdateEvent(noteView.ScoreEvents.TimeSignatureChangeEvents, item);
            });

            var insertMenuItems = new ToolStripItem[] { insertBpmItem, insertHighSpeedItem, insertTimeSignatureItem };

            isAbortAtLastNoteItem = new ToolStripMenuItem(MainFormStrings.AbortAtLastNote, null, (s, e) =>
            {
                var item = s as ToolStripMenuItem;
                item.Checked = !item.Checked;
                PreviewManager.IsStopAtLastNote = item.Checked;
                ApplicationSettings.Default.IsPreviewAbortAtLastNote = item.Checked;
            })
            {
                Checked = ApplicationSettings.Default.IsPreviewAbortAtLastNote
            };

            var playItem = new ToolStripMenuItem(MainFormStrings.Play, null, (s, e) => TogglePlay());

            var stopItem = new ToolStripMenuItem(MainFormStrings.Stop, null, (s, e) =>
            {
                PreviewManager.Stop();
            });

            var playMenuItems = new ToolStripItem[]
            {
                playItem, stopItem, new ToolStripSeparator(),
                isAbortAtLastNoteItem
            };

            var helpMenuItems = new ToolStripItem[]
            {
                new ToolStripMenuItem(MainFormStrings.Help, null, (s, e) => System.Diagnostics.Process.Start("https://github.com/paralleltree/Ched/wiki"), Keys.F1),
                new ToolStripMenuItem(MainFormStrings.VersionInfo, null, (s, e) => new VersionInfoForm().ShowDialog(this))
            };

            OperationManager.OperationHistoryChanged += (s, e) =>
            {
                redoItem.Enabled = OperationManager.CanRedo;
                undoItem.Enabled = OperationManager.CanUndo;
            };

            var menuStrip = new MenuStrip();
            menuStrip.Items.AddRange(new ToolStripItem[]
            {
                new ToolStripMenuItem(MainFormStrings.FileMenu, null, fileMenuItems),
                new ToolStripMenuItem(MainFormStrings.EditMenu, null, editMenuItems),
                new ToolStripMenuItem(MainFormStrings.ViewMenu, null, viewMenuItems),
                new ToolStripMenuItem(MainFormStrings.InsertMenu, null, insertMenuItems),
                new ToolStripMenuItem(MainFormStrings.PlayMenu, null, playMenuItems) { Enabled = PreviewManager.IsSupported },
                new ToolStripMenuItem(MainFormStrings.HelpMenu, null, helpMenuItems)
            });

            Action<ToolStripItemCollection> setForeColor = null;
            setForeColor = (items) =>
            {
                foreach (var item in items.OfType<ToolStripMenuItem>())
                {
                    item.ForeColor = Color.White;
                    if (item.HasDropDown) setForeColor(item.DropDownItems);
                }
            };

            setForeColor(menuStrip.Items);

            Action<ToolStripItemCollection> stripMnemonics = null;
            stripMnemonics = (items) =>
            {
                foreach (var item in items.OfType<ToolStripMenuItem>())
                {
                    // Remove accelerator key hints like "(&V)" from the text.
                    // e.g., "View(&V)" becomes "View".
                    item.Text = Regex.Replace(item.Text, @"\s*\(&.\)$", "");
                    if (item.HasDropDown) stripMnemonics(item.DropDownItems);
                }
            };
            stripMnemonics(menuStrip.Items);

            return menuStrip;
        }

        private ToolStrip CreateMainToolStrip(NoteView noteView)
        {
            var newFileButton = new ToolStripButton(MainFormStrings.NewFile, Resources.NewFileIcon, (s, e) => ClearFile())
            {
                DisplayStyle = ToolStripItemDisplayStyle.ImageAndText,
                TextImageRelation = TextImageRelation.ImageAboveText
            };
            var openFileButton = new ToolStripButton(MainFormStrings.OpenFile, Resources.OpenFileIcon, (s, e) => OpenFile())
            {
                DisplayStyle = ToolStripItemDisplayStyle.ImageAndText,
                TextImageRelation = TextImageRelation.ImageAboveText
            };
            var saveFileButton = new ToolStripButton(MainFormStrings.SaveFile, Resources.SaveFileIcon, (s, e) => SaveFile())
            {
                DisplayStyle = ToolStripItemDisplayStyle.ImageAndText,
                TextImageRelation = TextImageRelation.ImageAboveText
            };
            var exportButton = new ToolStripButton(MainFormStrings.Export, Resources.ExportIcon, (s, e) =>
            {
                if (!ExportManager.CanReExport)
                {
                    if (PluginManager.ScoreBookExportPlugins.Count() == 1)
                    {
                        ExportAs(PluginManager.ScoreBookExportPlugins.Single());
                        return;
                    }
                    MessageBox.Show(this, ErrorStrings.NotExported, ApplicationTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                HandleExport(ScoreBook, ExportManager.PrepareReExport());
            })
            {
                DisplayStyle = ToolStripItemDisplayStyle.ImageAndText,
                TextImageRelation = TextImageRelation.ImageAboveText
            };

            var cutButton = new ToolStripButton(MainFormStrings.Cut, Resources.CutIcon, (s, e) => noteView.CutSelectedNotes())
            {
                DisplayStyle = ToolStripItemDisplayStyle.ImageAndText,
                TextImageRelation = TextImageRelation.ImageAboveText,
                Margin = new Padding(20, 0, 0, 0)
            };
            var copyButton = new ToolStripButton(MainFormStrings.Copy, Resources.CopyIcon, (s, e) => noteView.CopySelectedNotes())
            {
                DisplayStyle = ToolStripItemDisplayStyle.ImageAndText,
                TextImageRelation = TextImageRelation.ImageAboveText
            };
            var pasteButton = new ToolStripButton(MainFormStrings.Paste, Resources.PasteIcon, (s, e) => noteView.PasteNotes())
            {
                DisplayStyle = ToolStripItemDisplayStyle.ImageAndText,
                TextImageRelation = TextImageRelation.ImageAboveText
            };

            var undoButton = new ToolStripButton(MainFormStrings.Undo, Resources.UndoIcon, (s, e) => OperationManager.Undo())
            {
                DisplayStyle = ToolStripItemDisplayStyle.ImageAndText,
                TextImageRelation = TextImageRelation.ImageAboveText,
                Enabled = false,
                Margin = new Padding(20, 0, 0, 0)
            };
            var redoButton = new ToolStripButton(MainFormStrings.Redo, Resources.RedoIcon, (s, e) => OperationManager.Redo())
            {
                DisplayStyle = ToolStripItemDisplayStyle.ImageAndText,
                TextImageRelation = TextImageRelation.ImageAboveText,
                Enabled = false
            };

            var penButton = new ToolStripButton(MainFormStrings.Pen, Resources.EditIcon, (s, e) => noteView.EditMode = EditMode.Edit)
            {
                DisplayStyle = ToolStripItemDisplayStyle.ImageAndText,
                TextImageRelation = TextImageRelation.ImageAboveText,
                Margin = new Padding(20, 0, 0, 0)
            };
            var selectionButton = new ToolStripButton(MainFormStrings.Selection, Resources.SelectionIcon, (s, e) => noteView.EditMode = EditMode.Select)
            {
                DisplayStyle = ToolStripItemDisplayStyle.ImageAndText,
                TextImageRelation = TextImageRelation.ImageAboveText
            };
            var eraserButton = new ToolStripButton(MainFormStrings.Eraser, Resources.EraserIcon, (s, e) => noteView.EditMode = EditMode.Erase)
            {
                DisplayStyle = ToolStripItemDisplayStyle.ImageAndText,
                TextImageRelation = TextImageRelation.ImageAboveText
            };

            var zoomInButton = new ToolStripButton(MainFormStrings.ZoomIn, Resources.ZoomInIcon)
            {
                Enabled = noteView.UnitBeatHeight < 1920,
                DisplayStyle = ToolStripItemDisplayStyle.ImageAndText,
                TextImageRelation = TextImageRelation.ImageAboveText,
                Margin = new Padding(20, 0, 0, 0)
            };
            var zoomOutButton = new ToolStripButton(MainFormStrings.ZoomOut, Resources.ZoomOutIcon)
            {
                Enabled = noteView.UnitBeatHeight > 30,
                DisplayStyle = ToolStripItemDisplayStyle.ImageAndText,
                TextImageRelation = TextImageRelation.ImageAboveText
            };

            zoomInButton.Click += (s, e) =>
            {
                noteView.UnitBeatHeight *= 2;
                ApplicationSettings.Default.UnitBeatHeight = (int)noteView.UnitBeatHeight;
                zoomOutButton.Enabled = CanZoomOut;
                zoomInButton.Enabled = CanZoomIn;
                UpdateThumbHeight();
            };

            zoomOutButton.Click += (s, e) =>
            {
                noteView.UnitBeatHeight /= 2;
                ApplicationSettings.Default.UnitBeatHeight = (int)noteView.UnitBeatHeight;
                zoomInButton.Enabled = CanZoomIn;
                zoomOutButton.Enabled = CanZoomOut;
                UpdateThumbHeight();
            };

            ZoomInButton = zoomInButton;
            ZoomOutButton = zoomOutButton;

            OperationManager.OperationHistoryChanged += (s, e) =>
            {
                undoButton.Enabled = OperationManager.CanUndo;
                redoButton.Enabled = OperationManager.CanRedo;
            };

            noteView.EditModeChanged += (s, e) =>
            {
                selectionButton.Checked = noteView.EditMode == EditMode.Select;
                penButton.Checked = noteView.EditMode == EditMode.Edit;
                eraserButton.Checked = noteView.EditMode == EditMode.Erase;
            };

            var spacer = new ToolStripLabel("") { AutoSize = false, Width = 1 };
            var items = new List<ToolStripItem>
            {
                spacer,
                newFileButton, openFileButton, saveFileButton, exportButton,
                cutButton, copyButton, pasteButton,
                undoButton, redoButton,
                penButton, selectionButton, eraserButton,
                zoomInButton, zoomOutButton
            };

            var mainStrip = new CenteredToolStrip(items.ToArray());
            mainStrip.ImageScalingSize = new Size(32, 32);
            mainStrip.ForeColor = Color.White;

            return mainStrip;
        }

        private ToolStrip CreateNewNoteTypeToolStrip(NoteView noteView)
        {
            var tapButton = new ToolStripButton("Tap Note", Resources.TapIcon, (s, e) => noteView.NewNoteType = NoteType.Tap)
            {
                DisplayStyle = ToolStripItemDisplayStyle.ImageAndText,
                TextImageRelation = TextImageRelation.ImageAboveText
            };
            var exTapButton = new ToolStripButton("Golden Tap Note", Resources.ExTapIcon, (s, e) => noteView.NewNoteType = NoteType.ExTap)
            {
                DisplayStyle = ToolStripItemDisplayStyle.ImageAndText,
                TextImageRelation = TextImageRelation.ImageAboveText
            };
            var holdButton = new ToolStripButton("Hold Note", Resources.HoldIcon, (s, e) => noteView.NewNoteType = NoteType.Hold)
            {
                DisplayStyle = ToolStripItemDisplayStyle.ImageAndText,
                TextImageRelation = TextImageRelation.ImageAboveText
            };
            var damageButton = new ToolStripButton("Damage Note", Resources.DamgeIcon, (s, e) => noteView.NewNoteType = NoteType.Damage)
            {
                DisplayStyle = ToolStripItemDisplayStyle.ImageAndText,
                TextImageRelation = TextImageRelation.ImageAboveText
            };

            var quantizeTicks = new int[]
            {
                4, 8, 12, 16, 24, 32, 48, 64, 96, 128, 192
            };
            var quantizeComboBox = new ToolStripComboBox("クォンタイズ")
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                AutoSize = false,
                Width = 80,
                Margin = new Padding(4, 0, 0, 0),
                BackColor = Color.FromArgb(45, 45, 48),
                ForeColor = Color.White
            };
            quantizeComboBox.Items.AddRange(quantizeTicks.Select(p => p + MainFormStrings.Division).ToArray());
            quantizeComboBox.Items.Add(MainFormStrings.Custom);

            quantizeComboBox.ComboBox.DrawMode = DrawMode.OwnerDrawFixed;
            quantizeComboBox.ComboBox.DrawItem += (s, e) =>
            {
                if (e.Index < 0) return;
                var combo = (ComboBox)s;
                var itemText = combo.GetItemText(combo.Items[e.Index]);

                Color backColor = (e.State & DrawItemState.Selected) != 0 ? Color.FromArgb(0, 120, 215) : Color.FromArgb(45, 45, 48);
                using (var b = new SolidBrush(backColor))
                {
                    e.Graphics.FillRectangle(b, e.Bounds);
                }

                TextRenderer.DrawText(e.Graphics, itemText, e.Font, e.Bounds, Color.White, TextFormatFlags.Left | TextFormatFlags.VerticalCenter);

                e.DrawFocusRectangle();
            };

            quantizeComboBox.ComboBox.DropDown += (s, e) =>
            {
                var comboBox = (ComboBox)s;
                int maxWidth = 0;
                using (Graphics g = comboBox.CreateGraphics())
                {
                    foreach (var item in comboBox.Items)
                    {
                        int width = (int)g.MeasureString(comboBox.GetItemText(item), comboBox.Font).Width;
                        if (width > maxWidth) maxWidth = width;
                    }
                }

                if (comboBox.Items.Count > comboBox.MaxDropDownItems) maxWidth += SystemInformation.VerticalScrollBarWidth;

                comboBox.DropDownWidth = Math.Max(comboBox.Width, maxWidth + 10);
            };

            quantizeComboBox.SelectedIndexChanged += (s, e) =>
            {
                if (quantizeComboBox.SelectedIndex == quantizeComboBox.Items.Count - 1)
                {
                    // ユーザー定義
                    var form = new CustomQuantizeSelectionForm(ScoreBook.Score.TicksPerBeat * 4);
                    if (form.ShowDialog(this) == DialogResult.OK)
                    {
                        noteView.QuantizeTick = form.QuantizeTick;
                    }
                }
                else
                {
                    noteView.QuantizeTick = noteView.UnitBeatTick * 4 / quantizeTicks[quantizeComboBox.SelectedIndex];
                }
                noteView.Focus();
            };
            quantizeComboBox.SelectedIndex = 1;

            noteView.NewNoteTypeChanged += (s, e) =>
            {
                tapButton.Checked = noteView.NewNoteType.HasFlag(NoteType.Tap);
                exTapButton.Checked = noteView.NewNoteType.HasFlag(NoteType.ExTap);
                holdButton.Checked = noteView.NewNoteType.HasFlag(NoteType.Hold);
                damageButton.Checked = noteView.NewNoteType.HasFlag(NoteType.Damage);
            };

            var spacer = new ToolStripLabel("") { AutoSize = false, Width = 1 };
            var items = new List<ToolStripItem>
            {
                spacer,
                tapButton, exTapButton, holdButton, damageButton,
                quantizeComboBox
            };

            var noteStrip = new CenteredToolStrip(items.ToArray());
            noteStrip.ImageScalingSize = new Size(32, 32);
            noteStrip.ForeColor = Color.White;

            return noteStrip;
        }
    }

    public class CenteredToolStrip : ToolStrip
    {
        public CenteredToolStrip(params ToolStripItem[] items) : base(items)
        {
            this.Layout += (s, e) => RecenterItems();
        }

        public void RecenterItems()
        {
            if (IsDisposed || Items.Count == 0) return;

            // Find the spacer label.
            var spacer = Items.OfType<ToolStripLabel>().FirstOrDefault(p => p.Text == "");
            if (spacer == null) return;

            // Calculate the total width of all items except the spacer.
            var totalItemsWidth = Items.Cast<ToolStripItem>()
                                      .Where(item => item != spacer)
                                      .Sum(item => item.Width + item.Margin.Horizontal);

            // Calculate the remaining width and set the spacer.
            var remainingWidth = DisplayRectangle.Width - totalItemsWidth;
            spacer.Width = Math.Max(0, remainingWidth / 2);
        }
    }
}