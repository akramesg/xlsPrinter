using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using XlsPrinter.Core;
using XlsPrinter.Spooler;

namespace XlsPrinter.UI
{
    public class MainForm : Form
    {
        private VirtualPrinterManager _printerManager;
        private EmfSpoolParser _spoolParser;
        private TableLayoutEngine _layoutEngine;

        private PrintJobDocument _currentDocument;
        private ReconstructedTable _currentTable;
        private int _currentPageIndex = 0;

        // Interactive region crop selection on the visual page preview (in mm)
        private RectangleF? _selectionRegionMm = null;
        private bool _isDraggingSelection = false;
        private Point _dragStartScreenPt;
        private Point _dragCurrentScreenPt;

        // UI Controls
        private ToolStrip _topToolbar;
        private ToolStripStatusLabel _statusLabel;
        private ToolStripStatusLabel _printerStateLabel;
        private SplitContainer _mainSplit;
        private PictureBox _previewCanvas;
        private DataGridView _dataGrid;
        private CheckedListBox _columnsCheckList;
        private CheckBox _chkRtlSheet;
        private CheckBox _chkReverseColumns;
        private CheckBox _chkReverseVisualArabic;
        private CheckBox _chkNormalizeArabicForms;
        private NumericUpDown _numRowTol;
        private NumericUpDown _numColTol;
        private Label _lblSelectionInfo;
        private Button _btnClearRegion;
        private NotifyIcon _trayIcon;
        private bool _isExplicitExit = false;
        private bool _startInTray = false;

        public MainForm(string initialFilePath) : this(initialFilePath, false)
        {
        }

        public MainForm(string initialFilePath, bool startInTray)
        {
            _startInTray = startInTray;
            _printerManager = new VirtualPrinterManager();
            _spoolParser = new EmfSpoolParser();
            _layoutEngine = new TableLayoutEngine();

            InitializeComponents();

            _printerManager.JobCaptured += OnPrintJobCaptured;
            _printerManager.StartWatching();
            UpdatePrinterInstallStatus();

            if (!string.IsNullOrEmpty(initialFilePath) && File.Exists(initialFilePath))
            {
                LoadSpoolOrEmfFile(initialFilePath);
            }
            else
            {
                LoadBuiltInSamplePrintJob();
            }
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            if (_startInTray)
            {
                this.BeginInvoke(new MethodInvoker(delegate { this.Hide(); }));
            }
        }

        private void InitializeComponents()
        {
            this.Text = "XLS Virtual Printer - Win32 GDI Print-to-Excel Converter (Arabic & Legacy ERP Edition)";
            this.Size = new Size(1220, 780);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Font = new Font("Tahoma", 9f);

            // Top Toolbar
            _topToolbar = new ToolStrip();
            _topToolbar.GripStyle = ToolStripGripStyle.Hidden;
            _topToolbar.Padding = new Padding(4, 2, 4, 2);

            ToolStripButton btnInstallPrinter = new ToolStripButton("Install Virtual Printer");
            btnInstallPrinter.Font = new Font("Tahoma", 9f, FontStyle.Bold);
            btnInstallPrinter.ForeColor = Color.DarkBlue;
            btnInstallPrinter.Click += delegate { InstallOrRepairVirtualPrinter(); };

            ToolStripButton btnTestPrint = new ToolStripButton("Simulate Classic Arabic Print");
            btnTestPrint.Click += delegate { SendActualGdiTestPrintOrSimulate(); };

            ToolStripButton btnOpenSpl = new ToolStripButton("Open .SPL / .EMF / .PRN...");
            btnOpenSpl.Click += delegate { PromptOpenSpoolFile(); };

            ToolStripSeparator sep1 = new ToolStripSeparator();

            ToolStripButton btnSelectAllRows = new ToolStripButton("Select All Rows");
            btnSelectAllRows.Click += delegate { SetAllRowsSelected(true); };

            ToolStripButton btnDeselectAllRows = new ToolStripButton("Clear Row Selection");
            btnDeselectAllRows.Click += delegate { SetAllRowsSelected(false); };

            ToolStripSeparator sep2 = new ToolStripSeparator();

            ToolStripButton btnExportXls = new ToolStripButton("Export Selected to .XLS (97-2003)");
            btnExportXls.Font = new Font("Tahoma", 9f, FontStyle.Bold);
            btnExportXls.ForeColor = Color.DarkGreen;
            btnExportXls.Click += delegate { ExportCurrentTable(".xls"); };

            ToolStripButton btnExportXlsx = new ToolStripButton("Export to .XLSX");
            btnExportXlsx.Font = new Font("Tahoma", 9f, FontStyle.Bold);
            btnExportXlsx.Click += delegate { ExportCurrentTable(".xlsx"); };

            ToolStripButton btnExportCsv = new ToolStripButton("Export to .CSV (UTF-8)");
            btnExportCsv.Click += delegate { ExportCurrentTable(".csv"); };

            ToolStripSeparator sep3 = new ToolStripSeparator();
            ToolStripButton btnViewLog = new ToolStripButton("View Debug Log");
            btnViewLog.Click += delegate {
                try {
                    string logFile = AppLogger.LogFilePath;
                    if (!File.Exists(logFile)) File.WriteAllText(logFile, "Debug log started.\r\n");
                    System.Diagnostics.Process.Start("notepad.exe", logFile);
                } catch { }
            };

            _topToolbar.Items.AddRange(new ToolStripItem[] {
                btnInstallPrinter, btnTestPrint, btnOpenSpl, sep1,
                btnSelectAllRows, btnDeselectAllRows, sep2,
                btnExportXls, btnExportXlsx, btnExportCsv, sep3, btnViewLog
            });

            // Status Bar
            StatusStrip statusBar = new StatusStrip();
            _printerStateLabel = new ToolStripStatusLabel("Checking printer...");
            _printerStateLabel.BorderSides = ToolStripStatusLabelBorderSides.Right;
            _statusLabel = new ToolStripStatusLabel("Ready");
            _statusLabel.Spring = true;
            _statusLabel.TextAlign = ContentAlignment.MiddleLeft;
            statusBar.Items.Add(_printerStateLabel);
            statusBar.Items.Add(_statusLabel);

            // Options Panel (Right / Top of Grid)
            Panel optionsPanel = new Panel();
            optionsPanel.Dock = DockStyle.Top;
            optionsPanel.Height = 115;
            optionsPanel.BackColor = Color.FromArgb(244, 246, 250);
            optionsPanel.Padding = new Padding(8);

            GroupBox grpArabic = new GroupBox();
            grpArabic.Text = "Arabic & Layout Options";
            grpArabic.Location = new Point(8, 6);
            grpArabic.Size = new Size(440, 100);

            _chkRtlSheet = new CheckBox();
            _chkRtlSheet.Text = "Right-to-Left (RTL) Excel Sheet View";
            _chkRtlSheet.Checked = false;
            _chkRtlSheet.AutoSize = true;
            _chkRtlSheet.Location = new Point(12, 22);

            _chkReverseColumns = new CheckBox();
            _chkReverseColumns.Text = "Order Columns Right-to-Left (Arabic Table)";
            _chkReverseColumns.Checked = false;
            _chkReverseColumns.AutoSize = true;
            _chkReverseColumns.Location = new Point(12, 46);
            _chkReverseColumns.CheckedChanged += delegate { RebuildTableFromDocument(); };

            _chkNormalizeArabicForms = new CheckBox();
            _chkNormalizeArabicForms.Text = "Normalize Arabic Glyph Forms (FE70->0600)";
            _chkNormalizeArabicForms.Checked = true;
            _chkNormalizeArabicForms.AutoSize = true;
            _chkNormalizeArabicForms.Location = new Point(12, 70);
            _chkNormalizeArabicForms.CheckedChanged += delegate { ReloadDocumentWithArabicSettings(); };

            _chkReverseVisualArabic = new CheckBox();
            _chkReverseVisualArabic.Text = "Reverse Visual Arabic Chars";
            _chkReverseVisualArabic.Checked = true;
            _chkReverseVisualArabic.AutoSize = true;
            _chkReverseVisualArabic.Location = new Point(250, 22);
            _chkReverseVisualArabic.CheckedChanged += delegate { ReloadDocumentWithArabicSettings(); };

            Label lblRowTol = new Label();
            lblRowTol.Text = "Row Y Tol (mm):";
            lblRowTol.AutoSize = true;
            lblRowTol.Location = new Point(250, 48);

            _numRowTol = new NumericUpDown();
            _numRowTol.DecimalPlaces = 1;
            _numRowTol.Increment = 0.5m;
            _numRowTol.Minimum = 0.5m;
            _numRowTol.Maximum = 15.0m;
            _numRowTol.Value = 2.2m;
            _numRowTol.Width = 55;
            _numRowTol.Location = new Point(355, 45);
            _numRowTol.ValueChanged += delegate {
                _layoutEngine.RowVerticalToleranceMm = (float)_numRowTol.Value;
                RebuildTableFromDocument();
            };

            Label lblColTol = new Label();
            lblColTol.Text = "Col X Tol (mm):";
            lblColTol.AutoSize = true;
            lblColTol.Location = new Point(250, 74);

            _numColTol = new NumericUpDown();
            _numColTol.DecimalPlaces = 1;
            _numColTol.Increment = 1.0m;
            _numColTol.Minimum = 1.0m;
            _numColTol.Maximum = 40.0m;
            _numColTol.Value = 8.0m;
            _numColTol.Width = 55;
            _numColTol.Location = new Point(355, 71);
            _numColTol.ValueChanged += delegate {
                _layoutEngine.ColumnHorizontalToleranceMm = (float)_numColTol.Value;
                RebuildTableFromDocument();
            };

            grpArabic.Controls.AddRange(new Control[] {
                _chkRtlSheet, _chkReverseColumns, _chkNormalizeArabicForms,
                _chkReverseVisualArabic, lblRowTol, _numRowTol, lblColTol, _numColTol
            });

            GroupBox grpColumns = new GroupBox();
            grpColumns.Text = "Columns to Export (Check/Uncheck)";
            grpColumns.Location = new Point(456, 6);
            grpColumns.Size = new Size(260, 100);

            _columnsCheckList = new CheckedListBox();
            _columnsCheckList.Dock = DockStyle.Fill;
            _columnsCheckList.CheckOnClick = true;
            _columnsCheckList.MultiColumn = true;
            _columnsCheckList.ColumnWidth = 75;
            _columnsCheckList.ItemCheck += OnColumnItemCheck;
            grpColumns.Controls.Add(_columnsCheckList);

            optionsPanel.Controls.Add(grpArabic);
            optionsPanel.Controls.Add(grpColumns);

            // SplitContainer: Left = Visual Page Preview + Rubberband Crop, Right = Interactive DataGridView
            _mainSplit = new SplitContainer();
            _mainSplit.Dock = DockStyle.Fill;
            _mainSplit.SplitterDistance = 470;

            // Left Panel: Visual Page Preview
            Panel leftHeader = new Panel();
            leftHeader.Dock = DockStyle.Top;
            leftHeader.Height = 34;
            leftHeader.BackColor = Color.FromArgb(230, 235, 242);

            _lblSelectionInfo = new Label();
            _lblSelectionInfo.Text = "Visual Page Preview (Drag mouse on page to filter area/table): Full Page";
            _lblSelectionInfo.AutoSize = false;
            _lblSelectionInfo.Location = new Point(8, 8);
            _lblSelectionInfo.Size = new Size(340, 20);

            _btnClearRegion = new Button();
            _btnClearRegion.Text = "Reset Area Filter";
            _btnClearRegion.Size = new Size(110, 24);
            _btnClearRegion.Location = new Point(350, 5);
            _btnClearRegion.Enabled = false;
            _btnClearRegion.Click += delegate {
                _selectionRegionMm = null;
                _btnClearRegion.Enabled = false;
                _lblSelectionInfo.Text = "Visual Page Preview (Drag mouse to select area): Full Page";
                RebuildTableFromDocument();
                _previewCanvas.Invalidate();
            };

            leftHeader.Controls.Add(_lblSelectionInfo);
            leftHeader.Controls.Add(_btnClearRegion);

            _previewCanvas = new PictureBox();
            _previewCanvas.Dock = DockStyle.Fill;
            _previewCanvas.BackColor = Color.DarkGray;
            _previewCanvas.Paint += OnPreviewCanvasPaint;
            _previewCanvas.MouseDown += OnPreviewMouseDown;
            _previewCanvas.MouseMove += OnPreviewMouseMove;
            _previewCanvas.MouseUp += OnPreviewMouseUp;
            _previewCanvas.Resize += delegate { _previewCanvas.Invalidate(); };

            _mainSplit.Panel1.Controls.Add(_previewCanvas);
            _mainSplit.Panel1.Controls.Add(leftHeader);

            // Right Panel: Options + Interactive DataGridView
            _dataGrid = new DataGridView();
            _dataGrid.Dock = DockStyle.Fill;
            _dataGrid.AllowUserToAddRows = false;
            _dataGrid.AllowUserToDeleteRows = false;
            _dataGrid.BackgroundColor = Color.White;
            _dataGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
            _dataGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            _dataGrid.CellValueChanged += OnGridCellValueChanged;
            _dataGrid.CurrentCellDirtyStateChanged += delegate {
                if (_dataGrid.IsCurrentCellDirty && _dataGrid.CurrentCell is DataGridViewCheckBoxCell)
                {
                    _dataGrid.CommitEdit(DataGridViewDataErrorContexts.Commit);
                }
            };

            _mainSplit.Panel2.Controls.Add(_dataGrid);
            _mainSplit.Panel2.Controls.Add(optionsPanel);

            this.Controls.Add(_mainSplit);
            this.Controls.Add(_topToolbar);
            this.Controls.Add(statusBar);

            // System Tray Icon so it can sit quietly waiting for print jobs from classic Win32 apps
            _trayIcon = new NotifyIcon();
            _trayIcon.Icon = SystemIcons.Application;
            _trayIcon.Text = "XLS Virtual Printer (Listening for Jobs)";
            _trayIcon.Visible = true;
            _trayIcon.DoubleClick += delegate {
                ShowAndActivate();
            };

            ContextMenu trayMenu = new ContextMenu();
            trayMenu.MenuItems.Add(new MenuItem("Open XLS Virtual Printer", delegate {
                ShowAndActivate();
            }));

            MenuItem autoStartItem = new MenuItem("Start with Windows");
            autoStartItem.Checked = IsAutoStartEnabled();
            autoStartItem.Click += delegate {
                bool newState = !autoStartItem.Checked;
                SetAutoStart(newState);
                autoStartItem.Checked = newState;
            };
            trayMenu.MenuItems.Add(autoStartItem);

            trayMenu.MenuItems.Add("-");
            trayMenu.MenuItems.Add(new MenuItem("Exit", delegate {
                _isExplicitExit = true;
                this.Close();
                Application.Exit();
            }));
            _trayIcon.ContextMenu = trayMenu;
        }

        private void UpdatePrinterInstallStatus()
        {
            bool installed = _printerManager.IsPrinterInstalled();
            if (installed)
            {
                _printerStateLabel.Text = "Printer Installed: '" + _printerManager.PrinterName + "' (Listening for Jobs)";
                _printerStateLabel.ForeColor = Color.DarkGreen;
            }
            else
            {
                _printerStateLabel.Text = "Printer Not Installed (Click 'Install Virtual Printer' or use Simulate/Open)";
                _printerStateLabel.ForeColor = Color.DarkRed;
            }
        }

        private void InstallOrRepairVirtualPrinter()
        {
            string msg;
            bool ok = _printerManager.InstallVirtualPrinter(out msg);
            UpdatePrinterInstallStatus();
            MessageBox.Show(this, msg, "Virtual Printer Setup", MessageBoxButtons.OK,
                ok ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
        }

        private void SendActualGdiTestPrintOrSimulate()
        {
            LoadBuiltInSamplePrintJob();
            _statusLabel.Text = "Simulated classic Win32 GDI print job with Arabic & English tabular data.";
        }

        private void LoadBuiltInSamplePrintJob()
        {
            byte[] emfBytes = SamplePrintGenerator.CreateSampleArabicAccountingEmf();
            _spoolParser.NormalizePresentationForms = _chkNormalizeArabicForms.Checked;
            _spoolParser.ReverseVisualArabic = _chkReverseVisualArabic.Checked;
            _currentDocument = _spoolParser.ParseBytes(emfBytes, "Arabic_Sales_Statement_Oct2026");
            _currentPageIndex = 0;
            _selectionRegionMm = null;
            RebuildTableFromDocument();
            _previewCanvas.Invalidate();
        }

        private void LoadSpoolOrEmfFile(string filePath)
        {
            _statusLabel.Text = "Loading print job: " + Path.GetFileName(filePath) + "...";
            this.Cursor = Cursors.WaitCursor;

            bool normForms = _chkNormalizeArabicForms.Checked;
            bool revArabic = _chkReverseVisualArabic.Checked;
            bool revCols = _chkReverseColumns.Checked;

            ThreadPool.QueueUserWorkItem(delegate
            {
                try
                {
                    EmfSpoolParser parser = new EmfSpoolParser();
                    parser.NormalizePresentationForms = normForms;
                    parser.ReverseVisualArabic = revArabic;

                    PrintJobDocument doc = parser.ParseFile(filePath);

                    TableLayoutEngine layout = new TableLayoutEngine();
                    layout.ReverseColumnsForRtl = revCols;
                    ReconstructedTable table = (doc.Pages.Count > 0)
                        ? layout.BuildTable(doc, 0, null)
                        : null;

                    this.BeginInvoke(new MethodInvoker(delegate
                    {
                        this.Cursor = Cursors.Default;
                        _currentDocument = doc;
                        _currentPageIndex = 0;
                        _selectionRegionMm = null;
                        _currentTable = table;

                        ApplyReconstructedTableToGrid();
                        _previewCanvas.Invalidate();
                        _statusLabel.Text = "Loaded print job from: " + filePath;
                    }));
                }
                catch (Exception ex)
                {
                    this.BeginInvoke(new MethodInvoker(delegate
                    {
                        this.Cursor = Cursors.Default;
                        MessageBox.Show(this, "Error parsing file:\n" + ex.Message, "Load Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }));
                }
            });
        }

        private void ReloadDocumentWithArabicSettings()
        {
            if (_currentDocument == null || _currentDocument.Pages.Count == 0) return;
            byte[] rawEmf = _currentDocument.Pages[0].RawEmfBytes;
            if (rawEmf != null && rawEmf.Length > 0)
            {
                _spoolParser.NormalizePresentationForms = _chkNormalizeArabicForms.Checked;
                _spoolParser.ReverseVisualArabic = _chkReverseVisualArabic.Checked;
                _currentDocument = _spoolParser.ParseBytes(rawEmf, _currentDocument.DocumentTitle);
                RebuildTableFromDocument();
                _previewCanvas.Invalidate();
            }
        }

        public void ShowAndActivate()
        {
            this.Show();
            if (this.WindowState == FormWindowState.Minimized)
                this.WindowState = FormWindowState.Normal;
            this.BringToFront();
            this.Activate();
        }

        private void OnPrintJobCaptured(object sender, PrintJobCapturedEventArgs e)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new EventHandler<PrintJobCapturedEventArgs>(OnPrintJobCaptured), sender, e);
                return;
            }

            AppLogger.Log(string.Format("MainForm: OnPrintJobCaptured received for Job #{0} '{1}' ({2} bytes)",
                e.JobId, e.DocumentTitle, e.SpoolBytes != null ? e.SpoolBytes.Length : 0));

            // Bring window to foreground even if minimized or hidden in system tray
            ShowAndActivate();

            _statusLabel.Text = string.Format("Processing captured job #{0}: '{1}'...", e.JobId, e.DocumentTitle);
            this.Cursor = Cursors.WaitCursor;

            bool normForms = _chkNormalizeArabicForms.Checked;
            bool revArabic = _chkReverseVisualArabic.Checked;
            bool revCols = _chkReverseColumns.Checked;

            // Process parsing & layout on background thread so the UI thread NEVER hangs or says (Not Responding)
            ThreadPool.QueueUserWorkItem(delegate
            {
                try
                {
                    AppLogger.Log("WorkerThread: Starting spool parsing...");
                    EmfSpoolParser parser = new EmfSpoolParser();
                    parser.NormalizePresentationForms = normForms;
                    parser.ReverseVisualArabic = revArabic;

                    PrintJobDocument doc = parser.ParseBytes(e.SpoolBytes, e.DocumentTitle);
                    AppLogger.Log(string.Format("WorkerThread: ParseBytes finished with {0} page(s)", doc.Pages.Count));

                    TableLayoutEngine layout = new TableLayoutEngine();
                    layout.ReverseColumnsForRtl = revCols;
                    ReconstructedTable table = (doc.Pages.Count > 0)
                        ? layout.BuildTable(doc, 0, null)
                        : null;
                    AppLogger.Log(string.Format("WorkerThread: BuildTable finished: {0} rows x {1} cols",
                        table != null ? table.Rows.Count : 0, table != null ? table.ColumnHeaders.Count : 0));

                    this.BeginInvoke(new MethodInvoker(delegate
                    {
                        try
                        {
                            AppLogger.Log("UI: Applying table and preview...");
                            this.Cursor = Cursors.Default;
                            _currentDocument = doc;
                            _currentPageIndex = 0;
                            _selectionRegionMm = null;
                            _currentTable = table;

                            ApplyReconstructedTableToGrid();
                            _previewCanvas.Invalidate();

                            _statusLabel.Text = string.Format("Captured Print Job #{0}: '{1}' ({2} pages, {3} rows)",
                                e.JobId, e.DocumentTitle, doc.Pages.Count, (table != null ? table.Rows.Count : 0));
                            AppLogger.Log("UI: Job capture successfully displayed in grid and canvas!");
                        }
                        catch (Exception uiex)
                        {
                            AppLogger.Log("UI Exception during job display: " + uiex.ToString());
                        }
                    }));
                }
                catch (Exception ex)
                {
                    AppLogger.Log("WorkerThread Exception: " + ex.ToString());
                    this.BeginInvoke(new MethodInvoker(delegate
                    {
                        this.Cursor = Cursors.Default;
                        _statusLabel.Text = "Error processing print job: " + ex.Message;
                    }));
                }
            });
        }

        private void RebuildTableFromDocument()
        {
            if (_currentDocument == null) return;

            _layoutEngine.ReverseColumnsForRtl = _chkReverseColumns.Checked;
            _currentTable = _layoutEngine.BuildTable(_currentDocument, 0, _selectionRegionMm);

            ApplyReconstructedTableToGrid();
        }

        private void ApplyReconstructedTableToGrid()
        {
            if (_currentTable == null) return;

            // Populate Columns CheckList
            _columnsCheckList.ItemCheck -= OnColumnItemCheck;
            _columnsCheckList.Items.Clear();
            for (int c = 0; c < _currentTable.ColumnHeaders.Count; c++)
            {
                string colLabel = "Col " + _currentTable.ColumnHeaders[c];
                _columnsCheckList.Items.Add(colLabel, _currentTable.SelectedColumns[c]);
            }
            _columnsCheckList.ItemCheck += OnColumnItemCheck;

            PopulateDataGrid();
        }

        private void PopulateDataGrid()
        {
            try
            {
                _dataGrid.SuspendLayout();
                _dataGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None; // Prevent per-row layout freeze
                _dataGrid.Columns.Clear();
                _dataGrid.Rows.Clear();

                if (_currentTable == null || _currentTable.ColumnHeaders == null || _currentTable.ColumnHeaders.Count == 0)
                {
                    return;
                }

                DataGridViewCheckBoxColumn chkCol = new DataGridViewCheckBoxColumn();
                chkCol.Name = "_ExportRow";
                chkCol.HeaderText = "Export";
                chkCol.Width = 50;
                _dataGrid.Columns.Add(chkCol);

                for (int c = 0; c < _currentTable.ColumnHeaders.Count; c++)
                {
                    DataGridViewTextBoxColumn txtCol = new DataGridViewTextBoxColumn();
                    txtCol.Name = "Col_" + c;
                    txtCol.HeaderText = _currentTable.ColumnHeaders[c];
                    txtCol.Visible = (c < _currentTable.SelectedColumns.Count) ? _currentTable.SelectedColumns[c] : true;
                    _dataGrid.Columns.Add(txtCol);
                }

                Font boldFont = null;
                try { boldFont = new Font(_dataGrid.Font, FontStyle.Bold); } catch { }

                // Build all rows in memory first and add in batch with Rows.AddRange for maximum speed
                DataGridViewRow[] gridRows = new DataGridViewRow[_currentTable.Rows.Count];
                for (int r = 0; r < _currentTable.Rows.Count; r++)
                {
                    GridRowData rowData = _currentTable.Rows[r];
                    DataGridViewRow row = new DataGridViewRow();
                    row.CreateCells(_dataGrid);
                    row.Cells[0].Value = rowData.Selected;

                    bool anyBold = false;
                    for (int c = 0; c < _currentTable.ColumnHeaders.Count; c++)
                    {
                        if (c < rowData.Cells.Length)
                        {
                            row.Cells[1 + c].Value = rowData.Cells[c].Text;
                            if (rowData.Cells[c].IsBold) anyBold = true;
                        }
                        else
                        {
                            row.Cells[1 + c].Value = "";
                        }
                    }
                    if (anyBold && boldFont != null)
                    {
                        row.DefaultCellStyle.Font = boldFont;
                    }
                    gridRows[r] = row;
                }

                _dataGrid.Rows.AddRange(gridRows);
                try
                {
                    _dataGrid.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
                }
                catch { }

                _statusLabel.Text = string.Format("Extracted {0} rows x {1} columns. Edit cells or uncheck rows/columns before exporting to .XLS.",
                    _currentTable.Rows.Count, _currentTable.ColumnHeaders.Count);
            }
            catch (Exception ex)
            {
                AppLogger.Log("PopulateDataGrid Exception: " + ex.ToString());
            }
            finally
            {
                _dataGrid.ResumeLayout();
            }
        }

        private void OnColumnItemCheck(object sender, ItemCheckEventArgs e)
        {
            if (_currentTable == null || e.Index < 0 || e.Index >= _currentTable.SelectedColumns.Count)
                return;

            bool isChecked = (e.NewValue == CheckState.Checked);
            _currentTable.SelectedColumns[e.Index] = isChecked;

            int gridColIdx = e.Index + 1;
            if (gridColIdx < _dataGrid.Columns.Count)
            {
                _dataGrid.Columns[gridColIdx].Visible = isChecked;
            }
        }

        private void OnGridCellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (_currentTable == null || e.RowIndex < 0 || e.RowIndex >= _currentTable.Rows.Count)
                return;

            GridRowData rowData = _currentTable.Rows[e.RowIndex];
            if (e.ColumnIndex == 0)
            {
                object val = _dataGrid.Rows[e.RowIndex].Cells[0].Value;
                rowData.Selected = (val is bool) && (bool)val;
            }
            else if (e.ColumnIndex >= 1)
            {
                int colIdx = e.ColumnIndex - 1;
                if (colIdx < rowData.Cells.Length)
                {
                    object val = _dataGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
                    rowData.Cells[colIdx].Text = (val != null) ? val.ToString() : string.Empty;
                }
            }
        }

        private void SetAllRowsSelected(bool selected)
        {
            if (_currentTable == null) return;
            for (int r = 0; r < _currentTable.Rows.Count; r++)
            {
                _currentTable.Rows[r].Selected = selected;
                if (r < _dataGrid.Rows.Count)
                    _dataGrid.Rows[r].Cells[0].Value = selected;
            }
        }

        // ================= Visual Page Preview & Rubberband Region Selection =================

        private RectangleF GetPageScreenRect(ExtractedPage page)
        {
            float pad = 16f;
            float availW = Math.Max(50f, _previewCanvas.ClientSize.Width - pad * 2);
            float availH = Math.Max(50f, _previewCanvas.ClientSize.Height - pad * 2);

            float pageW = Math.Max(10f, page.Width);
            float pageH = Math.Max(10f, page.Height);

            float scale = Math.Min(availW / pageW, availH / pageH);
            float drawW = pageW * scale;
            float drawH = pageH * scale;
            float drawX = (_previewCanvas.ClientSize.Width - drawW) / 2f;
            float drawY = (_previewCanvas.ClientSize.Height - drawH) / 2f;

            return new RectangleF(drawX, drawY, drawW, drawH);
        }

        private void OnPreviewCanvasPaint(object sender, PaintEventArgs e)
        {
            try
            {
                if (_previewCanvas == null || _previewCanvas.ClientSize.Width <= 10 || _previewCanvas.ClientSize.Height <= 10)
                    return;

                Graphics g = e.Graphics;
                g.SmoothingMode = SmoothingMode.HighQuality;
                g.Clear(Color.FromArgb(95, 102, 112));

                if (_currentDocument == null || _currentDocument.Pages.Count == 0)
                    return;

                if (_currentPageIndex < 0 || _currentPageIndex >= _currentDocument.Pages.Count)
                    return;

                ExtractedPage page = _currentDocument.Pages[_currentPageIndex];
                if (page == null) return;
                RectangleF pageRect = GetPageScreenRect(page);

                // Drop shadow + White Paper
                using (SolidBrush shadowBrush = new SolidBrush(Color.Black))
                using (SolidBrush paperBrush = new SolidBrush(Color.White))
                {
                    g.FillRectangle(shadowBrush, pageRect.X + 4, pageRect.Y + 4, pageRect.Width, pageRect.Height);
                    g.FillRectangle(paperBrush, pageRect.X, pageRect.Y, pageRect.Width, pageRect.Height);
                }

                float scaleX = pageRect.Width / Math.Max(1f, page.Width);
                float scaleY = pageRect.Height / Math.Max(1f, page.Height);

            // Render actual EMF metafile if available
            if (page.RawEmfBytes != null && page.RawEmfBytes.Length > 0)
            {
                try
                {
                    using (MemoryStream ms = new MemoryStream(page.RawEmfBytes))
                    using (Metafile mf = new Metafile(ms))
                    {
                        g.DrawImage(mf, pageRect);
                    }
                }
                catch { }
            }

            // If Arabic / RTL is present or reverse visual Arabic is enabled,
            // the raw EMF contains glyphs emitted by legacy drivers in inverted visual order.
            // Overlay the clean, normalized, correctly shaped text from page.Elements
            // so the visual preview displays proper Right-to-Left Arabic text.
            if (_chkReverseVisualArabic.Checked || page.Elements.Exists(delegate(ExtractedTextElement eItem) { return eItem.IsRtl || ArabicTextHelper.ContainsArabic(eItem.Text); }))
            {
                using (SolidBrush textCoverBrush = new SolidBrush(Color.White))
                using (SolidBrush textBrush = new SolidBrush(Color.FromArgb(20, 20, 20)))
                using (StringFormat rtlFmt = new StringFormat(StringFormatFlags.DirectionRightToLeft | StringFormatFlags.NoWrap))
                {
                    rtlFmt.Alignment = StringAlignment.Far;
                    rtlFmt.LineAlignment = StringAlignment.Center;

                    for (int i = 0; i < page.Elements.Count; i++)
                    {
                        ExtractedTextElement el = page.Elements[i];
                        if (el == null || string.IsNullOrEmpty(el.Text)) continue;

                        bool isArabic = el.IsRtl || ArabicTextHelper.ContainsArabic(el.Text);
                        if (isArabic)
                        {
                            try
                            {
                                float sx = pageRect.X + el.X * scaleX;
                                float sy = pageRect.Y + el.Y * scaleY;
                                float sw = Math.Max(4f, el.Width * scaleX);
                                float sh = Math.Max(4f, el.Height * scaleY);

                                if (float.IsNaN(sx) || float.IsInfinity(sx) || float.IsNaN(sy) || float.IsInfinity(sy) ||
                                    float.IsNaN(sw) || float.IsInfinity(sw) || float.IsNaN(sh) || float.IsInfinity(sh) ||
                                    sw <= 0f || sh <= 0f)
                                    continue;

                                RectangleF textRect = new RectangleF(sx - 1.5f, sy - 1f, sw + 3f, sh + 2f);

                                // Blank out inverted raw EMF glyphs underneath
                                g.FillRectangle(textCoverBrush, textRect);

                                // Draw clean, properly shaped and oriented Arabic text
                                float pixelSize = Math.Max(8f, Math.Min(48f, sh * 0.85f));
                                FontStyle fs = el.IsBold ? FontStyle.Bold : FontStyle.Regular;

                                Rectangle cellRect = new Rectangle(
                                    (int)Math.Floor(textRect.X),
                                    (int)Math.Floor(textRect.Y),
                                    (int)Math.Ceiling(textRect.Width),
                                    (int)Math.Ceiling(textRect.Height));

                                using (Font font = new Font("Tahoma", pixelSize, fs, GraphicsUnit.Pixel))
                                {
                                    // Prefer native Win32 GDI TextRenderer (DrawTextExW) which handles RTL/Arabic reliably
                                    // and avoids the known GDI+ DrawString ExternalException on Windows 10/11
                                    try
                                    {
                                        TextFormatFlags flags = TextFormatFlags.RightToLeft | TextFormatFlags.Right |
                                                                TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine |
                                                                TextFormatFlags.NoClipping | TextFormatFlags.NoPrefix;
                                        TextRenderer.DrawText(g, el.Text, font, cellRect, Color.FromArgb(20, 20, 20), flags);
                                    }
                                    catch
                                    {
                                        // Fallback to GDI+ DrawString with guarded exception handling
                                        try
                                        {
                                            g.DrawString(el.Text, font, textBrush, textRect, rtlFmt);
                                        }
                                        catch { }
                                    }
                                }
                            }
                            catch
                            {
                                // Safeguard: Individual element drawing error must never crash canvas paint
                            }
                        }
                    }
                }
            }

            // Draw bounding boxes around extracted text elements
            using (Pen activeBoxPen = new Pen(Color.FromArgb(140, 0, 120, 215), 1f))
            using (Pen excludedBoxPen = new Pen(Color.FromArgb(90, 180, 180, 180), 1f))
            {
                for (int i = 0; i < page.Elements.Count; i++)
                {
                    ExtractedTextElement el = page.Elements[i];
                    if (el == null) continue;
                    RectangleF elMm = new RectangleF(el.X, el.Y, Math.Max(2f, el.Width), Math.Max(2f, el.Height));
                    bool insideFilter = !_selectionRegionMm.HasValue || _selectionRegionMm.Value.IntersectsWith(elMm);

                    float sx = pageRect.X + el.X * scaleX;
                    float sy = pageRect.Y + el.Y * scaleY;
                    float sw = Math.Max(4f, el.Width * scaleX);
                    float sh = Math.Max(4f, el.Height * scaleY);

                    if (!float.IsNaN(sx) && !float.IsNaN(sy) && !float.IsNaN(sw) && !float.IsNaN(sh))
                    {
                        g.DrawRectangle(insideFilter ? activeBoxPen : excludedBoxPen, sx, sy, sw, sh);
                    }
                }
            }

            // Draw active region crop overlay (in mm)
            if (_selectionRegionMm.HasValue)
            {
                RectangleF rMm = _selectionRegionMm.Value;
                RectangleF rScreen = new RectangleF(
                    pageRect.X + rMm.X * scaleX,
                    pageRect.Y + rMm.Y * scaleY,
                    rMm.Width * scaleX,
                    rMm.Height * scaleY);

                using (Brush highlightBrush = new SolidBrush(Color.FromArgb(40, 0, 160, 80)))
                using (Pen borderPen = new Pen(Color.FromArgb(220, 0, 140, 60), 2f))
                {
                    borderPen.DashStyle = DashStyle.Dash;
                    g.FillRectangle(highlightBrush, rScreen);
                    g.DrawRectangle(borderPen, rScreen.X, rScreen.Y, rScreen.Width, rScreen.Height);
                }
            }

            // Draw live mouse drag rubberband rectangle
            if (_isDraggingSelection)
            {
                Rectangle dragRect = GetNormalizedDragRect(_dragStartScreenPt, _dragCurrentScreenPt);
                using (Brush dragBrush = new SolidBrush(Color.FromArgb(50, 0, 120, 215)))
                using (Pen dragPen = new Pen(Color.FromArgb(220, 0, 100, 200), 1.5f))
                {
                    g.FillRectangle(dragBrush, dragRect);
                    g.DrawRectangle(dragPen, dragRect);
                }
            }
            }
            catch
            {
                // Never allow a GDI / paint glitch to crash the WinForms message loop
            }
        }

        private void OnPreviewMouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left || _currentDocument == null || _currentDocument.Pages.Count == 0)
                return;

            _isDraggingSelection = true;
            _dragStartScreenPt = e.Location;
            _dragCurrentScreenPt = e.Location;
        }

        private void OnPreviewMouseMove(object sender, MouseEventArgs e)
        {
            if (!_isDraggingSelection) return;
            _dragCurrentScreenPt = e.Location;
            _previewCanvas.Invalidate();
        }

        private void OnPreviewMouseUp(object sender, MouseEventArgs e)
        {
            if (!_isDraggingSelection) return;
            _isDraggingSelection = false;

            Rectangle dragScreen = GetNormalizedDragRect(_dragStartScreenPt, e.Location);
            if (dragScreen.Width < 8 || dragScreen.Height < 8)
            {
                _previewCanvas.Invalidate();
                return;
            }

            ExtractedPage page = _currentDocument.Pages[_currentPageIndex];
            RectangleF pageRect = GetPageScreenRect(page);

            float scaleX = pageRect.Width / Math.Max(1f, page.Width);
            float scaleY = pageRect.Height / Math.Max(1f, page.Height);

            float mmX = Math.Max(0f, (dragScreen.X - pageRect.X) / scaleX);
            float mmY = Math.Max(0f, (dragScreen.Y - pageRect.Y) / scaleY);
            float mmW = Math.Min(page.Width - mmX, dragScreen.Width / scaleX);
            float mmH = Math.Min(page.Height - mmY, dragScreen.Height / scaleY);

            if (mmW > 2f && mmH > 2f)
            {
                _selectionRegionMm = new RectangleF(mmX, mmY, mmW, mmH);
                _btnClearRegion.Enabled = true;
                _lblSelectionInfo.Text = string.Format("Area Filter: ({0:F0},{1:F0}) {2:F0}x{3:F0} mm", mmX, mmY, mmW, mmH);
                RebuildTableFromDocument();
            }

            _previewCanvas.Invalidate();
        }

        private static Rectangle GetNormalizedDragRect(Point a, Point b)
        {
            int x = Math.Min(a.X, b.X);
            int y = Math.Min(a.Y, b.Y);
            int w = Math.Abs(a.X - b.X);
            int h = Math.Abs(a.Y - b.Y);
            return new Rectangle(x, y, w, h);
        }

        // ================= File Open & Excel Export =================

        private void PromptOpenSpoolFile()
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Title = "Open Windows Spool (.SPL), Enhanced Metafile (.EMF), or Text Print (.PRN)";
                ofd.Filter = "Print Spool & Metafile (*.spl;*.emf;*.prn;*.txt)|*.spl;*.emf;*.prn;*.txt|All Files (*.*)|*.*";
                if (ofd.ShowDialog(this) == DialogResult.OK)
                {
                    LoadSpoolOrEmfFile(ofd.FileName);
                }
            }
        }

        private void ExportCurrentTable(string defaultExt)
        {
            if (_currentTable == null || _currentTable.Rows.Count == 0)
            {
                MessageBox.Show(this, "No table rows available to export.", "Export", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Title = "Save Selected Print Data to Excel";
                sfd.Filter = "Excel 97-2003 Workbook (*.xls)|*.xls|Excel OpenXML Workbook (*.xlsx)|*.xlsx|UTF-8 CSV (*.csv)|*.csv";
                if (defaultExt == ".xlsx") sfd.FilterIndex = 2;
                else if (defaultExt == ".csv") sfd.FilterIndex = 3;
                else sfd.FilterIndex = 1;

                string baseName = (_currentDocument != null && !string.IsNullOrEmpty(_currentDocument.DocumentTitle))
                    ? SanitizeFileName(_currentDocument.DocumentTitle)
                    : "PrintOutput";
                sfd.FileName = baseName + defaultExt;

                if (sfd.ShowDialog(this) == DialogResult.OK)
                {
                    ExportTableToPath(_currentTable, sfd.FileName, _chkRtlSheet.Checked);
                    _statusLabel.Text = "Saved Excel file: " + sfd.FileName;
                }
            }
        }

        public static void ExportTableToPath(ReconstructedTable table, string filePath, bool rtlSheet)
        {
            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            if (ext == ".xlsx")
            {
                XlsxAndCsvWriter writer = new XlsxAndCsvWriter();
                writer.RightToLeftSheet = rtlSheet;
                writer.SaveToXlsx(table, filePath);
            }
            else if (ext == ".csv")
            {
                XlsxAndCsvWriter writer = new XlsxAndCsvWriter();
                writer.RightToLeftSheet = rtlSheet;
                writer.SaveToCsv(table, filePath);
            }
            else
            {
                Biff8XlsWriter writer = new Biff8XlsWriter();
                writer.RightToLeftSheet = rtlSheet;
                writer.SaveToXls(table, filePath);
            }
        }

        private static string SanitizeFileName(string name)
        {
            char[] invalid = Path.GetInvalidFileNameChars();
            for (int i = 0; i < invalid.Length; i++)
                name = name.Replace(invalid[i], '_');
            return name;
        }

        private const string RunRegistryKey = @"Software\Microsoft\Windows\CurrentVersion\Run";
        private const string RunValueName = "XlsVirtualPrinter";

        public static bool IsAutoStartEnabled()
        {
            try
            {
                using (Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(RunRegistryKey, false))
                {
                    if (key != null)
                    {
                        object val = key.GetValue(RunValueName);
                        return val != null;
                    }
                }
            }
            catch { }
            return false;
        }

        public static void SetAutoStart(bool enable)
        {
            try
            {
                using (Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(RunRegistryKey, true))
                {
                    if (key != null)
                    {
                        if (enable)
                        {
                            string exePath = Application.ExecutablePath;
                            key.SetValue(RunValueName, "\"" + exePath + "\" --tray");
                        }
                        else
                        {
                            key.DeleteValue(RunValueName, false);
                        }
                    }
                }
            }
            catch { }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing && !_isExplicitExit)
            {
                // Minimize to System Tray instead of terminating so background print watcher stays active
                e.Cancel = true;
                this.Hide();
                if (_trayIcon != null)
                {
                    _trayIcon.ShowBalloonTip(1500, "XLS Virtual Printer",
                        "Still listening in system tray. Print jobs will automatically pop up here.",
                        ToolTipIcon.Info);
                }
                return;
            }

            if (_trayIcon != null)
            {
                _trayIcon.Visible = false;
                _trayIcon.Dispose();
            }
            if (_printerManager != null)
            {
                _printerManager.Dispose();
            }
            base.OnFormClosing(e);
        }
    }
}
