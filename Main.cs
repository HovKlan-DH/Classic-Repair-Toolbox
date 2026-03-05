using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Commodore_Repair_Toolbox
{
    public partial class Main : Form
    {
        // Default values
        private static string buildType = ""; // Debug|Release
        private string onlineAvailableVersion = ""; // will be empty, if no newer version available
        private static string crtPage = "https://commodore-repair-toolbox.dk";
        private static string crtNewExecutable = crtPage + "/auto-update/Commodore-Repair-Toolbox.exe";
        private static string crtPageAutoUpdate = "/auto-update/";
        private static string crtPageFeedback = "/feedback-app/";

        private static bool _logInitialized = false;
        private static readonly object _logSync = new object();
        private static bool _syncPerformedAtLaunch = false;

        // For "Resources" tab
        private Panel _resourcesPanel;
        private ListView _resourcesLocalFiles;
        private ListView _resourcesWebLinks;
        private const int ResourcesNameColumnIndex = 1;
        private ListViewItem _hoverItem;
        private int _hoverSubItemIndex = -1;
        private Font _resourcesLinkFont;
        private Font _resourcesNormalFont;
        private static readonly Color _resourcesLocalHeaderBackColor = Color.LightGray;
        private static readonly Color _resourcesLinksHeaderBackColor = Color.LightGray;
        private Rectangle _hoverItemBounds = Rectangle.Empty;
        private bool _suppressResourcesSelection;

        // For "Help" tab
        private const int EM_SETRECT = 0x00B3;

        public static void InitializeLogging()
        {
            if (_logInitialized) return;
            lock (_logSync)
            {
                if (_logInitialized) return;
                CreateDebugOutputFile();
                _logInitialized = true;
            }
        }

        // Reference to the popup/info form
        private FormComponent componentInfoPopup = null;

        // Blinking of components
        private Timer blinkTimer;
        private bool blinkState = false;

        // Fullscreen mode
        private string windowState = "Maximized";
        private bool isFullscreen = false;
        private FormWindowState formPreviousWindowState;
        private FormBorderStyle formPreviousFormBorderStyle;
        private Rectangle previousBoundsForm;
        private Rectangle previousBoundsPanelBehindTab;
        private Rectangle previousBoundsFullscreenButton;

        // Main panel (left side) + image
        private CustomPanel panelZoom;
        private Panel panelImageMain;
        private Image image;

        // "Main" schematics (left-side of SplitContainer)
        private Label labelFile;
        private Label labelComponent;
        private Label labelRegion;
        private Label labelZoom;
        public static OverlayPanel overlayPanel;
        private Dictionary<string, Control> overlayLabelMap = new Dictionary<string, Control>();

        // Thumbnails (right-side of SplitContainer)
        private Dictionary<string, OverlayPanel> overlayPanelsList = new Dictionary<string, OverlayPanel>();
        private Dictionary<string, float> overlayListZoomFactors = new Dictionary<string, float>();
        private Color labelImageBgClr = Color.Khaki;
        private Color labelImageTxtClr = Color.Black;
        private Color labelImageHasElementsBgClr = Color.SteelBlue;
        private Color labelImageHasElementsTxtClr = Color.White;

        // Resizing of window/schematic
        private bool isResizedByMouseWheel = false;

        // Data loaded from Excel
        public static List<Hardware> classHardware = new List<Hardware>();
        public static string[] versionMismatch = {};

        // Overlays for main image
        private List<PictureBox> overlayComponentsTab = new List<PictureBox>();
        private Dictionary<int, Point> overlayComponentsTabOriginalLocations = new Dictionary<int, Point>();
        private Dictionary<int, Size> overlayComponentsTabOriginalSizes = new Dictionary<int, Size>();

        // Selected entries in the components list
        private List<string> listBoxComponentsSelectedText = new List<string>();

        // Version information
        public static string versionThis = "";
        private static string versionOnline = "";
        private static string versionOnlineTxt = "";

        // Current user selection
        public static string hardwareSelectedName;
        public static string boardSelectedName;
        private string boardSelectedFilename;
        public static string schematicSelectedName;
        private string schematicSelectedFile;
        public static float zoomFactor = 1.0f;
        private Point overlayPanelLastMousePos = Point.Empty;

        // "Shadow" lists for the classes - neded in "InitializeTabConfiguration()" function
        public static Dictionary<string, Dictionary<string, List<string>>> shadow_structure;

        // Misc
        private Dictionary<Control, EventHandler> clickEventHandlers = new Dictionary<Control, EventHandler>();
        private TabPage previousTab;
        private int thumbnailSelectedBorderWidth = 3;
        private Timer windowMoveStopTimer = new Timer();
        private Point windowLastLocation;
        bool thumbnailsSameWidth = false;
        int thumbnailsWidth = 0;
        public static string selectedRegion = ""; // "", "PAL" or "NTSC"

        // Polyline
        private PolylinesManagement polylinesManagement;
        private int zoomLevel = 1;
        private int panelTracesVisibleHeight = 0;
        private bool isPanelTracesVisible = false;


        // ###########################################################################################
        // Main form constructor.
        // ###########################################################################################

        public Main()
        {
            InitializeComponent();

            // Load configuration parameter "UpdateDataAtNextLaunch"
            string syncDataAtNextLaunchStr = Configuration.GetSetting("UpdateDataAtNextLaunch", "False");
            bool shouldSyncData = bool.TryParse(syncDataAtNextLaunchStr, out bool result) && result;
            if (shouldSyncData)
            {
                SyncFilesFromSource();
            }

            // Initialize "Always update data" checkbox from configuration (wire event after to avoid an extra save)
            this.checkBoxAlwaysUpdateData.Checked = Configuration.GetSetting("AlwaysUpdateData", "False") == "True";
            this.checkBoxAlwaysUpdateData.CheckedChanged += this.checkBoxAlwaysUpdateData_CheckedChanged;

            polylinesManagement = new PolylinesManagement(this);

            // Get build type
#if DEBUG
            buildType = "Debug";
#else
                buildType = "Release";
#endif

            // Enable double-buffering for smoother UI rendering
            EnableDoubleBuffering();

            // Get application versions - both this one and the one online
            versionThis = GetAssemblyVersion();
            GetOnlineVersion();

            // Load all files (Excel and configuration)
            LoadExcelData();

            Splashscreen.Current?.UpdateStatus("Finishing setting up the rest");

            InitializeTabConfiguration();
            InitializeTabOverviewOverviewGrid();
            InitializeTabHelp();

            // Attach "form load" event, which is triggered just before form is shown
            Load += Form_Load;

        }


        // ###########################################################################################
        // Event: Form load.
        // Triggered just before form is shown
        // ###########################################################################################

        private void Form_Load(object sender, EventArgs e)
        {
            windowState = Configuration.GetSetting("WindowState", "Maximized");
            if (Enum.TryParse(windowState, out FormWindowState state) && state != FormWindowState.Minimized)
            {
                this.WindowState = state;
            }

            tabControl.Dock = DockStyle.Fill;
            panelLabelsVisible.Visible = false;

            // Initialize the blink timer
            InitializeBlinkTimer();

            Shown += Form_Shown;
        }

        private async void Form_Shown(object sender, EventArgs e)
        {
            Debug.WriteLine("Form_Shown()");

            ApplyConfigSettings();

            // Attach various event handles
            AttachEventHandlers();

            AttachConfigurationSaveEvents();

            SetupNewBoard();

            UpdateComponentList("Form_Shown");

            SetRegionButtonColors();

            // Set initial focus to "textBoxFilterComponents"
            textBoxFilterComponents.Focus();

            // Initialize zoom level
            labelZoom.Text = $"Z:{zoomLevel}";

            StartDrawingPolylines();
            PopulatePolylineVisibilityPanel();

            // Set a "cross" cursor to visualize "drawing mode" when inside the overlay panel
            overlayPanel.Cursor = Cursors.Cross;

            // Show popup, if "versionMismatch" is not empy, that we have a version mismatch between one or more Excel data files nad the application version
            if (versionMismatch.Length > 0)
            {
                string message = "There is a version compatibility mismatch between one or more of the Excel data files and this application version. Please update the application to the newest available version, and afterwards update the data from online source.\r\n\r\n" +
                    "The following Excel data files do not match the application version:\r\n\r\n";
                foreach (string file in versionMismatch)
                {
                    message += "* " + file + "\r\n";
                }
                message += "\r\nThis may lead to unexpected behavior or errors in the application.";
                MessageBox.Show(
                    message,
                    "WARNING: Application and data version mismatch",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
            }

            // Wait 10 seconds before starting the background check.
            // Skip the background check if a full sync was already performed at launch.
            label13.TextAlign = ContentAlignment.MiddleCenter;
            await Task.Delay(10000);
            if (!_syncPerformedAtLaunch)
            {
                await Task.Run(() => checkFilesFromSource());
            }
        }

        private void Form_Closing(object sender, FormClosingEventArgs e)
        {
            PolylinesManagement.SavePolylinesToConfig();
        }


        // ###########################################################################################
        // Get the assembly version.
        // Will transform assembly information into a text string.
        // ###########################################################################################

        public static string GetAssemblyVersion()
        {
            try
            {
                Assembly assemblyInfo = Assembly.GetExecutingAssembly();
                string assemblyVersion = FileVersionInfo.GetVersionInfo(assemblyInfo.Location).FileVersion;
                string year = assemblyVersion.Substring(0, 4);
                string month = assemblyVersion.Substring(5, 2);
                string day = assemblyVersion.Substring(8, 2);
                string rev = assemblyVersion.Substring(11); // will be ignored in RELEASE builds
                switch (month)
                {
                    case "01": month = "January"; break;
                    case "02": month = "February"; break;
                    case "03": month = "March"; break;
                    case "04": month = "April"; break;
                    case "05": month = "May"; break;
                    case "06": month = "June"; break;
                    case "07": month = "July"; break;
                    case "08": month = "August"; break;
                    case "09": month = "September"; break;
                    case "10": month = "October"; break;
                    case "11": month = "November"; break;
                    case "12": month = "December"; break;
                    default: month = "Unknown"; break;
                }
                day = day.TrimStart(new Char[] { '0' }); // remove leading zero
                day = day.TrimEnd(new Char[] { '.' }); // remove last dot
                string date = year + "-" + month + "-" + day;

                // Beautify revision and build-type 
                rev = "(rev. " + rev + ")";
                rev = buildType == "Debug" ? rev : "";
                string buildTypeTmp = buildType == "Debug" ? "# DEVELOPMENT " : "";

                // Set the application version
                return (date + " " + buildTypeTmp + rev).Trim();
            }
            catch (Exception ex)
            {
                DebugOutput("EXCEPTION in \"GetAssemblyVersion()\":");
                DebugOutput(ex.ToString());
                return null;
            }
        }


        // ###########################################################################################
        // Get the version available online (newest version).
        // ###########################################################################################

        private void GetOnlineVersion()
        {
            try
            {
                using (WebClient webClient = new WebClient())
                {
                    ServicePointManager.Expect100Continue = true;
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls13;
                    webClient.Headers.Add("user-agent", "CRT " + versionThis);

                    // Include some control POST data
                    var postData = new System.Collections.Specialized.NameValueCollection
                    {
                        { "control", "CRT" }
                    };

                    // Send the POST data to the server
                    byte[] responseBytes = webClient.UploadValues(crtPage + crtPageAutoUpdate, postData);

                    // Convert the "response bytes" to a human readable string
                    onlineAvailableVersion = Encoding.UTF8.GetString(responseBytes);

                    if (onlineAvailableVersion.Substring(0, 7) == "Version")
                    {
                        onlineAvailableVersion = onlineAvailableVersion.Substring(9);
                        if (onlineAvailableVersion != versionThis)
                        {
                            tabAbout.Text = "About*";
                            versionOnline = onlineAvailableVersion;

                            // Add a simple auto-update button docked to the top of the About tab
                            Panel aboutNotifyPanel = new Panel
                            {
                                Dock = DockStyle.Top,
                                Height = 140,
                                BackColor = Color.IndianRed,
                                Name = "panelAutoUpdate",
                            };

                            // Add banner text that states where to manually update new version
                            aboutNotifyPanel.Controls.Add(
                                new Label
                                {
                                    Text = "\"Commodore Repair Toolbox\" is now considered as \"legacy\" and a new and better product has taken over.\nNo more updates will be given to this, including no more data updates.\n\nPlease go to https://github.com/HovKlan-DH/Classic-Repair-Toolbox and download the new...\n\n\"Classic Repair Toolbox\".",
                                    AutoSize = false,
                                    TextAlign = ContentAlignment.MiddleCenter,
                                    Dock = DockStyle.Fill,
                                    Font = new Font("Calibri", 12F, FontStyle.Bold),
                                    ForeColor = Color.White
                                }
                            );


                            /*
                            Button autoUpdateButton = new Button
                            {
                                Text = "Auto-install New Version" + Environment.NewLine + versionOnline,
                                AutoSize = false,
                                Size = new Size(220, 50),
                                Location = new Point(11, 11),
                                Anchor = AnchorStyles.Top | AnchorStyles.Left,
                                BackColor = Color.IndianRed,
                                ForeColor = Color.White,
                                FlatStyle = FlatStyle.Flat
                            };
                            autoUpdateButton.FlatAppearance.BorderColor = Color.DarkRed;
                            autoUpdateButton.FlatAppearance.BorderSize = 2;
                            autoUpdateButton.Click += (sender, e) =>
                            {
                                AutoUpdater.PerformUpdate(crtNewExecutable, "--fetch-data");
                            };
                            */

                            //aboutNotifyPanel.Controls.Add(autoUpdateButton);

                            // Draw a solid IndianRed border around the notification panel
                            aboutNotifyPanel.Paint += (s, pe) =>
                            {
                                float penWidth = thumbnailSelectedBorderWidth;
                                using (Pen pen = new Pen(Color.IndianRed, penWidth))
                                {
                                    float offset = penWidth / 2;
                                    pe.Graphics.DrawRectangle(
                                        pen,
                                        offset,
                                        offset,
                                        aboutNotifyPanel.ClientSize.Width - penWidth,
                                        aboutNotifyPanel.ClientSize.Height - penWidth
                                    );
                                }
                            };

                            // Add to About tab (stays at back of Z-order so it docks above richTextBoxAbout)
                            tabAbout.Controls.Add(aboutNotifyPanel);

                            // Top-bar notification label (text only; the button lives in the About tab)
                            string newText = "Newer version is available; view \"About\" tab";
                            Size textSize = TextRenderer.MeasureText(newText, label13.Font);
                            label13.Text = newText;
                            label13.Width = textSize.Width + 2;
                            label13.Anchor = AnchorStyles.Top | AnchorStyles.Right;
                            label13.Visible = true;
                            label13.Location = new Point(panelBehindTab.Width - label13.Width - 2, 3);
                            label13.BringToFront();
                        }
                        else
                        {
                            versionOnline = "";
                        }
                    }
                    else
                    {
                        tabAbout.Text = "About*";
                        versionOnlineTxt = "<font color='IndianRed'>";
                        versionOnlineTxt += "<hr>";
                        versionOnlineTxt += "ERROR:<br />";
                        versionOnlineTxt += $"It was not possible to check for a newer version, as the server connection to <a href='https://commodore-repair-toolbox.dk' target='_blank'>https://commodore-repair-toolbox.dk</a> cannot be established - please check your connectivity.<br />The exact recieved HTTP error is:<br /><br /><b>{onlineAvailableVersion}</b>";
                        versionOnlineTxt += "<hr>";
                        versionOnlineTxt += "<br />";
                        versionOnlineTxt += "</font>";
                    }
                }
            }
            catch (Exception ex)
            {
                DebugOutput("EXCEPTION in \"GetOnlineVersion()\":");
                DebugOutput(ex.ToString());
                tabAbout.Text = "About*";
                versionOnlineTxt = "<font color='IndianRed'>";
                versionOnlineTxt += "<hr>";
                versionOnlineTxt += "ERROR:<br />";
                versionOnlineTxt += $"It was not possible to check for a newer version, as the server connection to <a href='https://commodore-repair-toolbox.dk' target='_blank'>https://commodore-repair-toolbox.dk</a> cannot be established - please check your connectivity.<br />The exact recieved HTTP error is:<br /><br /><b>{ex.Message}</b>";
                versionOnlineTxt += "<hr>";
                versionOnlineTxt += "<br />";
                versionOnlineTxt += "</font>";
            }
        }


        // ###########################################################################################
        // Load all Excel data, and have it ready for usage.
        // ###########################################################################################

        private void LoadExcelData()
        {
            DataStructure.GetAllData(classHardware);
        }


        // ###########################################################################################
        // Apply saved configuration settings to controls.
        // ###########################################################################################

        private void ApplyConfigSettings()
        {
            // Set default values first
            string defaultSelectedHardware = classHardware.FirstOrDefault()?.Name;
            string defaultSelectedBoard = classHardware.FirstOrDefault(h => h.Name == defaultSelectedHardware)?.Boards?.FirstOrDefault()?.Name;
            string defaultShowLabel = "True";
            string defaultShowTechnicalName = "False";
            string defaultShowFriendlyName = "False";
            string defaultShowLabelsHeight = "95";
            string defaultShowTraces = "True";
            string defaultShowTracesHeight = "0";
            string defaultKeyboardZoomEnabled = "0";

            // Load saved settings from configuration file - or set default if none exists
            string selectedHardwareVal = Configuration.GetSetting("HardwareSelected", defaultSelectedHardware);
            string selectedBoardVal = Configuration.GetSetting("BoardSelected", defaultSelectedBoard);
            string selectedShowLabel = Configuration.GetSetting("ShowLabel", defaultShowLabel);
            string selectedShowTechnicalName = Configuration.GetSetting("ShowTechnicalName", defaultShowTechnicalName);
            string selectedShowFriendlyName = Configuration.GetSetting("ShowFriendlyName", defaultShowFriendlyName);
            string showLabelsHeight = Configuration.GetSetting("ShowLabelsHeight", defaultShowLabelsHeight);
            string showTraces = Configuration.GetSetting("ShowTraces", defaultShowTraces);
            string showTracesHeight = Configuration.GetSetting("ShowTracesHeight", defaultShowTracesHeight);
            string userEmail = Configuration.GetSetting("UserEmail", "");
            selectedRegion = Configuration.GetSetting("SelectedRegion", "PAL");
            bool keyboardZoomEnabled = Configuration.GetSetting("KeyboardZoomEnabled", defaultKeyboardZoomEnabled) == "1" ? true : false;

            textBoxEmail.Text = userEmail; // set email address in "Feedback" tab

            // Populate all hardware in combobox - and select
            foreach (Hardware hardware in classHardware)
            {
                comboBoxHardware.Items.Add(hardware.Name);
            }
            int indexHardware = comboBoxHardware.Items.IndexOf(selectedHardwareVal);
            if (indexHardware == -1)
            {
                indexHardware = comboBoxHardware.Items.IndexOf(defaultSelectedHardware);
                hardwareSelectedName = defaultSelectedHardware;
            }
            comboBoxHardware.SelectedIndex = indexHardware;
            hardwareSelectedName = comboBoxHardware.SelectedItem.ToString();
            textBox5.Text = ConvertStringToLabel(hardwareSelectedName); // feedback info           

            // Populate all boards in combobox, based on selected hardware - and select
            var hw = classHardware.FirstOrDefault(h => h.Name == hardwareSelectedName);
            if (hw != null)
            {
                foreach (var board in hw.Boards)
                {
                    comboBoxBoard.Items.Add(board.Name);
                }
                int indexBoard = comboBoxBoard.Items.IndexOf(selectedBoardVal);
                if (indexBoard == -1)
                {
                    indexBoard = comboBoxBoard.Items.IndexOf(defaultSelectedBoard);
                    boardSelectedName = defaultSelectedBoard;
                }
                comboBoxBoard.SelectedIndex = indexBoard;
                boardSelectedName = comboBoxBoard.SelectedItem.ToString();
                textBox1.Text = ConvertStringToLabel(boardSelectedName); // feedback info           
            }

            // Set the "Labels visible" checkboxes
            checkBox1.Checked = selectedShowLabel == "True" ? true : false;
            checkBox2.Checked = selectedShowTechnicalName == "True" ? true : false;
            checkBox3.Checked = selectedShowFriendlyName == "True" ? true : false;

            // .. and the height for its panel
            panelLabelsVisible.Height = Convert.ToInt32(showLabelsHeight);

            panelTracesVisibleHeight = Convert.ToInt32(showTracesHeight);
            isPanelTracesVisible = bool.TryParse(showTraces, out bool result) && result;

            checkBoxKeyboardZoom.Checked = keyboardZoomEnabled;
        }


        // ###########################################################################################
        // Create/overwrite the debug output logfile
        // ###########################################################################################

        private static void CreateDebugOutputFile()
        {
            string filePath = Path.Combine(Application.StartupPath, "Commodore-Repair-Toolbox.log");
            using (StreamWriter writer = new StreamWriter(filePath, false))
            {
                writer.WriteLine("Debug output file created [" + DateTime.Now + "]");
            }
        }

        public static void DebugOutput(string text)
        {
            // Ensure logging is initialized even if this is called early
            if (!_logInitialized)
                InitializeLogging();

            string filePath = Path.Combine(Application.StartupPath, "Commodore-Repair-Toolbox.log");
            try
            {
                lock (_logSync)
                {
                    using (StreamWriter writer = new StreamWriter(filePath, true))
                    {
                        writer.WriteLine(text);
                    }
                }
            }
            catch { /* swallow logging errors to avoid impacting app */ }

            Debug.WriteLine(text);
        }


        // ###########################################################################################
        // Attach necessary event handlers.
        // ###########################################################################################

        private void AttachEventHandlers()
        {
            ResizeBegin += Form_ResizeBegin;
            ResizeEnd += Form_ResizeEnd;
            comboBoxHardware.SelectedIndexChanged += comboBoxHardware_SelectedIndexChanged;
            comboBoxBoard.SelectedIndexChanged += comboBoxBoard_SelectedIndexChanged;
            checkBoxBlink.CheckedChanged += CheckBoxBlink_CheckedChanged;
            textBoxFilterComponents.TextChanged += TextBoxFilterComponents_TextChanged;
            splitContainerSchematics.Paint += SplitContainer1_Paint;
            comboBoxHardware.DropDownClosed += ComboBox_DropDownClosed;
            comboBoxBoard.DropDownClosed += ComboBox_DropDownClosed;
            listBoxCategories.SelectedIndexChanged += ListBoxCategories_SelectedIndexChanged;
            AttachClickEventsToFocusFilterComponents(this);
            textBoxEmail.TextChanged += TextBoxEmail_TextChanged;
            tabControl.SelectedIndexChanged += TabControl_SelectedIndexChanged;
            splitContainerSchematics.SplitterMoved += SplitContainer1_SplitterMoved;
            listBoxComponents.SelectedIndexChanged += listBoxComponents_SelectedIndexChanged;
            buttonRegionPal.Click += ButtonRegionPal_Click;
            buttonRegionNtsc.Click += ButtonRegionNtsc_Click;
            windowMoveStopTimer.Interval = 200;
            windowMoveStopTimer.Tick += MoveStopTimer_Tick;
            Move += Form_Move;
            FormClosing += Form_Closing;
        }


        // ###########################################################################################
        // Attach event handlers for saving configuration settings.
        // ###########################################################################################

        private void AttachConfigurationSaveEvents()
        {
            // Selection of new hardware
            comboBoxHardware.SelectedIndexChanged += (s, e) =>
            {
                Configuration.SaveSetting("HardwareSelected", hardwareSelectedName);
            };

            // Selection of new board
            comboBoxBoard.SelectedIndexChanged += (s, e) =>
            {
                Configuration.SaveSetting("BoardSelected", boardSelectedName);
            };
        }


        // ###########################################################################################
        // Attach a click event to all controls with a name, and have them focus
        // "textBoxFilterComponents" so it is ready for text input for filtering.
        // ###########################################################################################

        private void AttachClickEventsToFocusFilterComponents(Control parent)
        {
            if (!(parent is ComboBox))
            {
                EventHandler handler = (s, e) => textBoxFilterComponents.Focus();
                parent.Click += handler;
                clickEventHandlers[parent] = handler;
            }
            foreach (Control child in parent.Controls)
            {
                if (child != null)
                {
                    AttachClickEventsToFocusFilterComponents(child);
                }
            }
        }

        private void RemoveClickEventsToFocusFilterComponents()
        {
            foreach (var kvp in clickEventHandlers)
            {
                kvp.Key.Click -= kvp.Value;
            }
            clickEventHandlers.Clear();
        }


        // ###########################################################################################
        // Events: Resize.
        // ###########################################################################################

        private void Form_ResizeBegin(object sender, EventArgs e)
        {
            SuspendLayout();
        }

        private void Form_ResizeEnd(object sender, EventArgs e)
        {
            ResizeTabImage();
            ResumeLayout();
            ReadaptThumbnails();
        }

        private void Form_Move(object sender, EventArgs e)
        {
            // Disable event for "splitter moved", as it otherwise ruins the saved data
            splitContainerSchematics.SplitterMoved -= SplitContainer1_SplitterMoved;

            // Enable/reset the time to detect when the movemen has stopped
            windowLastLocation = this.Location;
            windowMoveStopTimer.Stop();
            windowMoveStopTimer.Start();
        }


        // ###########################################################################################
        // What happens when window movement has stopped
        // ###########################################################################################

        private void MoveStopTimer_Tick(object sender, EventArgs e)
        {
            if (this.Location == windowLastLocation)
            {
                windowMoveStopTimer.Stop();

                // Save the new "window state" and load and apply the splitter position for the new state
                windowState = this.WindowState.ToString();
                Configuration.SaveSetting("WindowState", windowState);
                LoadAndApplySplitterPosition();

                ReadaptThumbnails();

                // We can now reenable the event for "splitter moved"
                splitContainerSchematics.SplitterMoved += SplitContainer1_SplitterMoved;
            }
        }


        // ###########################################################################################
        // Overview tab
        // ###########################################################################################

        private void InitializeTabOverviewOverviewGrid()
        {
            // Hide the left row-header column entirely
            dataGridViewOverview.RowHeadersVisible = false;

            dataGridViewOverview.DefaultCellStyle.Font = new Font("Calibri", 11F);
            dataGridViewOverview.ColumnHeadersDefaultCellStyle.Font = new Font("Calibri", 11F, FontStyle.Bold);
            dataGridViewOverview.EnableHeadersVisualStyles = false;
            dataGridViewOverview.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0xDD, 0xDD, 0xDD);
            dataGridViewOverview.ColumnHeadersDefaultCellStyle.ForeColor = Color.Black;

            dataGridViewOverview.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridViewOverview.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells;
            dataGridViewOverview.RowTemplate.Resizable = DataGridViewTriState.False;

            // Prevent any visible selection/highlight (WITHOUT breaking LinkColumn clicks)
            dataGridViewOverview.MultiSelect = false;
            dataGridViewOverview.SelectionMode = DataGridViewSelectionMode.CellSelect;
            dataGridViewOverview.DefaultCellStyle.SelectionBackColor = dataGridViewOverview.DefaultCellStyle.BackColor;
            dataGridViewOverview.DefaultCellStyle.SelectionForeColor = dataGridViewOverview.DefaultCellStyle.ForeColor;

            // Clear selection immediately after it changes, but keep CurrentCell intact
            dataGridViewOverview.SelectionChanged -= DataGridViewOverview_SelectionChanged_Clear;
            dataGridViewOverview.SelectionChanged += DataGridViewOverview_SelectionChanged_Clear;

            dataGridViewOverview.Columns.Clear();

            dataGridViewOverview.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Type",
                HeaderText = "Type",
                DataPropertyName = "Type",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            });

            dataGridViewOverview.Columns.Add(new DataGridViewLinkColumn
            {
                Name = "Component",
                HeaderText = "Component",
                DataPropertyName = "ComponentLabel",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells,
                LinkColor = Color.FromArgb(0x06, 0x45, 0xAD),
                ActiveLinkColor = Color.FromArgb(0x0B, 0x00, 0x80),
                VisitedLinkColor = Color.FromArgb(0x06, 0x45, 0xAD)
            });

            dataGridViewOverview.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "TechnicalName",
                HeaderText = "Technical name",
                DataPropertyName = "TechnicalName",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells,
                DefaultCellStyle = { WrapMode = DataGridViewTriState.True }
            });

            dataGridViewOverview.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "FriendlyName",
                HeaderText = "Friendly name",
                DataPropertyName = "FriendlyName",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells,
                DefaultCellStyle = { WrapMode = DataGridViewTriState.True }
            });

            dataGridViewOverview.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "PartNumber",
                HeaderText = "Part-number",
                DataPropertyName = "PartNumber",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            });

            dataGridViewOverview.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "ShortDescription",
                HeaderText = "Short description",
                DataPropertyName = "ShortDescription",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                MinimumWidth = 150,
                DefaultCellStyle = { WrapMode = DataGridViewTriState.True }
            });

            dataGridViewOverview.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Notes",
                HeaderText = "Notes",
                DataPropertyName = "Notes",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                MinimumWidth = 200,
                DefaultCellStyle = { WrapMode = DataGridViewTriState.True }
            });

            dataGridViewOverview.Columns.Add(new DataGridViewLinkColumn
            {
                Name = "LocalFiles",
                HeaderText = "Local files",
                DataPropertyName = "LocalFilesText",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            });

            dataGridViewOverview.Columns.Add(new DataGridViewLinkColumn
            {
                Name = "WebLinks",
                HeaderText = "Web links",
                DataPropertyName = "WebLinksText",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            });

            dataGridViewOverview.CellContentClick -= DataGridViewOverview_CellContentClick;
            dataGridViewOverview.CellContentClick += DataGridViewOverview_CellContentClick;
        }

        private void DataGridViewOverview_SelectionChanged_Clear(object sender, EventArgs e)
        {
            if (!dataGridViewOverview.IsHandleCreated) return;

            // Keep link-click functionality: do NOT set CurrentCell = null here
            dataGridViewOverview.ClearSelection();
        }

        private void UpdateTabOverviewNew(Board selectedBoard)
        {
            if (selectedBoard == null)
            {
                dataGridViewOverview.DataSource = null;
                return;
            }

            var visibleDisplays = listBoxComponents.Items.Cast<string>().ToHashSet();

            var rows = new List<OverviewRow>();

            foreach (var comp in selectedBoard.Components ?? Enumerable.Empty<BoardComponents>())
            {
                if (!visibleDisplays.Contains(comp.NameDisplay)) continue;

                string compLabel = comp.Label;

                string baseKey = $"UserData|{hardwareSelectedName}|{boardSelectedName}|{compLabel}";
                string oneLinerKey = $"{baseKey}|Oneliner";
                string notesKey = $"{baseKey}|Notes";

                string shortDescr = Configuration.GetSetting(oneLinerKey, "");
                if (string.IsNullOrEmpty(shortDescr))
                {
                    shortDescr = comp.OneLiner;
                }

                string notes = Configuration.GetSetting(notesKey, "");
                if (string.IsNullOrEmpty(notes))
                {
                    if (comp.ComponentImages != null && comp.ComponentImages.Count > 0)
                    {
                        var firstNote = comp.ComponentImages
                            .Select(img => img.Note)
                            .FirstOrDefault(note => !string.IsNullOrEmpty(note));
                        notes = firstNote ?? "";
                    }
                    else
                    {
                        notes = "";
                    }
                }

                var localFilePaths = new List<string>();
                if (comp.LocalFiles != null)
                {
                    foreach (var f in comp.LocalFiles)
                    {
                        string filePath = Path.GetFullPath(DataPaths.Resolve(f.FileName));
                        localFilePaths.Add(filePath);
                    }
                }

                var webUrls = new List<string>();
                if (comp.ComponentLinks != null)
                {
                    foreach (var link in comp.ComponentLinks)
                    {
                        if (!string.IsNullOrWhiteSpace(link.Url))
                        {
                            webUrls.Add(link.Url);
                        }
                    }
                }

                rows.Add(new OverviewRow
                {
                    Type = comp.Type,
                    ComponentLabel = compLabel,
                    TechnicalName = comp.NameTechnical,
                    FriendlyName = (comp.NameFriendly ?? "").Replace("?", ""),
                    PartNumber = comp.Partnumber,
                    ShortDescription = shortDescr,
                    Notes = notes,

                    LocalFilesText = localFilePaths.Count > 0 ? $"#{localFilePaths.Count}" : "",
                    WebLinksText = webUrls.Count > 0 ? $"#{webUrls.Count}" : "",

                    LocalFilePaths = localFilePaths,
                    WebUrls = webUrls
                });
            }

            dataGridViewOverview.DataSource = rows;
        }

        private sealed class OverviewRow
        {
            public string Type { get; set; }
            public string ComponentLabel { get; set; }
            public string TechnicalName { get; set; }
            public string FriendlyName { get; set; }
            public string PartNumber { get; set; }
            public string ShortDescription { get; set; }
            public string Notes { get; set; }

            // Display text in grid
            public string LocalFilesText { get; set; }
            public string WebLinksText { get; set; }

            // Targets used when clicking
            public List<string> LocalFilePaths { get; set; }
            public List<string> WebUrls { get; set; }
        }                

        private void DataGridViewOverview_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            var row = dataGridViewOverview.Rows[e.RowIndex];
            var model = row.DataBoundItem as OverviewRow;
            if (model == null) return;

            string col = dataGridViewOverview.Columns[e.ColumnIndex].Name;

            if (col == "Component")
            {
                var foundHardware = classHardware.FirstOrDefault(h => h.Name == hardwareSelectedName);
                var foundBoard = foundHardware?.Boards.FirstOrDefault(b => b.Name == boardSelectedName);
                var comps = foundBoard?.Components.Where(c => c.Label == model.ComponentLabel).ToList();

                if (comps != null && comps.Count > 0)
                {
                    ShowComponentPopup(comps);
                }
                return;
            }

            if (col == "LocalFiles" && model.LocalFilePaths != null && model.LocalFilePaths.Count > 0)
            {
                string filePath = model.LocalFilePaths[0]; // simple: open first; can be enhanced to choose
                if (File.Exists(filePath))
                {
                    OpenExternalTarget(filePath, isUrl: false);
                }
                else
                {
                    DebugOutput("File [" + filePath + "] does not exists!");
                }
                return;
            }

            if (col == "WebLinks" && model.WebUrls != null && model.WebUrls.Count > 0)
            {
                OpenExternalTarget(model.WebUrls[0], isUrl: true);
                return;
            }
        }


        // ###########################################################################################
        // Initialize and update the tab for "Overview".
        // Will load new content from board data file.
        // ###########################################################################################     

        public async void UpdateTabNewOverview(Board selectedBoard)
        {
            UpdateTabOverviewNew(selectedBoard);
        }

        private void CloseComponentPopup()
        {
            if (componentInfoPopup != null && !componentInfoPopup.IsDisposed)
            {
                componentInfoPopup.Close();
                componentInfoPopup = null;
            }
        }


        // ###########################################################################################
        // Initialize and update the tab for "Resources".
        // Will load new content from board data file.
        // ###########################################################################################

        private void InitializeTabResourcesNew()
        {
            if (_resourcesPanel != null)
            {
                return;
            }

            _resourcesPanel = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                BackColor = Color.White
            };

            var header = new Label
            {
                Text = "Resources for troubleshooting and information",
                Font = new Font("Calibri", 14, FontStyle.Bold),
                AutoSize = true,
                Location = new Point(12, 12)
            };

            var localHeader = new Label
            {
                Text = "Local files",
                Font = new Font("Calibri", 12, FontStyle.Bold),
                AutoSize = true,
                Location = new Point(12, header.Bottom + 14)
            };

            _resourcesNormalFont = new Font("Calibri", 11F, FontStyle.Regular);
            _resourcesLinkFont = new Font(_resourcesNormalFont, FontStyle.Underline);

            _resourcesLocalFiles = new ListView
            {
                View = View.Details,
                FullRowSelect = false,
                HideSelection = true,
                Location = new Point(12, localHeader.Bottom + 6),
                Width = tabResources.ClientSize.Width - 24,
                Height = 220,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                HeaderStyle = ColumnHeaderStyle.Nonclickable,
                Font = _resourcesNormalFont,
                BorderStyle = BorderStyle.None
            };
            _resourcesLocalFiles.Columns.Add("Category", 200);
            _resourcesLocalFiles.Columns.Add("Name", 700);
            _resourcesLocalFiles.HotTracking = false;
            _resourcesLocalFiles.HoverSelection = false;
            _resourcesLocalFiles.DoubleBuffered(true);

            var linksHeader = new Label
            {
                Text = "Links",
                Font = new Font("Calibri", 12, FontStyle.Bold),
                AutoSize = true,
                Location = new Point(12, _resourcesLocalFiles.Bottom + 18)
            };

            _resourcesWebLinks = new ListView
            {
                View = View.Details,
                FullRowSelect = false,
                HideSelection = true,
                Location = new Point(12, linksHeader.Bottom + 6),
                Width = tabResources.ClientSize.Width - 24,
                Height = 220,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                HeaderStyle = ColumnHeaderStyle.Nonclickable,
                Font = _resourcesNormalFont,
                BorderStyle = BorderStyle.None
            };
            _resourcesWebLinks.Columns.Add("Category", 200);
            _resourcesWebLinks.Columns.Add("Name", 700);
            _resourcesWebLinks.HotTracking = false;
            _resourcesWebLinks.HoverSelection = false;
            _resourcesWebLinks.DoubleBuffered(true);

            ApplyResourcesListViewLinkBehavior(_resourcesLocalFiles);
            ApplyResourcesListViewLinkBehavior(_resourcesWebLinks);

            tabResources.Resize += (s, e) =>
            {
                LayoutResourcesControls();
            };

            _resourcesPanel.Controls.Add(header);
            _resourcesPanel.Controls.Add(localHeader);
            _resourcesPanel.Controls.Add(_resourcesLocalFiles);
            _resourcesPanel.Controls.Add(linksHeader);
            _resourcesPanel.Controls.Add(_resourcesWebLinks);

            tabResources.Controls.Add(_resourcesPanel);

            ApplyResourcesListViewHeaderStyle(_resourcesLocalFiles, _resourcesLocalHeaderBackColor);
            ApplyResourcesListViewHeaderStyle(_resourcesWebLinks, _resourcesLinksHeaderBackColor);
        }

        private void ApplyResourcesListViewLinkBehavior(ListView lv)
        {
            lv.MouseDown -= ResourcesListView_MouseDown;
            lv.MouseDown += ResourcesListView_MouseDown;

            lv.MouseUp -= ResourcesListView_MouseUp;
            lv.MouseUp += ResourcesListView_MouseUp;

            lv.MouseMove -= ResourcesListView_MouseMove;
            lv.MouseMove += ResourcesListView_MouseMove;

            lv.MouseLeave -= ResourcesListView_MouseLeave;
            lv.MouseLeave += ResourcesListView_MouseLeave;

            // Important: don't use ItemActivate when using MouseUp
            lv.ItemActivate -= ResourcesListView_ItemActivate;
        }

        private void ResourcesListView_ItemActivate(object sender, EventArgs e)
        {
            var lv = (ListView)sender;

            // Must have a focused/activated item
            var item = lv.FocusedItem;
            if (item == null)
            {
                return;
            }

            string target = item.Tag as string;
            if (string.IsNullOrWhiteSpace(target))
            {
                DebugOutput("[Resources] Target is empty/null");
                return;
            }

            ResourcesOpenTarget(target, lv == _resourcesLocalFiles);

            // Prevent the row from staying selected
            lv.SelectedIndices.Clear();
            lv.FocusedItem = null;

            FocusFilterTextBox();
        }

        private void ResourcesListView_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            if (!_suppressResourcesSelection) return;

            if (e.IsSelected)
            {
                e.Item.Selected = false;
                e.Item.Focused = false;

                var lv = (ListView)sender;
                lv.Invalidate(e.Item.Bounds);
            }
        }

        private void FocusFilterTextBox()
        {
            if (tabControl.SelectedTab == tabSchematics)
            {
                // When on Schematics, focus should stay on the filter textbox by design
                textBoxFilterComponents.Focus();
                return;
            }

            // Also allow focus behavior on other tabs if you want it globally:
            textBoxFilterComponents.Focus();
        }

        private void ResourcesListView_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) return;

            FocusFilterTextBox();

            var lv = (ListView)sender;
            var hit = lv.HitTest(e.Location);

            int subIndex = hit.Item != null && hit.SubItem != null
                ? hit.Item.SubItems.IndexOf(hit.SubItem)
                : -1;

            _suppressResourcesSelection = (hit.Item != null && subIndex == ResourcesNameColumnIndex);

            if (_suppressResourcesSelection)
            {
                lv.SelectedIndices.Clear();
                lv.FocusedItem = null;
            }
        }

        private static void SetListViewColumnWidths(ListView lv, int categoryWidth)
        {
            if (lv == null) return;
            if (lv.Columns.Count < 2) return;

            // Keep first column fixed
            lv.Columns[0].Width = categoryWidth;

            // Remaining space goes to column 1 ("Name")
            int scrollbarWidth = lv.Items.Count > 0 ? SystemInformation.VerticalScrollBarWidth : 0;

            // Leave a bit of padding for borders/owner-draw artifacts
            int padding = 6;

            int available = lv.ClientSize.Width - categoryWidth - scrollbarWidth - padding;
            lv.Columns[1].Width = Math.Max(50, available);
        }


        private void ResourcesListView_MouseMove(object sender, MouseEventArgs e)
        {
            var lv = (ListView)sender;
            var hit = lv.HitTest(e.Location);

            int subIndex = hit.Item != null && hit.SubItem != null
                ? hit.Item.SubItems.IndexOf(hit.SubItem)
                : -1;

            lv.Cursor = (hit.Item != null && subIndex == ResourcesNameColumnIndex)
                ? Cursors.Hand
                : Cursors.Default;

            // Only repaint rows when the hovered item changes (avoid full Invalidate flicker)
            if (!ReferenceEquals(_hoverItem, hit.Item) || _hoverSubItemIndex != subIndex)
            {
                // repaint previous hovered row
                if (!_hoverItemBounds.IsEmpty)
                {
                    lv.Invalidate(_hoverItemBounds);
                }

                _hoverItem = hit.Item;
                _hoverSubItemIndex = subIndex;

                // repaint new hovered row
                _hoverItemBounds = hit.Item != null ? hit.Item.Bounds : Rectangle.Empty;
                if (!_hoverItemBounds.IsEmpty)
                {
                    lv.Invalidate(_hoverItemBounds);
                }
            }
        }

        private void ResourcesListView_MouseLeave(object sender, EventArgs e)
        {
            var lv = (ListView)sender;
            lv.Cursor = Cursors.Default;

            if (!_hoverItemBounds.IsEmpty)
            {
                lv.Invalidate(_hoverItemBounds);
            }

            _hoverItem = null;
            _hoverSubItemIndex = -1;
            _hoverItemBounds = Rectangle.Empty;
        }

        private void ResourcesListView_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) return;

            DebugOutput("[Resources] MouseUp fired");

            FocusFilterTextBox();

            _suppressResourcesSelection = false;

            var lv = (ListView)sender;
            var hit = lv.HitTest(e.Location);

            int subIndex = hit.Item != null && hit.SubItem != null
                ? hit.Item.SubItems.IndexOf(hit.SubItem)
                : -1;

            if (hit.Item == null || subIndex != ResourcesNameColumnIndex)
            {
                return;
            }

            // Prevent the row from staying selected when clicking the "Name" link column
            if (hit.Item.Selected)
            {
                hit.Item.Selected = false;
                if (hit.Item.Focused)
                {
                    hit.Item.Focused = false;
                }

                lv.HideSelection = true;
                lv.Invalidate(hit.Item.Bounds);
            }

            string target = hit.Item.Tag as string;
            if (string.IsNullOrWhiteSpace(target))
            {
                DebugOutput("[Resources] Target is empty/null");
                return;
            }

            ResourcesOpenTarget(target, lv == _resourcesLocalFiles);
        }

        private void ResourcesOpenTarget(string target, bool isLocalFile)
        {
            if (string.IsNullOrWhiteSpace(target))
            {
                DebugOutput("[Resources] Target is empty/null");
                return;
            }

            target = target.Trim();
            DebugOutput("[Resources] Open: isLocalFile=" + isLocalFile + " target=[" + target + "]");

            OpenExternalTarget(target, isUrl: !isLocalFile);

            textBoxFilterComponents.Focus();
        }

        private void ApplyResourcesListViewHeaderStyle(ListView lv, Color headerBackColor)
        {
            lv.OwnerDraw = true;

            lv.DrawColumnHeader -= ResourcesListView_DrawColumnHeader;
            lv.DrawColumnHeader += ResourcesListView_DrawColumnHeader;

            lv.DrawSubItem -= ResourcesListView_DrawSubItem;
            lv.DrawSubItem += ResourcesListView_DrawSubItem;

            lv.Tag = headerBackColor; // store per-list header color
        }

        private void ResourcesListView_DrawColumnHeader(object sender, DrawListViewColumnHeaderEventArgs e)
        {
            var lv = (ListView)sender;
            Color backColor = lv.Tag is Color c ? c : SystemColors.Control;

            using (var backBrush = new SolidBrush(backColor))
            using (var textBrush = new SolidBrush(Color.Black))
            using (var borderPen = new Pen(Color.FromArgb(0xAA, 0xAA, 0xAA)))
            {
                e.Graphics.FillRectangle(backBrush, e.Bounds);
                e.Graphics.DrawRectangle(borderPen, new Rectangle(e.Bounds.X, e.Bounds.Y, e.Bounds.Width - 1, e.Bounds.Height - 1));

                TextRenderer.DrawText(
                    e.Graphics,
                    e.Header.Text,
                    e.Font,
                    e.Bounds,
                    Color.Black,
                    TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
            }
        }

        private void ResourcesListView_DrawSubItem(object sender, DrawListViewSubItemEventArgs e)
        {
            var lv = (ListView)sender;

            // Detect "hot" (hovered) item so we can paint a stable background
            Point clientPos = lv.PointToClient(Cursor.Position);
            var hit = lv.HitTest(clientPos);
            bool isHotItem = hit != null && hit.Item == e.Item;

            // IMPORTANT: do not render selection at all (prevents "row marked" effect)
            Color backColor = isHotItem
                ? Color.FromArgb(0xF2, 0xF2, 0xF2)
                : (e.SubItem.BackColor.IsEmpty ? lv.BackColor : e.SubItem.BackColor);

            Color foreColor = e.SubItem.ForeColor.IsEmpty ? lv.ForeColor : e.SubItem.ForeColor;

            using (var backBrush = new SolidBrush(backColor))
            {
                e.Graphics.FillRectangle(backBrush, e.Bounds);
            }

            Font font = e.SubItem.Font ?? e.Item.Font;

            TextRenderer.DrawText(
                e.Graphics,
                e.SubItem.Text,
                font,
                e.Bounds,
                foreColor,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
        }

        private void UpdateTabResourcesNew(Board selectedBoard)
        {
            InitializeTabResourcesNew();

            _resourcesLocalFiles.BeginUpdate();
            _resourcesWebLinks.BeginUpdate();
            try
            {
                _resourcesLocalFiles.Items.Clear();
                _resourcesWebLinks.Items.Clear();

                if (selectedBoard == null)
                {
                    return;
                }

                if (selectedBoard.BoardLocalFiles != null)
                {
                    foreach (var f in selectedBoard.BoardLocalFiles)
                    {
                        string filePath = Path.GetFullPath(Path.Combine(DataPaths.DataRoot, f.Datafile));

                        var item = new ListViewItem(f.Category ?? "")
                        {
                            Tag = filePath
                        };
                        item.UseItemStyleForSubItems = false;

                        var nameSub = item.SubItems.Add(f.Name ?? "");
                        nameSub.ForeColor = Color.FromArgb(0x06, 0x45, 0xAD);
                        nameSub.Font = _resourcesLinkFont;

                        _resourcesLocalFiles.Items.Add(item);
                    }
                }

                if (selectedBoard.BoardLinks != null)
                {
                    foreach (var l in selectedBoard.BoardLinks)
                    {
                        var item = new ListViewItem(l.Category ?? "")
                        {
                            Tag = l.Url
                        };
                        item.UseItemStyleForSubItems = false;

                        var nameSub = item.SubItems.Add(l.Name ?? "");
                        nameSub.ForeColor = Color.FromArgb(0x06, 0x45, 0xAD);
                        nameSub.Font = _resourcesLinkFont;

                        _resourcesWebLinks.Items.Add(item);
                    }
                }

                LayoutResourcesControls();
            }
            finally
            {
                _resourcesLocalFiles.EndUpdate();
                _resourcesWebLinks.EndUpdate();
            }
        }

        private void LayoutResourcesControls()
        {
            if (_resourcesPanel == null || _resourcesLocalFiles == null || _resourcesWebLinks == null)
            {
                return;
            }

            var header = _resourcesPanel.Controls.OfType<Label>()
                .FirstOrDefault(l => l.Text == "Resources for troubleshooting and information");

            var localHeader = _resourcesPanel.Controls.OfType<Label>()
                .FirstOrDefault(l => l.Text == "Local files");

            var linksHeader = _resourcesPanel.Controls.OfType<Label>()
                .FirstOrDefault(l => l.Text == "Links");

            if (header == null || localHeader == null || linksHeader == null)
            {
                return;
            }

            int width = tabResources.ClientSize.Width - 24;
            _resourcesLocalFiles.Width = width;
            _resourcesWebLinks.Width = width;

            SetListViewColumnWidths(_resourcesLocalFiles, 200);
            SetListViewColumnWidths(_resourcesWebLinks, 200);

            _resourcesLocalFiles.Height = CalculateListViewHeight(_resourcesLocalFiles);
            _resourcesWebLinks.Height = CalculateListViewHeight(_resourcesWebLinks);

            linksHeader.Location = new Point(12, _resourcesLocalFiles.Bottom + 18);
            _resourcesWebLinks.Location = new Point(12, linksHeader.Bottom + 6);
        }

        private static int CalculateListViewHeight(ListView lv)
        {
            const int minHeight = 60;
            const int maxHeight = 600;

            // Height = header + rows + a little padding for borders
            int headerHeight = lv.Font.Height + 10;
            int rowHeight = lv.Font.Height + 6;
            int rows = Math.Max(1, lv.Items.Count);

            int desired = headerHeight + (rowHeight * rows) + 6;
            if (desired < minHeight) return minHeight;
            if (desired > maxHeight) return maxHeight;
            return desired;
        }


        // ###########################################################################################
        // Initialize the tab for "Help".
        // ###########################################################################################

        private void InitializeTabConfiguration()
        {

            // Create a panel per hardware
            int x = button2.Location.X; // start X-position
            int y = textBox3.Location.Y + textBox3.Height + 5; // start Y-position
            int spacing = 15; // space between panels           

            foreach (var hardware in shadow_structure)
            {
                string hardwareName = hardware.Key;
                var boardsDict = hardware.Value;

                Panel hardwarePanel = new Panel
                {
                    Name = $"panel_{hardwareName.Replace(" ", "_")}",
                    BorderStyle = BorderStyle.FixedSingle,
                    BackColor = Color.WhiteSmoke,
                    Size = new Size(500, 100),
                    Location = new Point(x, y),
                    AutoScroll = true
                };

                Label label = new Label
                {
                    Text = hardwareName,
                    Dock = DockStyle.Top,
                    Font = new Font("Calibri", 12, FontStyle.Bold)
                };
                hardwarePanel.Controls.Add(label);

                int boardY = label.Height + 5;

                foreach (var board in boardsDict)
                {
                    string boardName = board.Key;
                    List<string> schematicNames = board.Value;

                    // Unique name for board checkbox
                    string boardCheckBoxName = $"{hardwareName}|{boardName}";

                    // Get configuration setting
                    string boardConfigKey = $"ConfigurationCheckBoxState|{boardCheckBoxName}";
                    bool boardChecked = Configuration.GetSetting(boardConfigKey, "True") == "True";

                    CheckBox boardCheckBox = new CheckBox
                    {
                        Name = boardCheckBoxName,
                        Text = boardName,
                        Font = new Font("Calibri", 10, FontStyle.Bold),
                        Location = new Point(10, boardY),
                        AutoSize = true,
                        Checked = boardChecked
                    };
                    boardCheckBox.CheckedChanged += BoardOrSchematicCheckBox_CheckedChanged;
                    hardwarePanel.Controls.Add(boardCheckBox);
                    boardY += boardCheckBox.Height + 2;

                    // Add a checkbox for each schematic
                    foreach (var schematicName in schematicNames)
                    {
                        // Unique name for schematic checkbox
                        string schematicCheckBoxName = $"{hardwareName}|{boardName}|{schematicName}";

                        // Get configuration setting
                        string schematicConfigKey = $"ConfigurationCheckBoxState|{schematicCheckBoxName}";
                        bool schematicChecked = Configuration.GetSetting(schematicConfigKey, "True") == "True";

                        CheckBox schematicCheckBox = new CheckBox
                        {
                            Name = schematicCheckBoxName,
                            Text = schematicName,
                            Location = new Point(30, boardY),
                            AutoSize = true,
                            Checked = schematicChecked,
                            Enabled = boardCheckBox.Checked
                        };
                        schematicCheckBox.CheckedChanged += BoardOrSchematicCheckBox_CheckedChanged;
                        hardwarePanel.Controls.Add(schematicCheckBox);
                        boardY += schematicCheckBox.Height + 2;
                    }

                    // Add extra spacing after all schematics for a board
                    boardY += 10;
                }

                // Adjust panel height to fit all checkboxes
                hardwarePanel.Height = Math.Max(100, boardY + 10);

                // Add the panel to a parent container
                panel2.Controls.Add(hardwarePanel);

                y += hardwarePanel.Height + spacing;
            }

            // Insert a "dummy label" at the bottom of the panel, to have some space
            Label label2 = new Label
            {
                Text = "",
                Location = new Point(0, y - 25)
            };
            panel2.Controls.Add(label2);
        }


        // ###########################################################################################
        // Event handler for checkbox changes in the "Configuration" tab.
        // ###########################################################################################

        private void BoardOrSchematicCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            var checkBox = sender as CheckBox;
            if (checkBox == null) return;

            // Save the state
            string configKey = $"ConfigurationCheckBoxState|{checkBox.Name}";
            Configuration.SaveSetting(configKey, checkBox.Checked.ToString());

            // Split the name to determine if it is a board or schematic checkbox
            var parts = checkBox.Name.Split('|');

            // This is a board checkbox
            if (parts.Length == 2)
            {
                Panel parentPanel = checkBox.Parent as Panel;
                if (parentPanel == null) return;

                // Find all schematic checkboxes for this board
                foreach (Control ctrl in parentPanel.Controls)
                {
                    if (ctrl is CheckBox schematicCheckBox && schematicCheckBox != checkBox)
                    {
                        var schematicParts = schematicCheckBox.Name.Split('|');
                        if (schematicParts.Length == 3 &&
                            schematicParts[0] == parts[0] && schematicParts[1] == parts[1])
                        {
                            schematicCheckBox.Enabled = checkBox.Checked;
                            if (checkBox.Checked)
                            {
                                schematicCheckBox.Checked = true; // Select all siblings when board is enabled
                            }
                            else
                            {
                                schematicCheckBox.Checked = false;
                            }
                        }
                    }
                }
            }

            // This is a schematic checkbox
            else if (parts.Length == 3)
            {
                Panel parentPanel = checkBox.Parent as Panel;
                if (parentPanel == null) return;

                // Find the parent board checkbox
                CheckBox boardCheckBox = null;
                bool anySchematicChecked = false;
                foreach (Control ctrl in parentPanel.Controls)
                {
                    if (ctrl is CheckBox cb)
                    {
                        var cbParts = cb.Name.Split('|');
                        if (cbParts.Length == 2 && cbParts[0] == parts[0] && cbParts[1] == parts[1])
                        {
                            boardCheckBox = cb;
                        }
                        else if (cbParts.Length == 3 && cbParts[0] == parts[0] && cbParts[1] == parts[1])
                        {
                            if (cb.Checked)
                                anySchematicChecked = true;
                        }
                    }
                }

                // If no schematic is checked, uncheck the board
                if (boardCheckBox != null && !anySchematicChecked)
                {
                    boardCheckBox.Checked = false;
                }
            }
        }


        // ###########################################################################################
        // Initialize the tab for "Help".
        // ###########################################################################################

        private void InitializeTabHelp()
        {
            try
            {
                richTextBoxHelp.LinkClicked -= RichTextBoxHelp_LinkClicked;
                richTextBoxHelp.LinkClicked += RichTextBoxHelp_LinkClicked;

                richTextBoxHelp.Rtf = BuildHelpRtf();

                // "Padding" inside the control
                SetRichTextBoxPadding(richTextBoxHelp, new Padding(10, 10, 10, 10));
                richTextBoxHelp.Resize -= RichTextBoxHelp_Resize_SetPadding;
                richTextBoxHelp.Resize += RichTextBoxHelp_Resize_SetPadding;
            }
            catch { /* best-effort */ }
        }

        private static string BuildHelpRtf()
        {
            var sb = new StringBuilder();

            // RTF header with font and color tables
            sb.Append(@"{\rtf1\ansi\deff0");
            sb.Append(@"{\fonttbl{\f0 Calibri;}{\f1 Consolas;}}");
            // colortbl: [0]=auto; [1]=text (black); [2]=keycap bg (light gray)
            sb.Append(@"{\colortbl ;\red0\green0\blue0;\red230\green234\blue238;}");
            sb.Append(@"\f0\fs22 "); // Calibri 11pt

            sb.Append(@"{\fs28{\b Help for application usage}}\line\line ");

            sb.Append(@"{{\b ""Schematics"" tab}\line\line ");

            sb.Append(@"{\b Mouse functions:}\line ");
            sb.Append(@"    * "+ KeycapRtf("Left-click") +@" on a component will show a information popup\line ");
            sb.Append(@"    * "+ KeycapRtf("Left-click") + @" + "+ KeycapRtf("Hold") +@" on a component will do one of three things:\line ");
            sb.Append(@"        - Start a new trace, when in ""empty"" space (not on top of component)\line ");
            sb.Append(@"        - Move trace marker, if mouse is on top of an existing trace marker\line ");
            sb.Append(@"        - Insert new trace marker, if mouse is on top of an existing trace, but not on top of a trace marker\line ");
            sb.Append(@"    * " + KeycapRtf("Right-click") + @" on a component will do one of three things:\line ");
            sb.Append(@"        - Toggle component highlight, if mouse is on top of a component\line ");
            sb.Append(@"        - Remove entire trace, if mouse is on top of an existing trace, but not on top of a trace marker\line ");
            sb.Append(@"        - Remove trace marker, if mouse is on top of an existing trace marker\line ");
            sb.Append(@"    * " + KeycapRtf("Right-click") + @" + " + KeycapRtf("Hold") + @" will pan the image\line ");
            sb.Append(@"    * " + KeycapRtf("Scrollwheel") + @" will zoom in/out\line ");
            sb.Append(@"\line ");

            sb.Append(@"{\b Keyboard functions:}\line ");
            sb.Append(@"    * " + KeycapRtf("F11") + @" will toggle fullscreen\line ");
            sb.Append(@"    * " + KeycapRtf("ESCAPE") + @" (in fullscreen) will exit fullscreen\line ");
            sb.Append(@"    * " + KeycapRtf("ENTER") + @" will toggle blinking for selected components\line ");
            sb.Append(@"    * " + KeycapRtf("ALT") + @" + " + KeycapRtf("A") + @" will select all components in ""Component list""\line ");
            sb.Append(@"    * " + KeycapRtf("ALT") + @" + " + KeycapRtf("C") + @" will clear all selections and show all components in ""Component list""\line ");
            sb.Append(@"    * Start typing anywhere to filter component list(view ""Filtering"")\line ");
            sb.Append(@"\line ");

            sb.Append(@"{\b Filtering:}\line ");
            sb.Append(@"    * Filtering supports multi-word / character searching, divided by a whitespace\line ");
            sb.Append(@"    * When doing a multi - word searching then the order is not important\line ");
            sb.Append(@"    * Filtering is case-insensitive\line ");
            sb.Append(@"    * Examples of a multi - word / character search:\line ");
            sb.Append(@"        - Typing "+ KeycapRtf("U6 | CPU | 8502") + @" will find the component "+ KeycapRtf("U6 | CPU | 8502") +@"\line ");
            sb.Append(@"        - Typing "+ KeycapRtf("CPU 8502 U6") + @" will find the component "+ KeycapRtf("U6 | CPU | 8502") +@"\line ");
            sb.Append(@"        - Typing "+ KeycapRtf("5 8 U P c 6") + @" will find the component "+ KeycapRtf("U6 | CPU | 8502") +@"\line ");
            sb.Append(@"\line ");

            sb.Append(@"{\b PAL/NTSC:}\line ");
            sb.Append(@"    * {\b PAL}\line ");
            sb.Append(@"        - Filters all components and images where the region is either set to PAL or not relevant (generic)\line ");
            sb.Append(@"    * {\b NTSC}\line ");
            sb.Append(@"        - Filters all components and images where the region is either set to NTSC or not relevant (generic)\line ");
            sb.Append(@"    * The component information popup shows a counter on the two region buttons for the number of images relevant specifically for this region or if images are generic\line ");
            sb.Append(@"\line ");

            sb.Append(@"{\b Component selection:}\line ");
            sb.Append(@"    * When a component is selected, then it will also visualize if component is part of thumbnail in list - view:\line ");
            sb.Append(@"        - Appending an asterisk/* as first character in thumbnail label\line ");
            sb.Append(@"        - Background color of thumbnail label changes to blue\line ");
            sb.Append(@"    * You cannot highlight a component in image, if its component category is unselected\line ");
            sb.Append(@"\line ");

            sb.Append(@"{\b Component selection:}\line ");
            sb.Append(@"    * When a component is selected, then it will also visualize if component is part of thumbnail in list - view:\line ");
            sb.Append(@"        - Appending an asterisk/* as first character in thumbnail label\line ");
            sb.Append(@"        - Background color of thumbnail label changes to blue\line ");
            sb.Append(@"    * You cannot highlight a component in image, if its component category is unselected\line ");
            sb.Append(@"\line ");

            sb.Append(@"{\b Circuit tracing:}\line ");
            sb.Append(@"    * When mouse cursor shows a ""cross"", then you can start drawing a new trace\line ");
            sb.Append(@"    * Holding down " + KeycapRtf("SHIFT") + @" while drawing a trace, will vertically and horizontally align the trace to its neighbour markers\line ");
            sb.Append(@"\line ");

            sb.Append(@"{\b ""Labels visible"" panel:}\line ");
            sb.Append(@"    * If no components are selected for the specific schematic, then the panel with checkboxes will not be shown\line ");
            sb.Append(@"    * The panel can be toggled minimized/maximized with the ""M"" button\line ");
            sb.Append(@"    * When only one checkbox is selected, then it will replace whitespaces in label with new-lines to condense the text\line ");
            sb.Append(@"\line ");

            sb.Append(@"{\b ""Traces visible"" panel:}\line ");
            sb.Append(@"    * If no traces are drawn for the specific schematic, then the panel with checkboxes will not be shown\line ");
            sb.Append(@"    * The panel can be toggled minimized/maximized with the ""M"" button\line ");
            sb.Append(@"    * Traces can be toggled hidden/shown based on their color - this applies to all traces within the selected hardware/board\line ");
            sb.Append(@"\line ");

            sb.Append(@"------------------------------------------------------------------------------------------------------------------------------------ \line ");
            sb.Append(@"\line ");

            sb.Append(@"{\fs28{\b Component information popup}}\line ");
            sb.Append(@"\line ");

            sb.Append(@"If multiple images are available for the selected component, then it will show ""Image 1 of x"" in the top right corner.\line ");
            sb.Append(@"\line ");

            sb.Append(@"{\b Mouse functions:}\line ");
            sb.Append(@"    * " + KeycapRtf("Left-click") + @" in the image area will change back to first image (typically the pinout)\line ");
            sb.Append(@"    * " + KeycapRtf("Scrollwheel") + @" will change image, if multiple images\line ");
            sb.Append(@"\line ");

            sb.Append(@"{\b Keyboard  functions:}\line ");
            sb.Append(@"    * Arrow keys " + KeycapRtf("LEFT") + @" "+ KeycapRtf("RIGHT") +@" "+ KeycapRtf("UP") + @" "+ KeycapRtf("DOWN") + @" will change image, if multiple images\line ");
            sb.Append(@"    * " + KeycapRtf("SPACE") + @" will change back to first image (typical the pinout)\line ");
            sb.Append(@"    * " + KeycapRtf("CTRL") + @" + " + KeycapRtf("TAB") + @" will toggle between PAL and NTSC\line ");
            sb.Append(@"    * " + KeycapRtf("ESCAPE") + @" will close popup\line ");
            sb.Append(@"\line ");

            sb.Append(@"------------------------------------------------------------------------------------------------------------------------------------ \line ");
            sb.Append(@"\line ");

            sb.Append(@"{\fs28{\b Data update from online source}}\line ");
            sb.Append(@"\line ");

            sb.Append(@"It is possible to fetch the newest data from the online source.\line ");
            sb.Append(@"You can do this via the ""Configuration"" tab.\line ");
            sb.Append(@"\line ");
            sb.Append(@"If you have not modified any data in the Excel data files on your own, then there is no risks in doing this - go for it.\line ");
            sb.Append(@"If you do have modified some data in the Excel data files, then be aware that all Excel data files and all images will be overwritten, so do make a backup before you update.\line ");
            sb.Append(@"The update will not delete any files it does not officially know - e.g. if you have added some of your own files, they will not be deleted.\line ");
            sb.Append(@"The update will not delete any of your own component text modifications done through the component information popup.\line ");
            sb.Append(@"\line ");
            sb.Append(@"The update will happen at the {\b next} application launch, so you will not see the changes immediately.\line ");
            sb.Append(@"\line ");

            sb.Append(@"------------------------------------------------------------------------------------------------------------------------------------ \line ");
            sb.Append(@"\line ");

            sb.Append(@"{\fs28{\b Show or hide hardware and boards}}\line ");
            sb.Append(@"\line ");

            sb.Append(@"In ""Configuration"" you can select which hardware and boards you want to show or hide in the application.\line ");
            sb.Append(@"Per default it will show everything, but you can uncheck the ones you do not want to have in the application.\line ");
            sb.Append(@"In some cases there are redundant schematics, where it both have the original (scanned) Commodore schematics but also a new set of replica schematics, and in these cases you can decide if you want to hide one of these sets (or keeping both visible).\line ");
            sb.Append(@"\line ");

            sb.Append(@"Changing any checkbox will be effectuated at the {\b next} application launch, so you will not see the changes immediately.\line ");
            sb.Append(@"\line ");

            sb.Append(@"------------------------------------------------------------------------------------------------------------------------------------ \line ");
            sb.Append(@"\line ");

            sb.Append(@"{\fs28{\b Misc} }\line ");
            sb.Append(@"\line ");

            sb.Append(@"When there is a newer version available online, it will be marked with an asterisk (*) in the ""About"" tab, and it will show with a red label in top-right corner.\line ");
            sb.Append(@"Then navigate to the ""About"" tab and update to the new version, either via ""Auto-install"" or manually download from https://github.com/HovKlan-DH/Commodore-Repair-Toolbox/releases\line ");
            sb.Append(@"\line ");

            sb.Append(@"How-to add or update your own data:\line ");
            sb.Append(@"    * View https://github.com/HovKlan-DH/Commodore-Repair-Toolbox/wiki/Documentation\line ");
            sb.Append(@"\line ");

            sb.Append(@"Report a problem or comment something from either of these places:\line ");
            sb.Append(@"    * Through the ""Feedback"" tab\line ");
            sb.Append(@"    * Through https://github.com/HovKlan-DH/Commodore-Repair-Toolbox/issues\line ");
            sb.Append(@"\line");

            sb.Append("}");

            return sb.ToString();
        }

        private static string KeycapRtf(string text)
        {
            // base font for keys a tad smaller than body
            var inner = $"\\f1\\fs20\\cf1\\highlight2\\~{EscapeRtf(text)}\\~\\highlight0";
            return "{" + inner + "}";
        }

        private static string EscapeRtf(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return s.Replace("\\", "\\\\").Replace("{", "\\{").Replace("}", "\\}");
        }

        private void RichTextBoxHelp_Resize_SetPadding(object sender, EventArgs e)
        {
            SetRichTextBoxPadding(richTextBoxHelp, new Padding(12, 10, 12, 10));
        }

        private void RichTextBoxHelp_LinkClicked(object sender, LinkClickedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(e.LinkText))
                return;

            try
            {
                OpenExternalTarget(e.LinkText, isUrl: true);
            }
            catch (Exception ex)
            {
                DebugOutput("ERROR: Failed opening URL [" + e.LinkText + "]: " + ex);
            }
        }


        // ###########################################################################################
        // Initialize the tab for "About".
        // ###########################################################################################

        private void InitializeTabAboutNew()
        {
            try
            {
                richTextBoxAbout.DetectUrls = true;

                richTextBoxAbout.LinkClicked -= RichTextBoxHelp_LinkClicked;
                richTextBoxAbout.LinkClicked += RichTextBoxHelp_LinkClicked;

                SetRichTextBoxPadding(richTextBoxAbout, new Padding(10, 10, 10, 10));
                richTextBoxAbout.Resize -= RichTextBoxAbout_Resize_SetPadding;
                richTextBoxAbout.Resize += RichTextBoxAbout_Resize_SetPadding;
            }
            catch { /* best-effort */ }
        }

        private void RichTextBoxAbout_Resize_SetPadding(object sender, EventArgs e)
        {
            SetRichTextBoxPadding(richTextBoxAbout, new Padding(10, 10, 10, 10));
        }

        private void UpdateTabAboutNew()
        {
            InitializeTabAboutNew();

            // Credits for selected board
            var foundHardware = classHardware.FirstOrDefault(h => h.Name == hardwareSelectedName);
            var foundBoard = foundHardware?.Boards.FirstOrDefault(b => b.Name == boardSelectedName);

            richTextBoxAbout.Rtf = BuildAboutRtf(foundBoard);
        }

        private static string BuildAboutRtf(Board board)
        {
            var sb = new StringBuilder();

            // RTF header with font + color tables
            sb.Append(@"{\rtf1\ansi\deff0");
            sb.Append(@"{\fonttbl{\f0 Calibri;}{\f1 Consolas;}}");

            // "colortbl" index:
            // 0 = auto
            // 1 = black
            // 2 = IndianRed
            // 3 = keycap bg (light gray) (if you want to keep it around)
            sb.Append(@"{\colortbl ;\red0\green0\blue0;\red205\green92\blue92;\red230\green234\blue238;}");
            sb.Append(@"\f0\fs22 "); // Calibri 11pt

            sb.Append(@"{\fs28\b Commodore Repair Toolbox\b0}\line\line ");

            sb.Append(@"You are running version \b ");
            sb.Append(EscapeRtf(versionThis));
            sb.Append(@"\b0  (64-bit)\line\line ");

            if (!string.IsNullOrWhiteSpace(versionOnline))
            {
//                sb.Append(@"\cf2 There is a newer version available. Click the above ""{\b Auto-install New Version\b0}"" button to install it, or download it manually from below link.\line ");
                sb.Append(@"\cf2 There is a newer version available. Download it manually from below link.\line ");
//                sb.Append(@"View the {\i Changelog\i0} here, https://github.com/HovKlan-DH/Commodore-Repair-Toolbox/releases\line\line ");
                sb.Append(@"View the {\i Changelog\i0} here, https://github.com/HovKlan-DH/Classic-Repair-Toolbox/releases\line\line ");
                sb.Append(@"\cf0");
            }

            sb.Append(@"All programming done by Dennis Helligsø (dennis@commodore-repair-toolbox.dk).\line\line ");

//            sb.Append(@"Visit official project home page at https://github.com/HovKlan-DH/Commodore-Repair-Toolbox\line\line ");
            sb.Append(@"Visit official project home page at https://github.com/HovKlan-DH/Classic-Repair-Toolbox\line\line ");

            sb.Append(@"------------------------------------------------------------------------------------------------------------------------------------\line\line ");

            sb.Append(@"Credits and recognition for this board data go to:\line\line ");
            AppendBoardCreditsRtf(sb, board);

            sb.Append(@"\line ");
            sb.Append(@"------------------------------------------------------------------------------------------------------------------------------------\line\line ");

            sb.Append(@"\b A comment from the developer, Dennis:\b0\line\line ");

            sb.Append(@"\i ");
            sb.Append(EscapeRtf(
                "I have been repairing Commodore 64/128 computers for some years, but I still consider myself as a novice in this world of hardware - I am more a software person, which you may have guessed having this tool here. I often forget where and what to check, and I struggle to find again all the relevant resources and schematics, not to mention the struggle to find the components in the schematics - a pure mess and quite inefficient. I did often refer to the \"Mainboards\" section of https://myoldcomputer.nl/technical-info/mainboards, and I noticed that Jeroen did have a prototype of an application named \"Repair Help\", and it did have the easy layout I was looking for (I did get a copy of it). However, it was never finalized from him, so I took upon myself to create something similar, and a couple of years later (including a long hiatus) I did come up with this quite similar looking application, though expanded with additional functionalities and data points."
            ));
            sb.Append(@"\line\line ");

            sb.Append(EscapeRtf(
                "The longer-term goal is that the tool will cover all the earlier Commodore computers, and ideally also its most used peripherals, and I will continue to add new and refine data for myself (when doing my own diagnostics and repairing), but I will most likely not be able to do this myself alone. I would really appreciate some help with this, so if you have the willingness to contribute, then please reach out to me, and I will happily explain the nitty-gritty details. It is actually quite easy when having tried it once."
            ));
            sb.Append(@"\line\line ");

            sb.Append(EscapeRtf(
                "If you see anything that can be better, e.g. bad data quality or improvements for the tool, then do reach out to me. Of course I would also be happy, if you would send a comment from the \"Feedback\" tab - even if you do not like the tool, as constructive criticism is always welcome :-)"
            ));
            sb.Append(@"\line\line ");

            sb.Append(EscapeRtf("// Dennis"));
            sb.Append(@"\i0\line ");

            sb.Append("}");
            return sb.ToString();
        }

        private static void AppendBoardCreditsRtf(StringBuilder sb, Board board)
        {
            if (board?.BoardCredits == null || board.BoardCredits.Count == 0)
            {
                sb.Append(@"(No credits found for this board)\line\line ");
                return;
            }

            var grouped = board.BoardCredits
                .GroupBy(c => c.Category)
                .ToDictionary(
                    g => g.Key,
                    g => g.GroupBy(c2 => string.IsNullOrWhiteSpace(c2.SubCategory) ? "Unnamed" : c2.SubCategory)
                          .ToDictionary(sg => sg.Key, sg => sg.ToList())
                );

            foreach (var category in grouped)
            {
                sb.Append(@"\b ");
                sb.Append(EscapeRtf(category.Key));
                sb.Append(@"\b0:\line ");

                foreach (var subcat in category.Value)
                {
                    if (!string.Equals(subcat.Key, "Unnamed", StringComparison.OrdinalIgnoreCase))
                    {
                        sb.Append(@"        \b ");
                        sb.Append(EscapeRtf(subcat.Key));
                        sb.Append(@"\b0\line ");
                    }

                    foreach (var credit in subcat.Value)
                    {
                        sb.Append(@"                ");
                        sb.Append(FormatBulletRtf(credit.Contact == "" ? credit.Name : (credit.Name + ", " + credit.Contact)));
                        sb.Append(@"\line ");
                    }
                }

                sb.Append(@"\line ");
            }

            sb.Append(@"When people contribute a substantial amount of data, they can choose to be listed here.\line ");
            sb.Append(@"Either the real name or a handle can be chosen and contact address is optional (email, GitHub or personal web page).\line ");
        }

        private static string FormatBulletRtf(string text)
        {
            // Simple bullet with hanging indent, no nested list semantics
            return @"{\pard\li360\fi-180\pntext\u8226?  " + EscapeRtf(text) + @"}";
        }


        // ###########################################################################################
        // Initialize the tab for "About".
        // ###########################################################################################

        private void InitializeBlinkTimer()
        {
            blinkTimer = new Timer();
            blinkTimer.Interval = 500;
            blinkTimer.Tick += BlinkTimer_Tick;
        }


        // ###########################################################################################
        // What happens when a combobox is closed.
        // Applicable for "Hardware" or "Board" comboboxes.
        // ###########################################################################################

        private void ComboBox_DropDownClosed(object sender, EventArgs e)
        {
            textBoxFilterComponents.Text = "";
            textBoxFilterComponents.Focus();
        }



        // ###########################################################################################
        // Do an update of the component list, taking selection and filtering in to consideration.
        // ###########################################################################################

        private void UpdateComponentSelection()
        {
            // Debug
#if DEBUG
            StackTrace stackTrace = new StackTrace();
            StackFrame callerFrame = stackTrace.GetFrame(1);
            MethodBase callerMethod = callerFrame.GetMethod();
            string callerName = callerMethod.Name;
            Debug.WriteLine("[UpdateComponentSelection] called from [" + callerName + "]");
#endif

            // Build a Display -> Label map once (outside the loop)
            var hw = classHardware.FirstOrDefault(h => h.Name == hardwareSelectedName);
            var bd = hw?.Boards.FirstOrDefault(b => b.Name == boardSelectedName);
            var displayToLabel = bd?.Components
                .Where(c => !string.IsNullOrEmpty(c.NameDisplay))
                .GroupBy(c => c.NameDisplay)
                .ToDictionary(g => g.Key, g => g.First().Label); // assume NameDisplay unique

            var listBoxComponentsSelectedClone = listBoxComponents.SelectedItems.Cast<object>().ToList();

            foreach (var item in listBoxComponents.Items)
            {
                string display = item.ToString();
                // Fallback to display if no mapping (prevents null issues)
                string label = (displayToLabel != null && displayToLabel.TryGetValue(display, out var lbl))
                    ? lbl
                    : display;

                if (listBoxComponentsSelectedClone.Contains(item))
                {
                    AddSelectedComponentIfNotInList(label);
                }
                else
                {
                    RemoveSelectedComponentIfInList(label);
                }
            }

            UpdateShowOfSelectedComponents();

            // Update overlays
            ShowOverlaysAccordingToComponentList();
        }

        private void UpdateComponentList(string from)
        {
#if DEBUG
            StackTrace stackTrace = new StackTrace();
            StackFrame callerFrame = stackTrace.GetFrame(1);
            MethodBase callerMethod = callerFrame.GetMethod();
            string callerName = callerMethod.Name;
            Debug.WriteLine("[UpdateComponentList] called from [" + callerName + "]");
#endif

            // Decouple the event handler to avoid continuous recycling
            listBoxComponents.SelectedIndexChanged -= listBoxComponents_SelectedIndexChanged;

            var listBoxComponentsSelectedClone = listBoxComponents.SelectedItems.Cast<object>().ToList(); // create a list of selected components
            string filterText = textBoxFilterComponents.Text.ToLower();
            string[] searchTerms = filterText.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries); // Split filter text into terms

            // Deselect all components if we have an empty filter search
            if (from == "TextBoxFilterComponents_TextChanged" && searchTerms.Length == 0)
            {
                listBoxComponentsSelectedClone.Clear();
                listBoxComponentsSelectedText.Clear();
            }

            // Convert the selected items to a "HashSet" for faster lookup
            var selectedComponents = new HashSet<string>(listBoxComponentsSelectedText);

            listBoxComponents.BeginUpdate(); // suspend redrawing this listBox while updating it
            listBoxComponents.Items.Clear();

            // Walk through all components for selected hardware and board
            var foundHardware = classHardware.FirstOrDefault(h => h.Name == hardwareSelectedName);
            var foundBoard = foundHardware?.Boards.FirstOrDefault(b => b.Name == boardSelectedName);
            if (foundBoard != null && foundBoard.Components != null)
            {
                foreach (BoardComponents comp in foundBoard.Components)
                {
                    string componentLabel = comp.Label;
                    string componentCategory = comp.Type;
                    string componentDisplay = comp.NameDisplay;

                    // Region filtering
                    bool regionMatches = string.IsNullOrEmpty(comp.Region) ||
                                         string.IsNullOrEmpty(Main.selectedRegion) ||
                                         string.Equals(comp.Region, Main.selectedRegion, StringComparison.OrdinalIgnoreCase);

                    if (!regionMatches)
                        continue; // skip this component if region does not match

                    if (listBoxCategories.SelectedItems.Contains(componentCategory))
                    {
                        // Check if the component matches all search terms
                        bool matchesAllTerms = searchTerms.Length == 0 || searchTerms.All(term => componentDisplay.ToLower().Contains(term));

                        if (matchesAllTerms)
                        {
                            listBoxComponents.Items.Add(componentDisplay);

                            if (listBoxComponentsSelectedClone.Contains(componentDisplay) || (from == "TextBoxFilterComponents_TextChanged" && searchTerms.Length > 0))
                            {
                                if (!selectedComponents.Contains(componentLabel))
                                {
                                    AddSelectedComponentIfNotInList(componentLabel);
                                }
                                int index = listBoxComponents.Items.IndexOf(componentDisplay);
                                if (index >= 0)
                                {
                                    listBoxComponents.SetSelected(index, true);
                                }
                            }
                            else
                            {
                                RemoveSelectedComponentIfInList(componentLabel);
                            }
                        }
                        else
                        {
                            RemoveSelectedComponentIfInList(componentLabel);
                        }
                    }
                    else
                    {
                        RemoveSelectedComponentIfInList(componentLabel);
                    }
                }
            }

            listBoxComponents.EndUpdate(); // resume redrawing of this specific listBox

            UpdateShowOfSelectedComponents();
            ShowOverlaysAccordingToComponentList();
            UpdateTabNewOverview(GetSelectedBoardClass());

            listBoxComponents.SelectedIndexChanged += listBoxComponents_SelectedIndexChanged;
        }

        private void AddSelectedComponentIfNotInList(string componentLabel)
        {
            if (!listBoxComponentsSelectedText.Contains(componentLabel))
            {
                listBoxComponentsSelectedText.Add(componentLabel);
            }

        }

        private void RemoveSelectedComponentIfInList(string componentLabel)
        {
            if (listBoxComponentsSelectedText.Contains(componentLabel))
            {
                listBoxComponentsSelectedText.Remove(componentLabel);
            }
        }

        private void listBoxComponents_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateComponentSelection();
        }

        private void UpdateShowOfSelectedComponents()
        {
            if (listBoxComponentsSelectedText.Count > 0)
            {
                labelComponents.Text = $"Component list ({listBoxComponentsSelectedText.Count} selected)";
            }
            else
            {
                labelComponents.Text = "Component list";
            }
        }


        // ###########################################################################################
        // Handling highlighting of overlays for "Main" and thumbnail images.
        // ###########################################################################################

        private void HighlightOverlays(string scope)
        {
            // Debug
#if DEBUG
            StackTrace stackTrace = new StackTrace();
            StackFrame callerFrame = stackTrace.GetFrame(1);
            MethodBase callerMethod = callerFrame.GetMethod();
            string callerName = callerMethod.Name;
            Debug.WriteLine("[HighlightOverlays(" + scope + ")] called from [" + callerName + "]");
#endif

            var hw = classHardware.FirstOrDefault(h => h.Name == hardwareSelectedName);
            var bd = hw?.Boards.FirstOrDefault(b => b.Name == boardSelectedName);
            if (bd == null) return;

            // Convert the selected items to a "HashSet" for faster lookup
            var selectedComponents = new HashSet<string>(listBoxComponentsSelectedText);

            // Main
            if (scope == "tab")
            {
                // Draw overlays on the main image
                if (overlayPanel == null) return;
                var bf = bd?.Files?.FirstOrDefault(f => f.Name == schematicSelectedName);
                if (bf == null) return;

                DisposeAllControls(overlayPanel);
                overlayPanel.Overlays.Clear();

                Color colorZoom = Color.FromName(bf.HighlightColorTab);
                int opacityZoom = bf.HighlightOpacityTab;

                bool hasHighlightedOverlays = false;

                for (int i = 0; i < overlayComponentsTab.Count; i++)
                {
                    Rectangle rect = new Rectangle(
                        (int)(overlayComponentsTabOriginalLocations[i].X * zoomFactor),
                        (int)(overlayComponentsTabOriginalLocations[i].Y * zoomFactor),
                        (int)(overlayComponentsTabOriginalSizes[i].Width * zoomFactor),
                        (int)(overlayComponentsTabOriginalSizes[i].Height * zoomFactor)
                    );

                    string componentLabel = bd.Components.FirstOrDefault(cb => cb.Label == overlayComponentsTab[i].Name)?.Label ?? "";

                    // Get "technical name" - first with region, then first or otherwise it should be empty
                    var compTech = bd.Components
                        .FirstOrDefault(cb => cb.Label == overlayComponentsTab[i].Name &&
                                              !string.IsNullOrEmpty(selectedRegion) &&
                                              !string.IsNullOrEmpty(cb.Region) &&
                                              string.Equals(cb.Region, selectedRegion, StringComparison.OrdinalIgnoreCase))
                        ?? bd.Components
                        .FirstOrDefault(cb => cb.Label == overlayComponentsTab[i].Name &&
                                              string.IsNullOrEmpty(cb.Region))
                        ?? bd.Components
                        .FirstOrDefault(cb => cb.Label == overlayComponentsTab[i].Name);
                    string componentTechName = compTech?.NameTechnical ?? "";

                    // Get "friendly name" - first with region, then first or otherwise it should be empty
                    var compFriendly = bd.Components
                        .FirstOrDefault(c => c.Label == overlayComponentsTab[i].Name &&
                                             !string.IsNullOrEmpty(selectedRegion) &&
                                             !string.IsNullOrEmpty(c.Region) &&
                                             string.Equals(c.Region, selectedRegion, StringComparison.OrdinalIgnoreCase))
                        ?? bd.Components
                        .FirstOrDefault(c => c.Label == overlayComponentsTab[i].Name &&
                                             string.IsNullOrEmpty(c.Region))
                        ?? bd.Components
                        .FirstOrDefault(c => c.Label == overlayComponentsTab[i].Name);
                    string componentFriendlyName = compFriendly?.NameFriendly ?? "";
                    string componentDisplay = bd.Components
                        .FirstOrDefault(cb => cb.Label == overlayComponentsTab[i].Name
                        && string.Equals(cb.Region, selectedRegion, StringComparison.OrdinalIgnoreCase))
                        ?.NameDisplay ?? "";

                    // Check if the component is selected in component list
                    bool highlighted = selectedComponents.Contains(overlayComponentsTab[i].Name);

                    overlayPanel.Overlays.Add(new OverlayInfo
                    {
                        Bounds = rect,
                        Color = colorZoom,
                        Opacity = opacityZoom,
                        Highlighted = highlighted,
                        ComponentLabel = componentLabel,
                        ComponentDisplay = overlayComponentsTab[i].Name
                    });

                    // Create a label, if the component is highlighted
                    if (highlighted)
                    {
                        hasHighlightedOverlays = true;

                        string labelText = checkBox1.Checked ? ConvertStringToLabel(componentLabel) : "";
                        labelText += checkBox2.Checked && componentTechName != "" ? Environment.NewLine + ConvertStringToLabel(componentTechName) : "";
                        labelText += checkBox3.Checked && componentFriendlyName != "" ? Environment.NewLine + ConvertStringToLabel(componentFriendlyName) : "";

                        // Overwrite the label text, if only one checkbox is enabled
                        labelText = checkBox1.Checked && !checkBox2.Checked && !checkBox3.Checked ? ConvertStringToLabel(componentLabel.Replace(" ", Environment.NewLine)) : labelText;
                        labelText = !checkBox1.Checked && checkBox2.Checked && !checkBox3.Checked ? ConvertStringToLabel(componentTechName.Replace(" ", Environment.NewLine)) : labelText;
                        labelText = !checkBox1.Checked && !checkBox2.Checked && checkBox3.Checked ? ConvertStringToLabel(componentFriendlyName.Replace(" ", Environment.NewLine)) : labelText;

                        // Weird command but it will remove empty newlines at beginning and end
                        labelText = string.Join(Environment.NewLine, labelText.Split(new[] { Environment.NewLine }, StringSplitOptions.None).Where(line => !string.IsNullOrWhiteSpace(line)));

                        // Add a component label, if one of the checkboxes are checked
                        if (checkBox1.Checked || checkBox2.Checked || checkBox3.Checked)
                        {
                            Label label = new Label
                            {
                                Name = "Label for component",
                                Text = labelText,
                                AutoSize = true,
                                BackColor = Color.Khaki,
                                ForeColor = Color.Black,
                                Font = new Font("Calibri", 9, FontStyle.Bold),
                                BorderStyle = BorderStyle.FixedSingle,
                                TextAlign = ContentAlignment.MiddleCenter,
                                Enabled = false,
                                Tag = componentLabel,
                            };
                            label.DoubleBuffered(true);

                            // Calculate the label center/middle position within the rectangle
                            int labelX = rect.X + (rect.Width - label.PreferredWidth) / 2;
                            int labelY = rect.Y + (rect.Height - label.PreferredHeight) / 2;
                            label.Location = new Point(labelX, labelY);

                            // Have a mapping table, as this isfaster than seaching all controls for a specific name
                            overlayLabelMap[componentLabel] = label;

                            overlayPanel.Controls.Add(label);
                        }
                    }
                }
                overlayPanel.Invalidate();

                // Ensure the "Labels visible" panel is visible
                if (hasHighlightedOverlays)
                {
                    // Add the label panel to "panelMain"
                    panelMain.Controls.Add(panelLabelsVisible);
                    int visibleHeight = panelZoom.ClientRectangle.Height;
                    panelLabelsVisible.Location = new Point(0, visibleHeight - panelLabelsVisible.Height - 2); // bottom-left corner of the visible area
                    panelLabelsVisible.BringToFront();
                    panelLabelsVisible.Visible = true;
                }
                else
                {
                    panelLabelsVisible.Visible = false;
                }
            }

            // Thumbnail list
            else if (scope == "list")
            {
                if (bd.Files != null)
                {
                    // Draw overlays on each thumbnail
                    foreach (BoardOverlays bo in bd.Files)
                    {
                        if (!overlayPanelsList.ContainsKey(bo.Name)) continue;

                        OverlayPanel listPanel = overlayPanelsList[bo.Name];
                        listPanel.Overlays.Clear();

                        Color colorList = Color.FromName(bo.HighlightColorList);
                        int opacityList = bo.HighlightOpacityList;
                        float listZoom = overlayListZoomFactors[bo.Name];

                        if (bo?.Components != null)
                        {
                            foreach (var comp in bo.Components)
                            {
                                if (comp.Overlays == null) continue;

                                // Find component "display name"
                                string componentDisplay = bd.Components
                                    .FirstOrDefault(cb => cb.Label == comp.Label)?.NameDisplay ?? "";

                                // Check if the component is selected in component list
                                bool highlighted = selectedComponents.Contains(comp.Label);

                                foreach (var ov in comp.Overlays)
                                {
                                    Rectangle rect = new Rectangle(
                                        (int)(ov.Bounds.X * listZoom),
                                        (int)(ov.Bounds.Y * listZoom),
                                        (int)(ov.Bounds.Width * listZoom),
                                        (int)(ov.Bounds.Height * listZoom)
                                    );
                                    listPanel.Overlays.Add(new OverlayInfo
                                    {
                                        Bounds = rect,
                                        Color = colorList,
                                        Opacity = opacityList,
                                        Highlighted = highlighted,
                                        ComponentLabel = comp.Label,
                                        ComponentDisplay = componentDisplay
                                    });
                                }
                            }
                        }
                        listPanel.Invalidate();
                    }
                }
            }
        }


        // ###########################################################################################
        // Filtering of components.
        // Dividing string with whitespaces will do a multi-word search, where the
        // order is not imporant.
        // ###########################################################################################

        private void TextBoxFilterComponents_TextChanged(object sender, EventArgs e)
        {
            UpdateComponentList("TextBoxFilterComponents_TextChanged");
        }


        private void ListBoxCategories_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Debug
#if DEBUG
            StackTrace stackTrace = new StackTrace();
            StackFrame callerFrame = stackTrace.GetFrame(1);
            MethodBase callerMethod = callerFrame.GetMethod();
            string callerName = callerMethod.Name;
            Debug.WriteLine("[ListBoxCategories_SelectedIndexChanged] called from [" + callerName + "]");
#endif

            SaveSelectedCategories();
            UpdateComponentList("ListBoxCategories_SelectedIndexChanged");
        }


        // ###########################################################################################
        // Save the selected categories.
        // ###########################################################################################              

        private void SaveSelectedCategories()
        {
            // 1) Build a unique config key from hardware and board
            string configKey = $"SelectedCategories|{hardwareSelectedName}|{boardSelectedName}";

            // 2) Gather selected categories from "listBoxCategories"
            var selectedCategories = listBoxCategories.SelectedItems
                .Cast<object>()
                .Select(item => item.ToString());

            string joined = string.Join(";", selectedCategories);

            Configuration.SaveSetting(configKey, joined);
        }


        // ###########################################################################################
        // Load the selected categories from configuration file.
        // ###########################################################################################

        private bool LoadSelectedCategories()
        {
            // Debug
#if DEBUG
            StackTrace stackTrace = new StackTrace();
            StackFrame callerFrame = stackTrace.GetFrame(1);
            MethodBase callerMethod = callerFrame.GetMethod();
            string callerName = callerMethod.Name;
            Debug.WriteLine("[LoadSelectedCategories] called from [" + callerName + "]");
#endif

            // Get "SelectedCategories" from configuration file
            string configKey = $"SelectedCategories|{hardwareSelectedName}|{boardSelectedName}";
            string joined = Configuration.GetSetting(configKey, "");

            if (string.IsNullOrEmpty(joined)) return false;

            listBoxCategories.ClearSelected();
            string[] categories = joined.Split(';');
            foreach (string cat in categories)
            {
                int idx = listBoxCategories.Items.IndexOf(cat);
                if (idx >= 0)
                {
                    listBoxCategories.SetSelected(idx, true);
                }
            }
            return true;
        }


        // ###########################################################################################
        // Blink handling.
        // ###########################################################################################

        private void CheckBoxBlink_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxBlink.Checked)
            {
                BlinkTimer_Tick(null, null); // perform the first blink immediately
                blinkTimer.Start();
            }
            else
            {
                blinkTimer.Stop();
                ReHighlightSelectedComponents(); // make sure the components will get highlighted again after end blinking
            }
        }

        private void BlinkTimer_Tick(object sender, EventArgs e)
        {
            BlinkSelectedOverlays(blinkState);
            blinkState = !blinkState;
        }

        private void BlinkSelectedOverlays(bool state)
        {
            SuspendDrawing(panelImageMain);

            // Convert the selected items to a "HashSet" for faster lookup
            var selectedComponents = new HashSet<string>(listBoxComponentsSelectedText);

            // Handle the main/schematic image
            if (overlayPanel != null)
            {
                foreach (var overlay in overlayPanel.Overlays)
                {
                    //if (selectedComponents.Contains(overlay.ComponentDisplay))
                    if (selectedComponents.Contains(overlay.ComponentLabel))
                    {
                        overlay.Highlighted = state;

                        if (overlayLabelMap.TryGetValue(overlay.ComponentLabel, out var label))
                        {
                            label.Visible = state;
                        }
                    }
                }
            }

            ResumeDrawing(panelImageMain);

            // Handle the thumbnails
            foreach (var overlayPanel2 in overlayPanelsList.Values)
            {
                foreach (var overlay in overlayPanel2.Overlays)
                {
                    //if (selectedComponents.Contains(overlay.ComponentDisplay))
                    if (selectedComponents.Contains(overlay.ComponentLabel))
                    {
                        overlay.Highlighted = state;
                    }
                }
                overlayPanel2.Invalidate();
            }
        }

        private void ReHighlightSelectedComponents()
        {
            ShowOverlaysAccordingToComponentList();
        }


        // ###########################################################################################
        // Handle input of email address in "Feedback" tab.
        // ###########################################################################################

        private void TextBoxEmail_TextChanged(object sender, EventArgs e)
        {
            string email = textBoxEmail.Text;
            Configuration.SaveSetting("UserEmail", email);
        }


        // ###########################################################################################
        // Changing tab.
        // ###########################################################################################

        private void TabControl_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (previousTab == tabFeedback && tabControl.SelectedTab != tabFeedback)
            {
                // Changed TO tabFeedback
                textBoxFilterComponents.TextChanged += TextBoxFilterComponents_TextChanged;
                AttachClickEventsToFocusFilterComponents(this);
                textBoxFilterComponents.Focus();
            }
            else if (previousTab != tabFeedback && tabControl.SelectedTab == tabFeedback)
            {
                // Changed away FROM tabFeedback
                textBoxFilterComponents.TextChanged -= TextBoxFilterComponents_TextChanged;
                RemoveClickEventsToFocusFilterComponents();
                textBoxFeedback.Focus();
            }

            // Make sure to resize thumbnails, in case some resizing has happend in between
            if (tabControl.SelectedTab == tabSchematics)
            {
                ReadaptThumbnails();
            }

            previousTab = tabControl.SelectedTab;
        }


        // ###########################################################################################
        // Repaint the borders for all thumbnails (selected thumbnail will get marked)
        // ###########################################################################################

        private void DrawBorderInList()
        {
            foreach (Panel container in panelThumbnails.Controls.OfType<Panel>())
            {
                container.Paint -= Panel_Paint_Special;
                container.Paint += Panel_Paint_Special;
                container.Invalidate();
            }
        }

        private void Panel_Paint_Special(object sender, PaintEventArgs e)
        {
            Panel panel = (Panel)sender;
            string selectedContainer = schematicSelectedName + "_container";
            if (panel.Name == selectedContainer)
            {
                float penWidth = thumbnailSelectedBorderWidth;
                using (Pen pen = new Pen(Color.Red, penWidth))
                {
                    pen.DashStyle = System.Drawing.Drawing2D.DashStyle.Dash;
                    float offset = penWidth / 2;
                    e.Graphics.DrawRectangle(
                        pen,
                        offset,
                        offset,
                        panel.ClientSize.Width - penWidth,
                        panel.ClientSize.Height - penWidth
                    );
                }
            }
        }


        // ###########################################################################################
        // Setup the board after user selects it in "comboboxBoard".
        // ###########################################################################################

        private void SetupNewBoard()
        {
#if DEBUG
            StackTrace stackTrace = new StackTrace();
            StackFrame callerFrame = stackTrace.GetFrame(1);
            MethodBase callerMethod = callerFrame.GetMethod();
            string callerName = callerMethod.Name;
            Debug.WriteLine("[SetupNewBoard] called from [" + callerName + "]");
#endif

            listBoxComponents.Items.Clear();

            boardSelectedName = comboBoxBoard.SelectedItem.ToString();
            textBox1.Text = ConvertStringToLabel(boardSelectedName); // feedback info

            var selectedBoardClass = GetSelectedBoardClass();
            if (selectedBoardClass == null) return;

            // Save the selected board for the hardware
            string configKey = $"SelectedBoard|{hardwareSelectedName}";
            Configuration.SaveSetting(configKey, boardSelectedName);

            boardSelectedFilename = selectedBoardClass.DataFile;

            // Load selected thumbnail from configuration file, if already set
            configKey = $"SelectedThumbnail|{hardwareSelectedName}|{boardSelectedName}";
            schematicSelectedName = Configuration.GetSetting(configKey, null);

            // Select the schematic - check if we can find the current selection, but otherwise default to first schematic
            var selectedSchematic = selectedBoardClass?.Files?.FirstOrDefault(f => f.Name == schematicSelectedName);
            if (selectedSchematic == null)
            {
                selectedSchematic = selectedBoardClass?.Files?.FirstOrDefault();
            }
            schematicSelectedName = selectedSchematic?.Name;
            schematicSelectedFile = selectedSchematic?.SchematicFileName;
            textBox2.Text = ConvertStringToLabel(schematicSelectedName); // feedback info

            // Initialize UI
            InitializeComponentCategories();

            // If no config was found (or empty):
            listBoxCategories.BeginUpdate(); // suspend redrawing this listBox while updating it
            bool loaded = LoadSelectedCategories(); // attempt to select from config
            if (!loaded && listBoxCategories.Items.Count > 0)
            {
                // fallback: auto-select everything
                for (int i = 0; i < listBoxCategories.Items.Count; i++)
                {
                    listBoxCategories.SetSelected(i, true);
                }
            }
            listBoxCategories.EndUpdate(); // resume redrawing of this specific listBox

            LoadAndApplySplitterPosition();

            SuspendLayout();
            InitializeThumbnails();
            InitializeTabMain();
            UpdateTabResourcesNew(selectedBoardClass);
            UpdateTabAboutNew();

            // Load polylines after initializing thumbnails and tabs
            PolylinesManagement.LoadPolylines();
            PopulatePolylineVisibilityPanel();

            ResumeLayout();
        }

        private void LoadAndApplySplitterPosition()
        {
            // Debug
#if DEBUG
            StackTrace stackTrace = new StackTrace();
            StackFrame callerFrame = stackTrace.GetFrame(1);
            MethodBase callerMethod = callerFrame.GetMethod();
            string callerName = callerMethod.Name;
            Debug.WriteLine("[LoadAndApplySplitterPosition] called from [" + callerName + "]");
#endif
            
            var splitter = splitContainerSchematics;

            // Skip if layout/handle is not ready yet
            if (!splitter.IsHandleCreated || splitter.Width <= 0)
            {
                return;
            }

            // Compute min/max allowed distances based on current size and constraints
            int minDistance = splitter.Panel1MinSize;
            int maxDistance = splitter.Width - splitter.SplitterWidth - splitter.Panel2MinSize;

            // If constraints can’t be satisfied (e.g., window temporarily too small), skip now
            if (maxDistance < minDistance)
            {
                Debug.WriteLine($"[LoadAndApplySplitterPosition] Skipped (min={minDistance}, max={maxDistance}, width={splitter.Width}).");
                DebugOutput("WARNING: Failed to set [SplitterDistance]");
                return;
            }

            // Default to 90% of width, but clamp it immediately
            int defaultDistance = (int)Math.Round(splitter.Width * 0.9);
            defaultDistance = Math.Max(minDistance, Math.Min(defaultDistance, maxDistance));

            string configKey = $"SplitterPosition|{windowState}|{hardwareSelectedName}|{boardSelectedName}";
            string splitterPosVal = Configuration.GetSetting(configKey, defaultDistance.ToString());

            int desired = defaultDistance;
            if (!int.TryParse(splitterPosVal, out desired) || desired <= 0)
            { 
                desired = defaultDistance;
            }

            // Clamp desired to the valid range for the current size
            int clamped = Math.Max(minDistance, Math.Min(desired, maxDistance));

            try
            {
                if (splitter.SplitterDistance != clamped)
                {
                    splitter.SplitterDistance = clamped;

                    // Persist the clamped value to avoid future out-of-range loads
                    Configuration.SaveSetting(configKey, clamped.ToString());
                }
            }
            catch (InvalidOperationException ex)
            {
                DebugOutput("ERROR: Failed to set SplitterDistance: " + ex.Message);
            }
        }


        // ###########################################################################################
        // Get the class of the selected board.
        // ###########################################################################################

        public Board GetSelectedBoardClass()
        {
            var hardware = classHardware.FirstOrDefault(h => h.Name == hardwareSelectedName);
            return hardware?.Boards.FirstOrDefault(b => b.Name == boardSelectedName);
        }


        // ###########################################################################################
        // Convert special characters, so they can be shown in labels.
        // Currently I only know of "&" being a problem?
        // ###########################################################################################

        public static string ConvertStringToLabel(string str)
        {
            if (string.IsNullOrEmpty(str))
            {
                return str;
            }
            return str.Replace("&", "&&");
        }


        // ###########################################################################################
        // Enable double-buffering for smoother UI rendering.
        // ###########################################################################################

        private void EnableDoubleBuffering()
        {
            SetStyle(
                ControlStyles.AllPaintingInWmPaint |
                ControlStyles.OptimizedDoubleBuffer |
                ControlStyles.UserPaint,
                true
            );
            this.UpdateStyles();

            panelMain.DoubleBuffered(true);
            panelThumbnails.DoubleBuffered(true);
        }


        // ###########################################################################################
        // Enter fullscreen.
        // ###########################################################################################

        private void FullscreenModeEnter()
        {
            // Save (in variable only - not config file) current and set new window state
            formPreviousWindowState = WindowState;
            formPreviousFormBorderStyle = FormBorderStyle;
            previousBoundsForm = Bounds;
            previousBoundsPanelBehindTab = panelBehindTab.Bounds;
            previousBoundsFullscreenButton = buttonFullscreen.Bounds;
            FormBorderStyle = FormBorderStyle.None;
            WindowState = FormWindowState.Normal;
            Bounds = Screen.PrimaryScreen.Bounds;

            // Set bounds for fullscreen panel
            panelBehindTab.Dock = DockStyle.Fill;
            panelBehindTab.BringToFront();

            // Determine which tab should be maximized
            panelBehindTab.Controls.Remove(panelMain);
            panelBehindTab.Controls.Add(panelMain);

            // We hould not populate the filter textbox when inside fullscreen
            textBoxFilterComponents.Enabled = false;

            // Hide tabs, and show fullscreen panel
            tabControl.Visible = false;
            buttonFullscreen.Text = "Exit fullscreen";
            isFullscreen = true;
        }


        // ###########################################################################################
        // Exit fullscreen.
        // ###########################################################################################

        private void FullscreenModeExit()
        {
            // Restore previous window state
            panelBehindTab.Dock = DockStyle.None;
            FormBorderStyle = formPreviousFormBorderStyle;
            WindowState = formPreviousWindowState;
            Bounds = previousBoundsForm;
            panelBehindTab.Bounds = previousBoundsPanelBehindTab;
            buttonFullscreen.Bounds = previousBoundsFullscreenButton;

            // Determine which tab should be repopulated with the previous maximized panel
            panelBehindTab.Controls.Remove(panelMain);
            splitContainerSchematics.Panel1.Controls.Add(panelMain);

            // Re-enable the filter
            textBoxFilterComponents.Enabled = true;
            textBoxFilterComponents.Focus();

            // Show again the tabs, and hide the fullscreen panel
            tabControl.Visible = true;
            buttonFullscreen.Text = "Fullscreen";
            isFullscreen = false;
        }


        // ###########################################################################################
        // Handle the fullscreen button click.
        // ###########################################################################################

        private void buttonFullscreen_Click(object sender, EventArgs e)
        {
            SuspendLayout();
            if (!isFullscreen)
            {
                FullscreenModeEnter();
            }
            else
            {
                FullscreenModeExit();
                ReadaptThumbnails();
            }

            ResumeLayout();

            // Reposition fullscreen button when in fullscreen (can only be done when UI is rendered)
            if (isFullscreen)
            {
                buttonFullscreen.Location = new Point(panelBehindTab.Width - buttonFullscreen.Width - 25, panelBehindTab.Height - 55);
                buttonFullscreen.BringToFront();
            }
        }


        // ###########################################################################################
        // Check if the "Fullscreen" button should be enabled or not.
        // ###########################################################################################

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl.SelectedTab.Text == "Schematics")
            {
                buttonFullscreen.Enabled = true;
            }
            else
            {
                buttonFullscreen.Enabled = false;
            }
        }


        // ###########################################################################################
        // Initialize the main/schematic image.
        // ###########################################################################################

        private void InitializeTabMain()
        {
            // Debug
#if DEBUG
            StackTrace stackTrace = new StackTrace();
            StackFrame callerFrame = stackTrace.GetFrame(1);
            MethodBase callerMethod = callerFrame.GetMethod();
            string callerName = callerMethod.Name;
            Debug.WriteLine("[InitializeTabMain] called from [" + callerName + "]");
#endif

            // Reset zoom-level
            zoomLevel = 1;

            // Dispose the image, if one already exists
            if (image != null)
            {
                image.Dispose();
                image = null;
            }

            // Load main image
            string filePath = DataPaths.Resolve(schematicSelectedFile);
            image = Image.FromFile(
                filePath
            );

            // Clear old controls
            RemoveClosePopupOnClickHandlers();
            panelMain.Controls.Clear();
            overlayComponentsTab.Clear();
            overlayComponentsTabOriginalSizes.Clear();
            overlayComponentsTabOriginalLocations.Clear();

            // Build the arrays
            CreateOverlayArraysToTab();

            // Create scrolling container
            panelZoom = new CustomPanel
            {
                AutoScroll = true,
                Dock = DockStyle.Fill
            };
            panelZoom.DoubleBuffered(true);
            panelMain.Controls.Add(panelZoom);

            // Create panelImage for the main picture
            panelImageMain = new Panel
            {
                Size = image.Size,
                BackgroundImage = image,
                BackgroundImageLayout = ImageLayout.Zoom,
                Dock = DockStyle.None
            };
            panelImageMain.DoubleBuffered(true);
            panelZoom.Controls.Add(panelImageMain);

            // Create overlay panel on top
            overlayPanel = new OverlayPanel
            {
                Bounds = panelImageMain.ClientRectangle
            };
            panelImageMain.Controls.Add(overlayPanel);
            overlayPanel.BringToFront();

            // Define events for the overlay for the components
            overlayPanel.OverlayClicked += OverlayPanel_OverlayClicked;
            overlayPanel.OverlayHoverChanged += OverlayPanel_OverlayHoverChanged;
            overlayPanel.OverlayPanelMouseDown += OverlayPanel_OverlayPanelMouseDown;
            overlayPanel.OverlayPanelMouseMove += OverlayPanel_OverlayPanelMouseMove;
            overlayPanel.OverlayPanelMouseUp += OverlayPanel_OverlayPanelMouseUp;

            // Define event for the overlay panel itself for "draw mode"
            overlayPanel.MouseDown += polylinesManagement.panelImageMain_MouseDown;
            overlayPanel.MouseMove += polylinesManagement.panelImageMain_MouseMove;
            overlayPanel.MouseUp += polylinesManagement.panelImageMain_MouseUp;
            overlayPanel.Paint += polylinesManagement.panelImageMain_Paint;

            // Top-left label: file name
            labelFile = new Label
            {
                Name = "labelFile",
                Text = ConvertStringToLabel(schematicSelectedName),
                AutoSize = true,
                BackColor = Color.Khaki,
                ForeColor = Color.Black,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Calibri", 11),
                Location = new Point(5, 5),
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            labelFile.DoubleBuffered(true);
            panelMain.Controls.Add(labelFile);
            labelFile.BringToFront();

            // Top-left label: region
            Color regionColorBack = Color.IndianRed;
            Color regionColorText = Color.White;
            if (selectedRegion == "NTSC")
            {
                regionColorBack = Color.LightSteelBlue;
                regionColorText = Color.Black;
            }
            int regionStartX = 5 + labelFile.PreferredWidth + 5;
            labelRegion = new Label
            {
                Name = "labelRegion",
                Text = selectedRegion,
                AutoSize = true,
                BackColor = regionColorBack,
                ForeColor = regionColorText,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Calibri", 11),
                Location = new Point(regionStartX, 5),
                Anchor = AnchorStyles.Top | AnchorStyles.Left,

            };
            labelRegion.DoubleBuffered(true);
            panelMain.Controls.Add(labelRegion);
            labelRegion.BringToFront();

            // Top-left label: hovered component name
            labelComponent = new Label
            {
                Name = "labelComponent",
                AutoSize = true,
                BackColor = Color.Khaki,
                ForeColor = Color.Black,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Calibri", 13, FontStyle.Bold),
                Location = new Point(5, 30),
                Visible = false,
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            labelComponent.DoubleBuffered(true);
            panelMain.Controls.Add(labelComponent);
            labelComponent.BringToFront();

            labelZoom = new Label
            {
                AutoSize = true,
                BackColor = Color.Moccasin,
                ForeColor = Color.Black,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Calibri", 11, FontStyle.Regular)
            };
            labelZoom.DoubleBuffered(true);
            labelZoom.Text = $"Z:{zoomLevel}";
            labelZoom.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            panelMain.Controls.Add(labelZoom);
            labelZoom.BringToFront();
            labelZoom.Location = new Point(panelMain.ClientSize.Width - labelZoom.PreferredWidth - 20, 5);

            // Finish up
            ResizeTabImage();
            panelZoom.CustomMouseWheel += panelZoom_MouseWheel;
            panelZoom.Resize += panelZoom_Resize;

            // THEN re-run the attach method:
            AttachClosePopupOnClick(this);
        }

        // ###########################################################################################
        // Create overlay arrays for the main/schematic image.
        // ###########################################################################################

        private void CreateOverlayArraysToTab()
        {
            var hw = classHardware.FirstOrDefault(h => h.Name == hardwareSelectedName);
            var bd = hw?.Boards.FirstOrDefault(b => b.Name == boardSelectedName);
            if (bd == null) return;

            foreach (BoardOverlays bo in bd.Files)
            {
                if (bo.Name != schematicSelectedName) continue;
                if (bo?.Components != null)
                {
                    foreach (var comp in bo.Components)
                    {
                        string componentDisplay = bd.Components.FirstOrDefault(cb => cb.Label == comp.Label)?.NameDisplay ?? "";
                        if (comp.Overlays == null) continue;

                        foreach (var ov in comp.Overlays)
                        {
                            PictureBox overlayPictureBox = new PictureBox
                            {
                                Name = comp.Label,
                                Location = new Point(ov.Bounds.X, ov.Bounds.Y),
                                Size = new Size(ov.Bounds.Width, ov.Bounds.Height),
                                Tag = bo.Name
                            };

                            overlayComponentsTab.Add(overlayPictureBox);
                            int idx = overlayComponentsTab.Count - 1;
                            overlayComponentsTabOriginalSizes[idx] = overlayPictureBox.Size;
                            overlayComponentsTabOriginalLocations[idx] = overlayPictureBox.Location;
                        }
                    }
                }
            }
        }


        // ###########################################################################################
        // Initialize the list of thumbnails.
        // ###########################################################################################

        private void InitializeThumbnails()
        {
            // Debug
#if DEBUG
            StackTrace stackTrace = new StackTrace();
            StackFrame callerFrame = stackTrace.GetFrame(1);
            MethodBase callerMethod = callerFrame.GetMethod();
            string callerName = callerMethod.Name;
            Debug.WriteLine("[InitializeList] called from [" + callerName + "]");
#endif

            // Gracefully dispose all controls
            DisposeAllControls(panelThumbnails);
            panelThumbnails.Controls.Clear();

            // Clear all lists
            overlayPanelsList.Clear();
            overlayListZoomFactors.Clear();

            // Find relevant schematic images to show here
            var hw = classHardware.FirstOrDefault(h => h.Name == hardwareSelectedName);
            var bd = hw?.Boards.FirstOrDefault(b => b.Name == boardSelectedName);
            if (bd == null) return;
            if (bd.Files == null) return;

            // Walkthrough each schematic image for this board
            foreach (BoardOverlays schematic in bd.Files)
            {
                string filename = Path.Combine(DataPaths.DataRoot, schematic.SchematicFileName);

                // Panel that will hold the label and the image
                Panel panelThumbnail = new Panel
                {
                    Name = schematic.Name + "_container",
                    BorderStyle = BorderStyle.None,
                };

                Label labelThumbnail = new Label
                {
                    Text = schematic.Name,
                    AutoSize = false,
                    BorderStyle = BorderStyle.FixedSingle,
                    BackColor = Color.Khaki,
                    ForeColor = Color.Black,
                    TextAlign = ContentAlignment.MiddleLeft,
                    Font = new Font("Calibri", 9),
                };

                PictureBox panelImage = new PictureBox
                {
                    Name = schematic.Name,
                    BackgroundImage = Image.FromFile(filename),
                    BackgroundImageLayout = ImageLayout.Zoom,
                    BorderStyle = BorderStyle.FixedSingle,
                    Dock = DockStyle.None,
                };

                // Overlay panel for the image that will ensure we can click anywhere on the image (to select it)
                OverlayPanel overlayPanel = new OverlayPanel
                {
                    Dock = DockStyle.Fill
                };

                // Add panels to each other
                panelThumbnails.Controls.Add(panelThumbnail);
                panelThumbnail.Controls.Add(labelThumbnail);
                panelThumbnail.Controls.Add(panelImage);
                panelImage.Controls.Add(overlayPanel);

                overlayPanelsList[schematic.Name] = overlayPanel;

                panelThumbnail.DoubleBuffered(true);
                labelThumbnail.DoubleBuffered(true);
                panelImage.DoubleBuffered(true);
                overlayPanel.DoubleBuffered(true);

                // Attach "MouseDown" and "Clicked" events to the overlay panel
                overlayPanel.OverlayPanelMouseDown += (s, e2) =>
                {
                    if (e2.Button == MouseButtons.Left)
                        ThumbnailImageClicked(panelImage);
                };
                overlayPanel.OverlayClicked += (s, e2) =>
                {
                    if (e2.MouseArgs.Button == MouseButtons.Left)
                        ThumbnailImageClicked(panelImage);
                };
                labelThumbnail.MouseDown += (s, e2) =>
                {
                    if (e2.Button == MouseButtons.Left)
                        ThumbnailImageClicked(panelImage);
                };
                labelThumbnail.Click += (s, e2) =>
                {
                    if (((MouseEventArgs)e2).Button == MouseButtons.Left)
                        ThumbnailImageClicked(panelImage);
                };
            }
            ReadaptThumbnails();
            DrawBorderInList();
        }


        // ###########################################################################################
        // Resize the thumbnails.
        // ###########################################################################################

        private void ReadaptThumbnails()
        {
            // Debug
#if DEBUG
            StackTrace stackTrace = new StackTrace();
            StackFrame callerFrame = stackTrace.GetFrame(1);
            MethodBase callerMethod = callerFrame.GetMethod();
            string callerName = callerMethod.Name;
            Debug.WriteLine("[ReadaptThumbnails] called from [" + callerName + "]");
#endif

            thumbnailsSameWidth = false;
            thumbnailsWidth = 0;
            int thumbnailsWidthOld = 0;

            // Suspend continues redrawing of thumbnail panel
            SuspendDrawing(panelThumbnails);

            for (int i = 0; i <= 40; i++) // "40" is unrealistic here, as it should never go higher than 4
            {
                thumbnailsWidthOld = thumbnailsWidth;
                if (!thumbnailsSameWidth)
                {
                    ReadaptThumbnails_Retry();
                }
                else
                {
                    break;
                }

                // if we end up in a race condition, then end redrawing when most of the thumbnail image is visible
                if (i >= 3 && thumbnailsWidth < thumbnailsWidthOld)
                {
                    Debug.WriteLine("Race condition in [ReadaptThumbnails]");
                    panelThumbnails.AutoScrollMinSize = new Size(0, panelThumbnails.ClientSize.Height + 1);
                    break;
                }
                else
                {
                    panelThumbnails.AutoScrollMinSize = Size.Empty;
                }
            }

            HighlightOverlays("list");

            ResumeDrawing(panelThumbnails);
        }

        private void ReadaptThumbnails_Retry()
        {

            int yPosition = 3; // first thumbnail position
            int padding = 1; // padding from left and right borders
            int verticalSpaceBetweenThumbnails = 1;
            int availableWidthForThumbnailContainer = panelThumbnails.ClientSize.Width - (padding * 2);
            int availableWidthForThumbnailElement = availableWidthForThumbnailContainer - (thumbnailSelectedBorderWidth * 2);

            // Check if we can end the loop for continues redrawing
            thumbnailsSameWidth = availableWidthForThumbnailContainer == thumbnailsWidth ? true : false;

            // Walkthrough all (parent) panels (a panel is a thumbnail container)
            foreach (Panel panelThumbnail in panelThumbnails.Controls.OfType<Panel>())
            {
                // Get the two elements inside the panel
                Label labelImage = panelThumbnail.Controls[0] as Label;
                PictureBox pictureBoxImage = panelThumbnail.Controls[1] as PictureBox;

                // Calculate height for the label
                Size textSize = TextRenderer.MeasureText(labelImage.Text, labelImage.Font);
                int labelHeight = textSize.Height + 2;

                // Calculate new height for the image, based on fixed width
                float aspectRatio = (float)pictureBoxImage.BackgroundImage.Height / pictureBoxImage.BackgroundImage.Width;
                int newImageHeight = (int)(availableWidthForThumbnailElement * aspectRatio);

                // Set location and new size for all three elements
                panelThumbnail.Location = new Point(padding, yPosition + panelThumbnails.AutoScrollPosition.Y); // "panelThumbnails.AutoScrollPosition.Y" will take into consideration, if there is a scrollbar visible, and then correctly position the thumbnail
                panelThumbnail.Size = new Size(availableWidthForThumbnailContainer, labelHeight + newImageHeight + (thumbnailSelectedBorderWidth * 2));
                labelImage.Location = new Point(thumbnailSelectedBorderWidth, thumbnailSelectedBorderWidth);
                labelImage.Size = new Size(availableWidthForThumbnailElement, labelHeight + thumbnailSelectedBorderWidth);
                pictureBoxImage.Location = new Point(thumbnailSelectedBorderWidth, labelHeight + thumbnailSelectedBorderWidth - 1); // -1 to avoid having a double-line
                pictureBoxImage.Size = new Size(availableWidthForThumbnailElement, newImageHeight + 1); // +1 to compensate for the above -1

                yPosition += panelThumbnail.Height + verticalSpaceBetweenThumbnails;

                float scaleFactor = (float)pictureBoxImage.Width / pictureBoxImage.BackgroundImage.Width;
                string key = panelThumbnail.Name.Replace("_container", "");

                // Update the aspect ratio used, in the overlay list
                overlayListZoomFactors[key] = scaleFactor;
                if (overlayPanelsList.ContainsKey(key))
                {
                    overlayPanelsList[key].Bounds = pictureBoxImage.ClientRectangle;
                }
            }
            thumbnailsWidth = availableWidthForThumbnailContainer;
        }


        // ###########################################################################################
        // Refresh the labels of the thumbnails (show if component is included), based on selected components.
        // ###########################################################################################

        private void RefreshThumbnailLabels()
        {
            var hw = classHardware.FirstOrDefault(h => h.Name == hardwareSelectedName);
            var bd = hw?.Boards.FirstOrDefault(b => b.Name == boardSelectedName);
            if (bd == null) return;
            if (bd.Files == null) return;

            // Convert the selected items to a "HashSet" for faster lookup
            var selectedComponents = new HashSet<string>(listBoxComponentsSelectedText);

            foreach (var file in bd.Files)
            {
                // Now search for the container panel (instead of panel with file.Name)
                var container = panelThumbnails.Controls
                    .OfType<Panel>()
                    .FirstOrDefault(p => p.Name == file.Name + "_container");
                if (container == null) continue;

                // Find the label inside the container panel.
                var labelListFile = container.Controls.OfType<Label>().FirstOrDefault();
                if (labelListFile == null) continue;

                bool hasSelectedComponent = selectedComponents.Any(selectedComponentText =>
                {
                    // Find component "label"
                    string componentLabel = bd.Components
                        .FirstOrDefault(cb => cb.Label == selectedComponentText)?.Label ?? "";

                    // Find component "bounds" for the label
                    var compBounds = file.Components.FirstOrDefault(c => c.Label == componentLabel);

                    return compBounds != null && compBounds.Overlays != null && compBounds.Overlays.Count > 0;
                });

                if (hasSelectedComponent)
                {
                    labelListFile.Text = "* " + ConvertStringToLabel(file.Name);
                    labelListFile.BackColor = labelImageHasElementsBgClr;
                    labelListFile.ForeColor = labelImageHasElementsTxtClr;
                }
                else
                {
                    labelListFile.Text = ConvertStringToLabel(file.Name);
                    labelListFile.BackColor = labelImageBgClr;
                    labelListFile.ForeColor = labelImageTxtClr;
                }
            }
        }


        // ###########################################################################################
        // Handle the click on a thumbnail image.
        // ###########################################################################################

        private void ThumbnailImageClicked(PictureBox pan)
        {
            // Debug
#if DEBUG
            StackTrace stackTrace = new StackTrace();
            StackFrame callerFrame = stackTrace.GetFrame(1);
            MethodBase callerMethod = callerFrame.GetMethod();
            string callerName = callerMethod.Name;
            Debug.WriteLine("[ThumbnailImageClicked] called from [" + callerName + "]");
#endif

            // Clear all current overlays
            overlayPanel.Overlays.Clear();

            schematicSelectedName = pan.Name;

            // Save the board specific selected thumbnail to configuration file
            string configKey = $"SelectedThumbnail|{hardwareSelectedName}|{boardSelectedName}";
            Configuration.SaveSetting(configKey, schematicSelectedName);

            textBox2.Text = ConvertStringToLabel(schematicSelectedName); // feedback info

            var hw = classHardware.FirstOrDefault(h => h.Name == hardwareSelectedName);
            var bd = hw?.Boards.FirstOrDefault(b => b.Name == boardSelectedName);
            if (bd != null)
            {
                var file = bd.Files.FirstOrDefault(f => f.Name == schematicSelectedName);
                if (file != null)
                {
                    schematicSelectedFile = file.SchematicFileName;
                    InitializeTabMain();  // load the selected image to "Main"
                }
            }

            // Ensure thumbnail border gets updated
            DrawBorderInList();

            // Make sure polylines for the new image are loaded
            if (!PolylinesManagement.imagePolylines.ContainsKey(schematicSelectedName))
            {
                PolylinesManagement.imagePolylines[schematicSelectedName] = new List<List<Point>>();
            }

            // Update the polylines reference to use the correct list for this schematic
            PolylinesManagement.polylines = PolylinesManagement.imagePolylines[schematicSelectedName];

            // Reset selection when changing images
            PolylinesManagement.selectedPolylineIndex = -1;
            PolylinesManagement.selectedMarker = (-1, -1);

            PopulatePolylineVisibilityPanel();
        }


        // ###########################################################################################
        // Initialize the component categories.
        // ###########################################################################################

        private void InitializeComponentCategories()
        {
            // Debug
#if DEBUG
            StackTrace stackTrace = new StackTrace();
            StackFrame callerFrame = stackTrace.GetFrame(1);
            MethodBase callerMethod = callerFrame.GetMethod();
            string callerName = callerMethod.Name;
            Debug.WriteLine("[InitializeComponentCategories] called from [" + callerName + "]");
#endif

            listBoxCategories.BeginUpdate(); // suspend redrawing this listBox while updating it
            listBoxCategories.Items.Clear();

            var foundHardware = classHardware.FirstOrDefault(h => h.Name == hardwareSelectedName);
            var foundBoard = foundHardware?.Boards.FirstOrDefault(b => b.Name == boardSelectedName);

            if (foundBoard?.Components != null)
            {
                // Use a "HashSet" to track added categories for faster lookups
                var addedTypes = new HashSet<string>();

                foreach (BoardComponents component in foundBoard.Components)
                {
                    if (!string.IsNullOrEmpty(component.Type) && addedTypes.Add(component.Type))
                    {
                        listBoxCategories.Items.Add(component.Type);
                    }
                }
            }
            AdjustListBoxCategoriesHeight();
            listBoxCategories.EndUpdate(); // resume redrawing of this specific listBox
        }

        private void AdjustListBoxCategoriesHeight()
        {
            int itemCount = listBoxCategories.Items.Count;
            if (itemCount == 0) return;
            int itemHeight = listBoxCategories.ItemHeight;
            int borderHeight = SystemInformation.BorderSize.Height * 2;
            int totalHeight = (itemHeight * itemCount) + borderHeight + 15;
            listBoxCategories.Height = totalHeight;

            labelComponents.Location = new Point(
                labelComponents.Location.X,
                listBoxCategories.Location.Y + listBoxCategories.Height + 10
            );

            listBoxComponents.Location = new Point(
                listBoxComponents.Location.X,
                labelComponents.Location.Y + labelComponents.Height + 5
            );

            // Calculate the height so listBoxComponents extends down to just above textBoxFilterComponents
            int bottomY = textBoxFilterComponents.Location.Y;
            int topY = listBoxComponents.Location.Y;
            listBoxComponents.Height = bottomY - topY - 5; // 5px padding, adjust as needed
        }


        // ###########################################################################################
        // Resize the main/schematic image.
        // ###########################################################################################

        private void ResizeTabImage()
        {
            if (image == null) return;

            int scrollbarVerticalWidth = SystemInformation.VerticalScrollBarWidth;
            int scrollbarHorizontalHeight = SystemInformation.HorizontalScrollBarHeight;

            int availableWidth = panelZoom.Width - scrollbarVerticalWidth;
            int availableHeight = panelZoom.Height - scrollbarHorizontalHeight;

            float xZoomFactor = (float)availableWidth / image.Width;
            float yZoomFactor = (float)availableHeight / image.Height;
            zoomFactor = Math.Min(xZoomFactor, yZoomFactor);

            panelImageMain.Size = new Size(
                (int)(image.Width * zoomFactor),
                (int)(image.Height * zoomFactor)
            );

            // Always enforce scrollbars to avoid the weird "first zoom-in" flickering
            panelZoom.AutoScroll = true;
            panelZoom.AutoScrollMinSize = new Size(panelZoom.Width + 1, panelZoom.Height + 1);

            PopulatePolylineVisibilityPanel();

            HighlightOverlays("tab");
        }


        // ###########################################################################################
        // Resize the tab image when the panel is resized by mousewheel or keys.
        // ###########################################################################################

        private void panelZoom_Resize(object sender, EventArgs e)
        {
            if (!isResizedByMouseWheel)
            {
                ResizeTabImage();
            }
            isResizedByMouseWheel = false;
        }

        private void panelZoom_MouseWheel(object sender, MouseEventArgs e)
        {
            PerformZoom(e.Delta, e.Location);
        }

        private void PerformZoom(int delta, Point clientPos)
        {
            if (panelZoom == null || panelImageMain == null || image == null) return;

            ControlUpdateHelper.BeginControlUpdate(panelZoom);
            try
            {
                float oldZoomFactor = zoomFactor;
                bool hasZoomChanged = false;

                // Zoom IN
                if (delta > 0) 
                {
                    if (zoomFactor * 1.5f <= 4.5f)
                    {
                        zoomFactor *= 1.5f;
                        hasZoomChanged = true;
                        zoomLevel++;
                    }
                }

                // Zoom OUT
                else
                {
                    // Only zoom out if the image is bigger than the container
                    if (zoomLevel > 1)
                    {
                        if (panelImageMain.Width > panelZoom.Width || panelImageMain.Height > panelZoom.Height)
                        {
                            zoomFactor /= 1.5f;
                            hasZoomChanged = true;
                            zoomLevel--;
                        }
                    } else
                    {
                        zoomLevel = 1;
                    }
                }

                if (hasZoomChanged)
                {
                    isResizedByMouseWheel = true;

                    // 2) Calculate new size
                    Size newSize = new Size(
                        (int)(image.Width * zoomFactor),
                        (int)(image.Height * zoomFactor)
                    );

                    // If cursor is outside panelZoom, default to center
                    if (!panelZoom.ClientRectangle.Contains(clientPos))
                    {
                        clientPos = new Point(panelZoom.ClientSize.Width / 2, panelZoom.ClientSize.Height / 2);
                    }

                    // 3) Figure out how to keep the same "point under mouse"
                    Point mousePosition = new Point(
                        clientPos.X - panelZoom.AutoScrollPosition.X,
                        clientPos.Y - panelZoom.AutoScrollPosition.Y
                    );

                    Point newScrollPosition = new Point(
                        (int)(mousePosition.X * (zoomFactor / oldZoomFactor)),
                        (int)(mousePosition.Y * (zoomFactor / oldZoomFactor))
                    );

                    // 4) Apply the new size
                    panelImageMain.Size = newSize;

                    // 5) Update the scroll position
                    panelZoom.AutoScrollPosition = new Point(
                        newScrollPosition.X - clientPos.X,
                        newScrollPosition.Y - clientPos.Y
                    );

                    // 6) Re-highlight overlays (so they scale properly)
                    HighlightOverlays("tab");

                    labelZoom.Text = $"Z:{zoomLevel}";
                }
            }
            finally
            {
                ControlUpdateHelper.EndControlUpdate(panelZoom);
            }
        }


        // ###########################################################################################
        // Highlight all overlays according to the selected components.
        // ###########################################################################################

        private void ShowOverlaysAccordingToComponentList()
        {
            HighlightOverlays("tab");
            HighlightOverlays("list");

            // Refresh the thumbnails, to show if thumbnail includes the selected components
            RefreshThumbnailLabels();
        }


        // ###########################################################################################
        // Handle click-event for component highlights - both left- and rightclick events.
        // ###########################################################################################

        private void OverlayPanel_OverlayClicked(object sender, OverlayClickedEventArgs e)
        {
            // First, check if we're clicking on a marker
            bool clickedOnMarker = false;
            int clickedPolylineIndex = -1;
            int clickedPointIndex = -1;

            if (e.MouseArgs.Button == MouseButtons.Left)
            {
                // Check if clicking on an existing marker
                for (int i = 0; i < PolylinesManagement.polylines.Count; i++)
                {
                    for (int j = 0; j < PolylinesManagement.polylines[i].Count; j++)
                    {
                        Point scaledMarker = new Point(
                            (int)(PolylinesManagement.polylines[i][j].X * zoomFactor),
                            (int)(PolylinesManagement.polylines[i][j].Y * zoomFactor)
                        );

                        // Check if the click is on a marker
                        const int proximityThreshold = 10;
                        int effectiveRadius = 5 + proximityThreshold; // 5 is MarkerRadius from PolylinesManagement

                        if (Math.Pow(e.MouseArgs.Location.X - scaledMarker.X, 2) +
                            Math.Pow(e.MouseArgs.Location.Y - scaledMarker.Y, 2) <=
                            Math.Pow(effectiveRadius, 2))
                        {
                            clickedOnMarker = true;
                            clickedPolylineIndex = i;
                            clickedPointIndex = j;
                            break;
                        }
                    }
                    if (clickedOnMarker) break;
                }

                // If not clicking on a marker, check if clicking on a line segment
                if (!clickedOnMarker)
                {
                    for (int i = 0; i < PolylinesManagement.polylines.Count; i++)
                    {
                        for (int j = 0; j < PolylinesManagement.polylines[i].Count - 1; j++)
                        {
                            Point scaledStart = new Point(
                                (int)(PolylinesManagement.polylines[i][j].X * zoomFactor),
                                (int)(PolylinesManagement.polylines[i][j].Y * zoomFactor)
                            );
                            Point scaledEnd = new Point(
                                (int)(PolylinesManagement.polylines[i][j + 1].X * zoomFactor),
                                (int)(PolylinesManagement.polylines[i][j + 1].Y * zoomFactor)
                            );

                            // Get closest point on line segment
                            float dx = scaledEnd.X - scaledStart.X;
                            float dy = scaledEnd.Y - scaledStart.Y;

                            if (dx == 0 && dy == 0) continue; // Skip if line is a point

                            float t = ((e.MouseArgs.Location.X - scaledStart.X) * dx +
                                      (e.MouseArgs.Location.Y - scaledStart.Y) * dy) /
                                      (dx * dx + dy * dy);
                            t = Math.Max(0, Math.Min(1, t)); // Clamp t to [0,1]

                            Point closestPoint = new Point(
                                (int)(scaledStart.X + t * dx),
                                (int)(scaledStart.Y + t * dy)
                            );

                            // Check if click is near the line
                            const int proximityThreshold = 5;
                            if (Math.Abs(e.MouseArgs.Location.X - closestPoint.X) <= proximityThreshold &&
                                Math.Abs(e.MouseArgs.Location.Y - closestPoint.Y) <= proximityThreshold)
                            {
                                clickedPolylineIndex = i;
                                clickedPointIndex = j;

                                // Insert new marker immediately
                                Point newPointUnscaled = new Point(
                                    (int)(closestPoint.X / zoomFactor),
                                    (int)(closestPoint.Y / zoomFactor)
                                );
                                PolylinesManagement.polylines[i].Insert(j + 1, newPointUnscaled);

                                // Select the new marker
                                PolylinesManagement.selectedPolylineIndex = i;
                                PolylinesManagement.selectedMarker = (i, j + 1);

                                // Redraw and save
                                Main.overlayPanel.Invalidate();
                                PolylinesManagement.SavePolylinesToConfig();
                                return;
                            }
                        }
                    }
                }
            }

            // If clicked on marker, forward the event to polylines management
            if (clickedOnMarker)
            {
                // Store which polyline and point was selected
                PolylinesManagement.selectedPolylineIndex = clickedPolylineIndex;
                PolylinesManagement.selectedMarker = (clickedPolylineIndex, clickedPointIndex);

                // Manually trigger a mouse down event on the overlay panel to initiate marker movement
                MouseEventArgs newArgs = new MouseEventArgs(
                    e.MouseArgs.Button,
                    e.MouseArgs.Clicks,
                    e.MouseArgs.X,
                    e.MouseArgs.Y,
                    e.MouseArgs.Delta
                );

                // Forward the event to the polylines management
                polylinesManagement.panelImageMain_MouseDown(overlayPanel, newArgs);

                // Make sure to redraw
                overlayPanel.Invalidate();
                return;
            }

            // Continue with regular component overlay handling
            string componentClickedLabel = e.OverlayInfo.ComponentLabel;

            // Find component "display name"
            var hw = classHardware.FirstOrDefault(h => h.Name == hardwareSelectedName);
            var bd = hw?.Boards.FirstOrDefault(b => b.Name == boardSelectedName);
            string componentLabel = bd.Components.FirstOrDefault(cb => cb.Label == componentClickedLabel)?.Label ?? "";
            //            string componentDisplay = bd.Components.FirstOrDefault(cb => cb.Label == componentClickedLabel)?.NameDisplay ?? "";
            var compEntry = bd.Components
                .FirstOrDefault(cb => cb.Label == componentClickedLabel &&
                                      !string.IsNullOrEmpty(Main.selectedRegion) &&
                                      string.Equals(cb.Region, Main.selectedRegion, StringComparison.OrdinalIgnoreCase))
                ?? bd.Components
                .FirstOrDefault(cb => cb.Label == componentClickedLabel &&
                                      string.IsNullOrEmpty(cb.Region));
            string componentDisplay = compEntry?.NameDisplay ?? "";


            // Left-mouse click (select component and show popup)
            if (e.MouseArgs.Button == MouseButtons.Left)
            {
                // 1) HIGHLIGHT the overlay
                if (!listBoxComponentsSelectedText.Contains(componentLabel))
                {
                    listBoxComponentsSelectedText.Add(componentLabel);
                }
                int index = listBoxComponents.Items.IndexOf(componentDisplay);
                if (index >= 0)
                {
                    listBoxComponents.SetSelected(index, true);
                }

                // 2) SHOW the form for the clicked component
                var hardware = classHardware.FirstOrDefault(h => h.Name == hardwareSelectedName);
                var board = hardware?.Boards.FirstOrDefault(b => b.Name == boardSelectedName);
                var comps = board?.Components.Where(c => c.Label == componentClickedLabel).ToList();
                if (comps != null && comps.Count > 0)
                {
                    ShowComponentPopup(comps);
                }
            }

            // Right-mouse click (toggle component selection)
            else if (e.MouseArgs.Button == MouseButtons.Right)
            {
                if (!listBoxComponentsSelectedText.Contains(componentLabel))
                {
                    listBoxComponentsSelectedText.Add(componentLabel);
                    int index = listBoxComponents.Items.IndexOf(componentDisplay);
                    if (index >= 0)
                    {
                        listBoxComponents.SetSelected(index, true);
                    }
                }
                else
                {
                    listBoxComponentsSelectedText.Remove(componentLabel);
                    int index = listBoxComponents.Items.IndexOf(componentDisplay);
                    if (index >= 0)
                    {
                        listBoxComponents.SetSelected(index, false);
                    }
                }
                ShowOverlaysAccordingToComponentList();
            }

            // Refresh the highlight overlays
            ShowOverlaysAccordingToComponentList();
        }


        // ###########################################################################################
        // Handle mouse-hover event for component highlights.
        // ###########################################################################################

        private void OverlayPanel_OverlayHoverChanged(object sender, OverlayHoverChangedEventArgs e)
        {
            if (labelComponent == null) return;

            // If hovering on a component
            if (e.IsHovering)
            {
                overlayPanel.Cursor = Cursors.Hand;
                var hw = classHardware.FirstOrDefault(h => h.Name == hardwareSelectedName);
                var bd = hw?.Boards.FirstOrDefault(b => b.Name == boardSelectedName);

                // Get the "display name" based on selected region
                var comp = bd?.Components
                    .FirstOrDefault(c => c.Label == e.OverlayInfo.ComponentLabel &&
                                         !string.IsNullOrEmpty(selectedRegion) &&
                                         string.Equals(c.Region, selectedRegion, StringComparison.OrdinalIgnoreCase))
                    ?? bd?.Components
                    .FirstOrDefault(c => c.Label == e.OverlayInfo.ComponentLabel &&
                                         string.IsNullOrEmpty(c.Region))
                    ?? bd?.Components
                    .FirstOrDefault(c => c.Label == e.OverlayInfo.ComponentLabel);

                if (comp != null)
                {
                    labelComponent.Text = ConvertStringToLabel(comp.NameDisplay);
                }
                else
                {
                    labelComponent.Text = ConvertStringToLabel(e.OverlayInfo.ComponentLabel);
                }

                if (comp != null)
                {
                    labelComponent.Text = ConvertStringToLabel(comp.NameDisplay);
                }
                else
                {
                    labelComponent.Text = ConvertStringToLabel(e.OverlayInfo.ComponentLabel);
                }
                labelComponent.Visible = true;
            }
            else
            {
                overlayPanel.Cursor = Cursors.Cross;
                labelComponent.Visible = false;
            }
        }


        // ###########################################################################################
        // Allow panning when right-clicking (and hold) directly on a component overlay.
        // ###########################################################################################

        private void OverlayPanel_OverlayPanelMouseDown(object sender, MouseEventArgs e)
        {
            // Right-click drag on empty space
            if (e.Button == MouseButtons.Right)
            {
                overlayPanelLastMousePos = e.Location;
            }
        }

        private void OverlayPanel_OverlayPanelMouseMove(object sender, MouseEventArgs e)
        {
            // If user is dragging with right-click
            if (e.Button == MouseButtons.Right && overlayPanelLastMousePos != Point.Empty)
            {
                int dx = e.X - overlayPanelLastMousePos.X;
                int dy = e.Y - overlayPanelLastMousePos.Y;
                panelZoom.AutoScrollPosition = new Point(
                    -panelZoom.AutoScrollPosition.X - dx,
                    -panelZoom.AutoScrollPosition.Y - dy
                );
            }
        }

        private void OverlayPanel_OverlayPanelMouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                overlayPanelLastMousePos = Point.Empty;
            }
        }


        // ###########################################################################################
        // Close any open component informtion popup.
        // ###########################################################################################

        // Add near your other fields
        private readonly Dictionary<Control, MouseEventHandler> _closePopupMouseDownHandlers =
            new Dictionary<Control, MouseEventHandler>();

        private void AttachClosePopupOnClick(Control parent)
        {
            if (parent == null) return;

            if (!_closePopupMouseDownHandlers.ContainsKey(parent))
            {
                MouseEventHandler handler = (s, e) =>
                {
                    CloseComponentPopup();
                };

                parent.MouseDown += handler;
                _closePopupMouseDownHandlers[parent] = handler;
            }

            foreach (Control child in parent.Controls)
            {
                AttachClosePopupOnClick(child);
            }
        }

        private void RemoveClosePopupOnClickHandlers()
        {
            foreach (var kvp in _closePopupMouseDownHandlers)
            {
                kvp.Key.MouseDown -= kvp.Value;
            }

            _closePopupMouseDownHandlers.Clear();
        }


        // ###########################################################################################
        // Show the component information popup.
        // ###########################################################################################

        private void ShowComponentPopup(List<BoardComponents> comps)
        {
            componentInfoPopup = new FormComponent(comps, this);
            componentInfoPopup.Show(this);
            componentInfoPopup.BringToFront();
        }


        // ###########################################################################################
        // Dispose all controls in a parent control.
        // Proper handling of memory.
        // ###########################################################################################

        void DisposeAllControls(Control parent)
        {
            for (int i = parent.Controls.Count - 1; i >= 0; i--)
            {
                Control child = parent.Controls[i];
                DisposeAllControls(child);
                child.Dispose();
            }
        }


        // ###########################################################################################
        // "Clear" buttons
        // ###########################################################################################

        private void buttonClear_Click(object sender, EventArgs e)
        {
            ClearEverything();
        }

        private void ClearEverything()
        {
            listBoxComponents.ClearSelected();
            listBoxComponentsSelectedText.Clear();
            textBoxFilterComponents.Text = "";
        }


        // ###########################################################################################
        // "Mark all" button
        // ###########################################################################################

        private void buttonAll_Click(object sender, EventArgs e)
        {
            listBoxComponents.SelectedIndexChanged -= listBoxComponents_SelectedIndexChanged;

            listBoxComponents.BeginUpdate(); // suspend redrawing this listBox while updating it

            for (int i = 0; i < listBoxComponents.Items.Count; i++)
            {
                listBoxComponents.SetSelected(i, true);
            }

            listBoxComponents.EndUpdate(); // resume redrawing of this specific listBox

            listBoxComponents.SelectedIndexChanged += listBoxComponents_SelectedIndexChanged;
            UpdateComponentSelection();
        }


        // ###########################################################################################
        // "Hardware" combobox change.
        // ###########################################################################################       

        private async void comboBoxHardware_SelectedIndexChanged(object sender, EventArgs e)
        {
            ClearEverything();

            comboBoxBoard.Items.Clear();
            hardwareSelectedName = comboBoxHardware.SelectedItem.ToString();
            textBox5.Text = ConvertStringToLabel(hardwareSelectedName); // feedback info

            var hw = classHardware.FirstOrDefault(h => h.Name == hardwareSelectedName);
            if (hw != null)
            {

                string defaultBoardSelectedName = hw.Boards.FirstOrDefault()?.Name ?? "";

                // Load the saved "selected board" from the configuration file
                string configKey = $"SelectedBoard|{hardwareSelectedName}";
                boardSelectedName = Configuration.GetSetting(configKey, defaultBoardSelectedName);

                foreach (var board in hw.Boards)
                {
                    comboBoxBoard.Items.Add(board.Name);
                }

                // Find the index of "boardSelectedName" in the combobox
                int boardIndex = comboBoxBoard.Items.IndexOf(boardSelectedName);

                // Set the selected board
                if (boardIndex != -1)
                {
                    comboBoxBoard.SelectedIndex = boardIndex;
                }
                else
                {
                    // If not found, select the first item
                    comboBoxBoard.SelectedIndex = 0;
                }
            }
        }


        // ###########################################################################################
        // "Board" combobox change.
        // ###########################################################################################

        private void comboBoxBoard_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetupNewBoard();
        }


        // ###########################################################################################
        // Save and resize thumbnails when the splitter position changes.
        // ###########################################################################################

        private void SplitContainer1_SplitterMoved(object sender, SplitterEventArgs e)
        {
            // Save the board specific splitter position to configuration file
            string configKey = $"SplitterPosition|{windowState}|{hardwareSelectedName}|{boardSelectedName}";
            Configuration.SaveSetting(configKey, splitContainerSchematics.SplitterDistance.ToString());

            ReadaptThumbnails();
        }


        // ###########################################################################################
        // Paint event for "SplitContainer".
        // ###########################################################################################

        private void SplitContainer1_Paint(object sender, PaintEventArgs e)
        {
            SplitContainer splitContainer = sender as SplitContainer;
            if (splitContainer != null)
            {
                // Draw a custom line in the middle of the splitter
                int splitterWidth = splitContainer.SplitterWidth;
                int halfWidth = splitterWidth / 2;
                int x = splitContainer.SplitterDistance + halfWidth;
                int y1 = splitContainer.Panel1.ClientRectangle.Top;
                int y2 = splitContainer.Panel1.ClientRectangle.Bottom;

                using (Pen pen = new Pen(Color.Gray, 2))
                {
                    e.Graphics.DrawLine(pen, x, y1, x, y2);
                }
            }
        }


        // ###########################################################################################
        // Keyboard handling.
        // ###########################################################################################

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (tabControl.SelectedTab.Text == "Schematics")
            {
                // Handle zoom via +/-
                // Supports main keyboard (+/-) and numpad (Add/Subtract)
                if (checkBoxKeyboardZoom.Checked)
                {
                    Keys keyCode = keyData & Keys.KeyCode;
                    if (keyCode == Keys.Oemplus || keyCode == Keys.Add)
                    {
                        var clientPos = panelZoom != null ? panelZoom.PointToClient(Cursor.Position) : Point.Empty;
                        PerformZoom(+120, clientPos); // +120 mimics one wheel notch up
                        return true;
                    }
                    else if (keyCode == Keys.OemMinus || keyCode == Keys.Subtract)
                    {
                        var clientPos = panelZoom != null ? panelZoom.PointToClient(Cursor.Position) : Point.Empty;
                        PerformZoom(-120, clientPos); // -120 mimics one wheel notch down
                        return true;
                    }
                }

                // Fullscreen mode toggle
                if (keyData == Keys.F11)
                {
                    buttonFullscreen_Click(null, null);
                    return true;
                }
                else if (keyData == Keys.Escape && isFullscreen)
                {
                    buttonFullscreen_Click(null, null);
                    return true;
                }

                // Move selected polyline with arrow keys
                if (PolylinesManagement.selectedPolylineIndex != -1)
                {
                    int dx = 0, dy = 0;

                    if (keyData == Keys.Up)
                        dy = -1;
                    else if (keyData == Keys.Down)
                        dy = 1;
                    else if (keyData == Keys.Left)
                        dx = -1;
                    else if (keyData == Keys.Right)
                        dx = 1;

                    if (dx != 0 || dy != 0)
                    {
                        PolylinesManagement.MovePolyline(PolylinesManagement.selectedPolylineIndex, dx, dy);
                        overlayPanel.Invalidate();
                        PolylinesManagement.SavePolylinesToConfig();
                        return true;
                    }
                }
            }

            // Toggle "checkBoxBlink" with SPACE key
            if (keyData == Keys.Enter && tabControl.SelectedTab.Text != "Feedback")
            {
                checkBoxBlink.Checked = !checkBoxBlink.Checked;
                return true;
            }

            // Trigger the "buttonAll_Click" event
            if (keyData == (Keys.Alt | Keys.A))
            {
                buttonAll_Click(null, EventArgs.Empty);
                return true;
            }

            // Trigger the "buttonClear_Click" event
            if (keyData == (Keys.Alt | Keys.C))
            {
                buttonClear_Click(null, EventArgs.Empty);
                return true;
            }

            return base.ProcessCmdKey(ref msg, keyData);
        }


        // ###########################################################################################
        // Post feedback to server.
        // ###########################################################################################

        private void buttonSendFeedback_Click(object sender, EventArgs e)
        {
            string email = textBoxEmail.Text;
            string feedback = ConvertStringToLabel(textBoxFeedback.Text);

            // Validate the email address
            bool isValidEmail = true;
            if (email.Length > 0)
            {
                isValidEmail = IsValidEmail(email);
            }

            // Proceed if the email address is empty or valid
            if (isValidEmail)
            {
                var foundHardware = classHardware.FirstOrDefault(h => h.Name == hardwareSelectedName);
                var foundBoard = foundHardware?.Boards.FirstOrDefault(b => b.Name == boardSelectedName);
                string boardFile = foundBoard?.DataFile;
                string excelFilePath = Path.Combine(DataPaths.DataRoot, boardFile);

                try
                {
                    using (WebClient webClient = new WebClient())
                    {
                        ServicePointManager.Expect100Continue = true;
                        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls13;
                        webClient.Headers.Add("user-agent", "CRT " + versionThis);

                        // Build the data to send
                        var data = new NameValueCollection
                        {
                            { "version", versionThis },
                            { "hardware", hardwareSelectedName },
                            { "board", boardSelectedName },
                            { "schematic", schematicSelectedName },
                            { "filename", boardFile },
                            { "email", email },
                            { "feedback", feedback }
                        };

                        // Attach the binary Excel file if the checkbox is checked
                        if (checkBoxAttachExcel.Checked)
                        {
                            if (!IsFileLocked(excelFilePath))
                            {
                                byte[] fileBytes = File.ReadAllBytes(excelFilePath);
                                string fileBase64 = Convert.ToBase64String(fileBytes);
                                data["attachment"] = fileBase64;
                            }
                            else
                            {
                                MessageBox.Show("The Excel file currently has an exclusive lock, and cannot be read. Please close any application that might be using it and try again.",
                                    "ERROR: Cannot read data file",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Error);
                            }
                        }

                        // Ensure the temporary "UserData" file is attached
                        string userData = ExtractUserDataForSelectedBoard();
                        if (!string.IsNullOrEmpty(userData))
                        {
                            // Save user data to a temporary file
                            string tempUserDataFile = Path.Combine(Path.GetTempPath(), "UserData.txt");
                            File.WriteAllText(tempUserDataFile, userData);

                            // Attach the user data file
                            byte[] userDataBytes = File.ReadAllBytes(tempUserDataFile);
                            string userDataBase64 = Convert.ToBase64String(userDataBytes);
                            data["userDataAttachment"] = userDataBase64;

                            // Clean up the temporary file
                            File.Delete(tempUserDataFile);
                        }

                        // Send it to the server
                        var response = webClient.UploadValues(crtPage + crtPageFeedback, "POST", data);
                        string resultFromServer = Encoding.UTF8.GetString(response);
                        if (resultFromServer == "Success")
                        {
                            if (email.Length > 0)
                            {
                                string txt = "Feedback sent. Please allow for some time, if any response is needed.";
                                MessageBox.Show(txt,
                                    "OK: Feedback sent",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Information);
                            }
                            else
                            {
                                string txt = "Feedback sent. No response will be given, as you did not specify an email address.";
                                MessageBox.Show(txt,
                                    "OK: Feedback sent",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Information);
                            }
                        }
                        else
                        {
                            string txt = "No feedback sent - did you fill in some text or attached the Excel data file?";
                            MessageBox.Show(txt,
                                "ERROR: Feedback not sent",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
                        }
                    }
                }
                catch (WebException ex)
                {
                    MessageBox.Show("CRT cannot submit the feedback right now, please retry later. If the issue persists, then you can connect directly with the developer at [dennis@commodore-repair-toolbox.dk]." + Environment.NewLine + Environment.NewLine + "The exact received HTTP error is:" + Environment.NewLine + Environment.NewLine + ex.Message,
                        "ERROR: Cannot connect with server",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
            else
            {
                string txt = "Invalid email address [" + email + "].";
                MessageBox.Show(txt,
                    "ERROR: Feedback not sent",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }


        // ###########################################################################################
        // Extract user data for the selected board.
        // Will be attached as a file to the feedback email.
        // ###########################################################################################

        private string ExtractUserDataForSelectedBoard()
        {
            var userData = new StringBuilder();
            string baseKey = $"UserData|{hardwareSelectedName}|{boardSelectedName}|";

            // Iterate through all configuration keys
            foreach (var key in Configuration.GetAllKeys())
            {
                if (key.StartsWith(baseKey))
                {
                    string value = Configuration.GetSetting(key, "");
                    if (!string.IsNullOrEmpty(value))
                    {
                        // Remove only the "UserData|" prefix
                        string trimmedKey = key.Substring("UserData|".Length);
                        userData.AppendLine($"{trimmedKey}={value}");
                    }
                }
            }

            return userData.ToString();
        }


        // ###########################################################################################
        // Check if the file is exclusively locked by another application.
        // ###########################################################################################

        private static bool IsFileLocked(string filePath)
        {
            if (File.Exists(filePath))
            {
                try
                {
                    using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.None))
                    {
                        stream.Close();
                    }
                }
                catch (IOException)
                {
                    return true;
                }
            }
            return false;
        }


        // ###########################################################################################
        // Check if the email address syntax is valid.
        // ###########################################################################################

        private static bool IsValidEmail(string email)
        {
            if (string.IsNullOrWhiteSpace(email))
                return false;

            try
            {
                var addr = new System.Net.Mail.MailAddress(email);

                // Ensure the original email matches exactly and contains a valid domain and TLD
                return addr.Address == email &&
                       Regex.IsMatch(email, @"^[^@\s]+@[^@\s]+\.[a-zA-Z]{2,}$");
            }
            catch
            {
                return false;
            }
        }


        // ###########################################################################################
        // Populate the filename of the attached Excel file, to the feedback tab UI.
        // ###########################################################################################

        private void checkBoxAttachExcel_CheckedChanged(object sender, EventArgs e)
        {
            string txtForField = Path.GetFileName(boardSelectedFilename);
            textBox6.Text = checkBoxAttachExcel.Checked ? txtForField : "";
        }


        // ###########################################################################################
        // Save the "Always update data" preference to the configuration file.
        // When checked, every application launch will synchronize data files from the online source.
        // ###########################################################################################

        private void checkBoxAlwaysUpdateData_CheckedChanged(object sender, EventArgs e)
        {
            Configuration.SaveSetting("AlwaysUpdateData", this.checkBoxAlwaysUpdateData.Checked.ToString());
        }


        // ###########################################################################################
        // Define a custom method to suspend and resume drawing on a control.
        // Should be better than "SuspendLayout" and "ResumeLayout".
        // ###########################################################################################

        private const int WM_SETREDRAW = 0x000B;

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessage(IntPtr hWnd, int msg, bool wParam, int lParam);

        public static void SuspendDrawing(Control control)
        {
            SendMessage(control.Handle, WM_SETREDRAW, false, 0);
        }

        public static void ResumeDrawing(Control control)
        {
            SendMessage(control.Handle, WM_SETREDRAW, true, 0);
            control.Refresh();
        }


        // ###########################################################################################
        // Save the 3 checkboxes in "Labels visible".
        // ###########################################################################################

        private void checkBoxVisibleLabels_CheckedChanged(object sender, EventArgs e)
        {
            // Save the state of the checkboxes
            Configuration.SaveSetting("ShowLabel", checkBox1.Checked.ToString());
            Configuration.SaveSetting("ShowTechnicalName", checkBox2.Checked.ToString());
            Configuration.SaveSetting("ShowFriendlyName", checkBox3.Checked.ToString());

            // Refresh the overlays
            HighlightOverlays("tab");
        }


        // ###########################################################################################
        // Minimize and maximize the "Labels visible" panel.
        // ###########################################################################################

        private void TogglePanelLabelsVisibility_Click(object sender, EventArgs e)
        {
            int currentHeight = panelLabelsVisible.Height;
            int newHeight = 0;

            if (currentHeight == 95)
            {
                newHeight = 26;
            }
            else
            {
                newHeight = 95;
            }

            panelLabelsVisible.Height = newHeight;

            Configuration.SaveSetting("ShowLabelsHeight", newHeight.ToString());

            // Reposition the panel
            int visibleHeight = panelZoom.ClientRectangle.Height;
            panelLabelsVisible.Location = new Point(0, visibleHeight - panelLabelsVisible.Height - 2); // bottom-left corner of the visible area
        }


        // ###########################################################################################
        // At application launch, then immediately enable drawing of polylines.
        // ###########################################################################################

        private void StartDrawingPolylines()
        {
            // Ensure the current image has an entry in the dictionary
            if (!PolylinesManagement.imagePolylines.ContainsKey(Main.schematicSelectedName))
            {
                PolylinesManagement.imagePolylines[Main.schematicSelectedName] = new List<List<Point>>();
            }
        }


        // ###########################################################################################
        // Clicking color button in "Traces visible" panel.
        // ###########################################################################################

        private void buttonColorPolyline_Click(object sender, EventArgs e)
        {
            if (colorDialog1.ShowDialog() == DialogResult.OK)
            {
                // If we should only choose the color for the next new polyline
                if (PolylinesManagement.selectedPolylineIndex == -1)
                {
                    PolylinesManagement.LastSelectedPolylineColor = colorDialog1.Color;
                    return;
                }

                // Use the composite key to set the color
                var key = (schematicSelectedName, PolylinesManagement.selectedPolylineIndex);
                PolylinesManagement.polylineColors[key] = colorDialog1.Color;

                // Update the last selected color
                PolylinesManagement.LastSelectedPolylineColor = colorDialog1.Color;

                // Redraw the panel to apply the new color
                overlayPanel.Invalidate();

                // Update the visibility panel and counters
                PopulatePolylineVisibilityPanel();
            }

            PolylinesManagement.SavePolylinesToConfig();
        }



        // ###########################################################################################
        // Update the "Traces visible" panel.
        // Will completely recreate all panel elements.
        // Will add/remove new colors and show/hide panel.
        // ###########################################################################################

        public void PopulatePolylineVisibilityPanel()
        {
            panelZoom.Invalidate();

            SuspendDrawing(panel1);

            panel1.Controls.Clear();

            // Get the polyline colors for the selected schematic
            var relevantColors = PolylinesManagement.polylineColors
                 .Where(kvp => kvp.Key.ImageName == schematicSelectedName)
                 .GroupBy(kvp => kvp.Value.ToArgb()) // Group by ARGB value
                 .ToDictionary(group => Color.FromArgb(group.Key), group => group.Count()); // Convert back to Color and count

            foreach (var color in relevantColors.Keys)
            {
                if (!PolylinesManagement.CheckboxStates.ContainsKey(color))
                {
                    PolylinesManagement.CheckboxStates[color] = true; // Default to visible
                }
            }

            int yOffset = 0; // start position for the first element

            // Create a label to display the headline
            Label labelHeadline = new Label
            {
                Text = "Traces visible",
                AutoSize = true,
                Location = new Point(1, yOffset + 4),
                Font = new Font("Calibri", 10, FontStyle.Regular)
            };

            // Create a button to toggle visibility of panel
            Button buttonToggleTracesVisibility = new Button
            {
                Text = "M",
                Size = new Size(22, 23),
                BackColor = Color.LightGray,
            };
            buttonToggleTracesVisibility.Location = new Point(panel1.Width - buttonToggleTracesVisibility.Width - 2, yOffset);

            panel1.Controls.Add(labelHeadline);
            panel1.Controls.Add(buttonToggleTracesVisibility);
            panel1.Controls.Add(buttonTraceColor);
            panel1.Controls.Add(buttonTracesDelete);

            yOffset += labelHeadline.Height + 9;

            foreach (var colorEntry in relevantColors)
            {
                var color = colorEntry.Key;
                var count = colorEntry.Value;

                // Create the checkbox
                CheckBox checkBox = new CheckBox
                {
                    Text = "", // No text, as the color is visualized
                    Checked = false,
                    Tag = color,
                    AutoSize = true,
                    Location = new Point(5, yOffset) // Position next to the color panel
                };

                // Attach an event handler to toggle visibility
                checkBox.CheckedChanged += CheckBox_CheckedChanged;

                bool isChecked = PolylinesManagement.CheckboxStates.ContainsKey(color) ? PolylinesManagement.CheckboxStates[color] : true;
                checkBox.Checked = isChecked;

                // Create a panel to display the color
                Panel colorPanel = new Panel
                {
                    Size = new Size(50, 13), // Small square to represent the color
                    BackColor = color,
                    BorderStyle = BorderStyle.FixedSingle,
                    Location = new Point(25, yOffset) // Align with the checkbox
                };

                // Attach an event handler to the panel
                colorPanel.Click += ColorPanel_Click;

                // Create a label to display the counter
                Label counterLabel = new Label
                {
                    Text = $"({count})", // Show the count
                    AutoSize = true,
                    Location = new Point(80, yOffset), // Position after the color panel
                    Font = new Font("Calibri", 9, FontStyle.Regular)
                };
                counterLabel.Location = new Point(colorPanel.Right + 2, yOffset - 1);

                // Add the color panel, checkbox, and counter label to the manual panel
                panel1.Controls.Add(colorPanel);
                panel1.Controls.Add(checkBox);
                panel1.Controls.Add(counterLabel);

                yOffset += Math.Max(colorPanel.Height, checkBox.Height) + 5; // Adjust spacing
            }

            int buttonSize = (panel1.Width - 13) / 2;
            buttonTraceColor.Size = new Size(buttonSize, 23);
            buttonTracesDelete.Size = new Size(buttonSize, 23);
            buttonTraceColor.Location = new Point(5, yOffset);
            buttonTracesDelete.Location = new Point(panel1.Width - buttonTracesDelete.Width - 5, yOffset);

            // Add the label panel to "panelMain"
            panelMain.Controls.Add(panel1);

            // Attach the event handler for the button
            panel1.Click -= TogglePanelTracesVisibility_Click;
            panel1.Click += TogglePanelTracesVisibility_Click;
            labelHeadline.Click -= TogglePanelTracesVisibility_Click;
            labelHeadline.Click += TogglePanelTracesVisibility_Click;
            buttonToggleTracesVisibility.Click -= TogglePanelTracesVisibility_Click;
            buttonToggleTracesVisibility.Click += TogglePanelTracesVisibility_Click;

            if (isPanelTracesVisible)
            {
                panel1.Height = yOffset + buttonTraceColor.Height + 5;
                panelTracesVisibleHeight = panel1.Height;
                Configuration.SaveSetting("ShowTracesHeight", panelTracesVisibleHeight.ToString());
            }
            else
            {
                // "26" pixels equals it is minimized
                panel1.Height = 26;
            }

            int visibleHeight = panelZoom.ClientRectangle.Height;
            int visibleWidth = panelZoom.ClientRectangle.Width;
            panel1.Location = new Point(visibleWidth - panel1.Width - 2, visibleHeight - panel1.Height - 2); // bottom-left corner of the visible area

            ResumeDrawing(panel1);

            if (relevantColors.Count == 0)
            {
                panel1.Visible = false;
            }
            else
            {
                panel1.Visible = true;
                panel1.BringToFront();
            }

        }


        // ###########################################################################################
        // Checkbox event handler for toggling polyline visibility.
        // ###########################################################################################

        private void CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            var cb = sender as CheckBox;
            if (cb != null && cb.Tag is Color selectedColor)
            {
                polylinesManagement.TogglePolylineVisibility(selectedColor, cb.Checked);
            }
        }


        // ###########################################################################################
        // Clicking the color panel/bar in the "Traces visible" panel.
        // ###########################################################################################

        private void ColorPanel_Click(object sender, EventArgs e)
        {
            var panel = sender as Panel;
            if (panel != null && panel.BackColor is Color selectedColor)
            {
                // Toggle the visibility state
                bool isVisible = PolylinesManagement.CheckboxStates.ContainsKey(selectedColor)
                                 && PolylinesManagement.CheckboxStates[selectedColor];
                PolylinesManagement.CheckboxStates[selectedColor] = !isVisible;

                // Update the visibility of the polyline
                polylinesManagement.TogglePolylineVisibility(selectedColor, !isVisible);

                // Find the corresponding checkbox for the color
                foreach (Control control in panel1.Controls)
                {
                    if (control is CheckBox checkBox && checkBox.Tag is Color color && color == selectedColor)
                    {
                        // Set the checkbox state to match the new visibility state
                        checkBox.Checked = !isVisible;
                        break;
                    }
                }
            }
        }


        // ###########################################################################################
        // Minimize and maximize the "Traces visible" panel.
        // ###########################################################################################

        private void TogglePanelTracesVisibility_Click(object sender, EventArgs e)
        {
            int currentHeight = panel1.Height;

            if (currentHeight > 26)
            {
                int newHeight = 26;
                panel1.Height = newHeight;

                // Reposition the panel
                int visibleHeight = panelZoom.ClientRectangle.Height;
                int visibleWidth = panelZoom.ClientRectangle.Width;
                panel1.Location = new Point(visibleWidth - panel1.Width - 2, visibleHeight - panel1.Height - 2); // bottom-left corner of the visible area

                isPanelTracesVisible = false;
            }
            else
            {
                isPanelTracesVisible = true;
                PopulatePolylineVisibilityPanel();
            }

            Configuration.SaveSetting("ShowTraces", isPanelTracesVisible.ToString());
        }


        // ###########################################################################################
        // "Delete all traces" button.
        // Will delete only for the selected hardware and board.
        // ###########################################################################################

        private void buttonTracesDelete_Click(object sender, EventArgs e)
        {
            // Confirm deletion
            var confirmResult = MessageBox.Show(
                "Are you sure you want to delete all traces for the selected hardware and board?",
                "Confirm deletion of all traces",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning
            );

            if (confirmResult == DialogResult.Yes)
            {
                // Clear traces for the selected hardware and board
                var selectedHardware = classHardware.FirstOrDefault(h => h.Name == hardwareSelectedName);
                var selectedBoard = selectedHardware?.Boards.FirstOrDefault(b => b.Name == boardSelectedName);
                if (selectedBoard != null)
                {
                    PolylinesManagement.ClearTracesForBoard(selectedBoard);
                    PopulatePolylineVisibilityPanel();
                }
            }
        }


        // ###########################################################################################
        // Button click for updating all (local) files with newest content from online source.
        // Write a configuration file parameter that states that the files should be updated
        // at next launch.
        // ###########################################################################################

        private void button2_Click(object sender, EventArgs e)
        {
            var confirmResult = MessageBox.Show(
                "Please confirm that you will overwrite ALL \"CRT\" Excel data files and images!?\r\n\r\nData will be updated at next application launch.",
                "Data update overwrite confirmation",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning
            );

            if (confirmResult != DialogResult.Yes)
            {
                return;
            }

            // Write a configuration file parameter that states that the files will be updated at next launch
            Configuration.SaveSetting("UpdateDataAtNextLaunch", "True");

            button2.Text = "Will update data at next launch";
            button2.Enabled = false;
        }


        private void checkFilesFromSource()
        {

            // Fetch the list of files and checksums from the online source
            List<DataUpdate> checksumFromOnline;

            try
            {
                using (var webClient = new WebClient())
                {
                    ServicePointManager.Expect100Continue = true;
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls13;
                    webClient.Headers.Add("user-agent", "CRT " + versionThis);

                    string json = webClient.DownloadString("https://commodore-repair-toolbox.dk/auto-data/dataChecksums.json");
                    checksumFromOnline = DataUpdate.LoadFromJson(json);
                }
                DebugOutput("INFO: Fetched checksum list of [" + checksumFromOnline.Count + "] files from online source");

                List<LocalFiles> checksumFromLocal = GetAllReferencedLocalFiles();
                DebugOutput("INFO: Calculated checksum list of [" + checksumFromLocal.Count + "] files from local storage");

                // Find files present online but missing locally
                var missingLocal = checksumFromOnline
                    .Where(online => !checksumFromLocal.Any(local =>
                        string.Equals(local.File, online.File, StringComparison.OrdinalIgnoreCase)))
                    .ToList();

                // Find files present in both lists but with different checksums
                var differingChecksums = checksumFromOnline
                    .Join(
                        checksumFromLocal,
                        online => online.File,
                        local => local.File,
                        (online, local) => new { File = online.File, OnlineChecksum = online.Checksum, LocalChecksum = local.Checksum }
                    )
                    .Where(x => !string.Equals(x.OnlineChecksum, x.LocalChecksum, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                // Combine missingLocal and differingChecksums to get the list of files to transfer from online to local
                // Only include files that are present in the online list
                var filesToTransfer = missingLocal
                    .Select(f => f.File)
                    .Concat(differingChecksums.Select(f => f.File))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Where(file => checksumFromOnline.Any(online => string.Equals(online.File, file, StringComparison.OrdinalIgnoreCase)))
                    .ToList();

                // ---

                if (filesToTransfer.Count > 0)
                {
                    foreach (var file in filesToTransfer)
                    {
                        Debug.WriteLine($"INFO: File to transfer: {file}");
                    }

                    // Only show the data label, if there is no newer version available (this is more important)
                    if (!label13.Visible)
                    {
                        string newText = "Newer data available; update from \"Configuration\" tab";
                        Size textSize = TextRenderer.MeasureText(newText, label13.Font);
                        Action updateLabel = () =>
                        {
                            label13.Text = newText;
                            label13.Width = textSize.Width + 2;
                            label13.Visible = true;
                            label13.Location = new Point(panelBehindTab.Width - label13.Width - 2, 3);
                        };
                        if (label13.InvokeRequired)
                        {
                            label13.Invoke(updateLabel);
                        }
                        else
                        {
                            updateLabel();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                DebugOutput("EXCEPTION raised for fetching JSON file catalogue from online source:");
                DebugOutput(ex.ToString());
            }
        }


        // ###########################################################################################
        // Synchronize all local data files with the online source.
        // When showCompletionDialog is true, a MessageBox is shown after completion.
        // When false (e.g. after auto-update), progress is shown on the splash screen only.
        // ###########################################################################################

        public static void SyncFilesFromSource(bool showCompletionDialog = true)
        {
            _syncPerformedAtLaunch = true;

            // Ensure version string is available (may not be set yet if called before the constructor)
            if (string.IsNullOrEmpty(versionThis))
            {
                versionThis = GetAssemblyVersion();
            }

            // Fetch the list of files and checksums from the online source
            List<DataUpdate> checksumFromOnline;

            try
            {
                Splashscreen.Current?.UpdateStatus("Fetching file catalog from server...");

                using (var webClient = new WebClient())
                {
                    ServicePointManager.Expect100Continue = true;
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls13;
                    webClient.Headers.Add("user-agent", "CRT " + versionThis);

                    string json = webClient.DownloadString("https://commodore-repair-toolbox.dk/auto-data/dataChecksums.json");
                    checksumFromOnline = DataUpdate.LoadFromJson(json);
                }
                DebugOutput("INFO: Fetched checksum list of [" + checksumFromOnline.Count + "] files from online source");

                Splashscreen.Current?.UpdateStatus("Calculating local file checksums...");

                List<LocalFiles> checksumFromLocal = GetAllReferencedLocalFiles();
                DebugOutput("INFO: Calculated checksum list of [" + checksumFromLocal.Count + "] files from local storage");

                // Find files present online but missing locally
                var missingLocal = checksumFromOnline
                    .Where(online => !checksumFromLocal.Any(local =>
                        string.Equals(local.File, online.File, StringComparison.OrdinalIgnoreCase)))
                    .ToList();

                // Find files present in both lists but with different checksums
                var differingChecksums = checksumFromOnline
                    .Join(
                        checksumFromLocal,
                        online => online.File,
                        local => local.File,
                        (online, local) => new { File = online.File, OnlineChecksum = online.Checksum, LocalChecksum = local.Checksum }
                    )
                    .Where(x => !string.Equals(x.OnlineChecksum, x.LocalChecksum, StringComparison.OrdinalIgnoreCase))
                    .ToList();


                // Combine missingLocal and differingChecksums to get the list of files to transfer from online to local
                // Only include files that are present in the online list
                var filesToTransfer = missingLocal
                    .Select(f => f.File)
                    .Concat(differingChecksums.Select(f => f.File))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Where(file => checksumFromOnline.Any(online => string.Equals(online.File, file, StringComparison.OrdinalIgnoreCase)))
                    .ToList();

                // ---

                int fileIndex = 0;
                int fileCount = filesToTransfer.Count;

                foreach (var file in filesToTransfer)
                {
                    fileIndex++;

                    // Find the online file entry (assuming DataUpdate has File and Url or similar)
                    var onlineFile = checksumFromOnline.FirstOrDefault(f =>
                        string.Equals(f.File, file, StringComparison.OrdinalIgnoreCase));
                    if (onlineFile == null)
                        continue; // Skip if not found online

                    // Update splash screen with progress
                    string shortName = Path.GetFileName(file);
                    Splashscreen.Current?.UpdateStatus(
                        "Downloading file " + fileIndex + " of " + fileCount + ":" +
                        Environment.NewLine + shortName
                    );

                    // Ensure the directory exists
                    string localPath = Path.Combine(DataPaths.DataRoot, file.Replace("/", "\\"));
                    string directory = Path.GetDirectoryName(localPath);
                    if (!Directory.Exists(directory))
                    {
                        Directory.CreateDirectory(directory);
                    }

                    // Download and overwrite the file
                    try
                    {
                        using (var webClient = new WebClient())
                        {
                            ServicePointManager.Expect100Continue = true;
                            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls13;
                            webClient.Headers.Add("user-agent", "CRT " + versionThis);

                            string decodedUrl = WebUtility.UrlDecode(onlineFile.Url);

                            // Check if the file is locked before downloading
                            if (!IsFileLocked(localPath))
                            {
                                webClient.DownloadFile(decodedUrl, localPath);
                            }
                            else
                            {
                                MessageBox.Show("The file [" + localPath + "] currently has an exclusive lock, and cannot be updated from the online source.\r\n\r\nPlease close any application that might be using it and retry the synchronization.",
                                    "ERROR: Cannot update file",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Error);
                            }
                        }
                        DebugOutput($"INFO: Downloaded and replaced file [{file}] from online source");
                    }
                    catch (Exception ex)
                    {
                        DebugOutput("EXCEPTION raised for fetching a specific file from online source:");
                        DebugOutput(ex.ToString());
                    }
                }

                // Silent mode: just log and update splash
                if (filesToTransfer.Count > 0)
                {
                    DebugOutput("INFO: Data sync completed - updated [" + filesToTransfer.Count + "] file(s)");
                    Splashscreen.Current?.UpdateStatus("Data update completed (" + filesToTransfer.Count + " files)");
                }
                else
                {
                    DebugOutput("INFO: Data sync completed - all files are up to date");
                    Splashscreen.Current?.UpdateStatus("All data files are up to date");
                }

                // Save to the configuration file, that we now have updated
                Configuration.SaveSetting("UpdateDataAtNextLaunch", "False");
            }
            catch (Exception ex)
            {
                DebugOutput("EXCEPTION raised for fetching JSON file catalogue from online source:");
                DebugOutput(ex.ToString());
                MessageBox.Show(ex.ToString(), "Error fetching JSON file catalogue", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        // ###########################################################################################
        // Get all local files referenced in the application, including the Excel file itself.
        // ###########################################################################################

        public static List<LocalFiles> GetAllReferencedLocalFiles()
        {
            var localFiles = new List<LocalFiles>();

            // Defensive: if root not set, abort early
            if (string.IsNullOrWhiteSpace(DataPaths.DataRoot))
                return localFiles;

            // Normalize root (absolute, no trailing slash, forward slashes)
            string rootFull = Path.GetFullPath(DataPaths.DataRoot)
                .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            string rootFullNorm = rootFull.Replace('\\', '/');

            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var filePath in Directory.GetFiles(rootFull, "*.*", SearchOption.AllDirectories))
            {
                string full = Path.GetFullPath(filePath);
                string fullNorm = full.Replace('\\', '/');

                // Strip root prefix (case-insensitive on Windows)
                string relative = fullNorm.StartsWith(rootFullNorm, StringComparison.OrdinalIgnoreCase)
                    ? fullNorm.Substring(rootFullNorm.Length).TrimStart('/')
                    : fullNorm; // fallback (shouldn't normally happen)

                // Avoid duplicates
                if (seen.Add(relative))
                {
                    localFiles.Add(new LocalFiles
                    {
                        File = relative, // store relative path only
                        Checksum = GetFileChecksum(full) // use full path for hashing
                    });
                }
            }

            return localFiles;
        }


        // ###########################################################################################
        // Calculate the SHA256 checksum of a file.
        // ###########################################################################################

        private static string GetFileChecksum(string filePath)
        {
            try
            {
                using (var sha = SHA256.Create())
                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    var hash = sha.ComputeHash(stream);
                    return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
                }
            }
            catch (IOException ex)
            {
                // File is locked (e.g., by Excel)
                Debug.WriteLine($"Cannot read file {filePath}: {ex.Message}");
                return "Cannot set checksum of file";
            }
        }


        // ###########################################################################################

        public void SetRegionButtonColors()
        {
            if (selectedRegion == "NTSC")
            {
                buttonRegionPal.FlatStyle = FlatStyle.Standard;
                buttonRegionPal.FlatAppearance.BorderColor = SystemColors.ControlDark;
                buttonRegionPal.FlatAppearance.BorderSize = 1;
                buttonRegionPal.BackColor = SystemColors.Control;
                buttonRegionPal.ForeColor = SystemColors.ControlText;
                buttonRegionNtsc.BackColor = Color.LightSteelBlue;
                buttonRegionNtsc.ForeColor = Color.Black;
                labelRegion.BackColor = Color.LightSteelBlue;
                labelRegion.ForeColor = Color.Black;
            }

            // selectedRegion == "PAL"
            else
            {
                buttonRegionPal.FlatStyle = FlatStyle.Flat;
                buttonRegionPal.FlatAppearance.BorderColor = Color.DarkRed;
                buttonRegionPal.FlatAppearance.BorderSize = 1;
                buttonRegionPal.BackColor = Color.IndianRed;
                buttonRegionPal.ForeColor = Color.White;
                buttonRegionNtsc.BackColor = SystemColors.Control;
                buttonRegionNtsc.ForeColor = SystemColors.ControlText;
                labelRegion.BackColor = Color.IndianRed;
                labelRegion.ForeColor = Color.White;
            }
            labelRegion.Text = selectedRegion;
        }


        // ###########################################################################################

        private void ButtonRegionPal_Click(object sender, EventArgs e)
        {
            selectedRegion = "PAL";
            Configuration.SaveSetting("SelectedRegion", selectedRegion);
            SetRegionButtonColors();
            UpdateComponentList("ButtonRegionPal_Click");
        }


        // ###########################################################################################

        private void ButtonRegionNtsc_Click(object sender, EventArgs e)
        {
            selectedRegion = "NTSC";
            Configuration.SaveSetting("SelectedRegion", selectedRegion);
            SetRegionButtonColors();
            UpdateComponentList("ButtonRegionNtsc_Click");
        }


        // ###########################################################################################

        private void checkBoxKeyboardZoom_CheckedChanged(object sender, EventArgs e)
        {
            // Get value from checkbox and save it to configuration file
            string isKeyboardZoomEnabled = checkBoxKeyboardZoom.Checked ? "1" : "0";

            Configuration.SaveSetting("KeyboardZoomEnabled", isKeyboardZoomEnabled);
        }


        // ###########################################################################################


        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, ref Rectangle lParam);

        private static void SetRichTextBoxPadding(RichTextBox rtb, Padding padding)
        {
            if (rtb == null) return;
            if (!rtb.IsHandleCreated) rtb.CreateControl();

            // Rectangle is in client coordinates.
            Rectangle rect = rtb.ClientRectangle;
            rect.X += padding.Left;
            rect.Y += padding.Top;
            rect.Width -= padding.Left + padding.Right;
            rect.Height -= padding.Top + padding.Bottom;

            if (rect.Width < 1) rect.Width = 1;
            if (rect.Height < 1) rect.Height = 1;

            SendMessage(rtb.Handle, EM_SETRECT, IntPtr.Zero, ref rect);

            // Redraw
            rtb.Invalidate();
        }


        // ###########################################################################################
        // Open external targets (URL or local file) with best-effort approach.
        // Platform dependent.
        // ###########################################################################################

        public static void OpenExternalTarget(string target, bool isUrl)
        {
            if (string.IsNullOrWhiteSpace(target))
            {
                DebugOutput("[OpenExternalTarget] Target is empty/null");
                return;
            }

            target = target.Trim();

            if (isUrl && !target.Contains("://"))
            {
                target = "https://" + target;
            }

            bool isWine = Environment.OSVersion.Platform == PlatformID.Win32NT && IsRunningUnderWine();

            if (isUrl)
            {
                if (!Uri.TryCreate(target, UriKind.Absolute, out var uri))
                {
                    DebugOutput("[OpenExternalTarget] Invalid URL: [" + target + "]");
                    return;
                }

                target = uri.AbsoluteUri;
            }
            else
            {
                target = NormalizePossiblyWinePath(target);

                if (!File.Exists(target))
                {
                    DebugOutput("[OpenExternalTarget] File does not exist: [" + target + "]");
                    return;
                }
            }

            DebugOutput("[OpenExternalTarget] Open isUrl=" + isUrl + " isWine=" + isWine + " target=[" + target + "]");

            // Under Wine+Mono, ShellExecute frequently throws "Path not found" for Unix paths.
            if (!isWine)
            {
                if (TryOpenWithShellExecute(target))
                {
                    DebugOutput("[OpenExternalTarget] ShellExecute OK");
                    return;
                }
            }
            else
            {
                DebugOutput("[OpenExternalTarget] Skipping ShellExecute under Wine");
            }

            if (TryOpenWithPlatformTools(target, isUrl))
            {
                DebugOutput("[OpenExternalTarget] Platform tools OK");
                return;
            }

            DebugOutput("[OpenExternalTarget] FAILED to open target=[" + target + "]");
        }

        private static bool TryOpenWithShellExecute(string target)
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = target,
                    UseShellExecute = true
                });
                return true;
            }
            catch (Exception ex)
            {
                DebugOutput("[OpenExternalTarget] ShellExecute failed: " + ex);
                return false;
            }
        }

        private static bool IsRunningUnderWine()
        {
            // Method 1: Check for Wine-specific environment variables
            if (!string.IsNullOrEmpty(Environment.GetEnvironmentVariable("WINELOADERNOEXEC"))) return true;
            if (!string.IsNullOrEmpty(Environment.GetEnvironmentVariable("WINEPREFIX"))) return true;
            if (!string.IsNullOrEmpty(Environment.GetEnvironmentVariable("WINESERVER"))) return true;

            // Method 2: Check if paths look like Wine Z: drive mapping
            try
            {
                string appPath = System.Windows.Forms.Application.StartupPath;
                if (!string.IsNullOrEmpty(appPath) &&
                    appPath.Length >= 2 &&
                    (appPath[0] == 'Z' || appPath[0] == 'z') &&
                    appPath[1] == ':')
                {
                    return true;
                }
            }
            catch { }

            // Method 3: Check for wine_get_version export in ntdll (most reliable)
            try
            {
                IntPtr ntdll = GetModuleHandle("ntdll.dll");
                if (ntdll != IntPtr.Zero)
                {
                    IntPtr wineGetVersion = GetProcAddress(ntdll, "wine_get_version");
                    if (wineGetVersion != IntPtr.Zero)
                    {
                        return true;
                    }
                }
            }
            catch { }

            return false;
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true)]
        private static extern IntPtr GetProcAddress(IntPtr hModule, string lpProcName);


        private static bool TryOpenWithPlatformTools(string target, bool isUrl)
        {
            bool isWindows = Environment.OSVersion.Platform == PlatformID.Win32NT;
            bool isUnix = Environment.OSVersion.Platform == PlatformID.Unix;
            bool isMac = Environment.OSVersion.Platform == PlatformID.MacOSX;

            try
            {
                if (isWindows && IsRunningUnderWine())
                {
                    // target is already a Unix path (converted in OpenExternalTarget)
                    DebugOutput("[OpenExternalTarget] Wine mode: target=[" + target + "]");

                    // Must go through bash so the shell handles quoting properly
                    // Mono's Process.Start doesn't handle quoted arguments correctly
                    string escaped = target.Replace("'", "'\\''");
                    string bashCmd = "xdg-open '" + escaped + "'";

                    return RunToolAndLog("/bin/bash", "-c \"" + bashCmd.Replace("\"", "\\\"") + "\"");
                }

                if (isWindows)
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = "cmd",
                        Arguments = "/c start \"\" \"" + target + "\"",
                        CreateNoWindow = true,
                        UseShellExecute = false
                    });
                    return true;
                }

                if (isMac)
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = "open",
                        Arguments = "\"" + target + "\"",
                        UseShellExecute = false
                    });
                    return true;
                }

                if (isUnix)
                {
                    string escaped = target.Replace("'", "'\\''");
                    string bashCmd = "xdg-open '" + escaped + "'";
                    return RunToolAndLog("/bin/bash", "-c \"" + bashCmd.Replace("\"", "\\\"") + "\"");
                }

                return false;
            }
            catch (Exception ex)
            {
                DebugOutput("[OpenExternalTarget] Platform open failed: " + ex);
                return false;
            }
        }

        private static bool RunToolAndLog(string fileName, string arguments)
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = fileName,
                    Arguments = arguments,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardError = true,
                    RedirectStandardOutput = true
                };

                DebugOutput("[OpenExternalTarget] Exec: " + fileName + " " + arguments);

                using (var p = Process.Start(psi))
                {
                    if (p == null)
                    {
                        DebugOutput("[OpenExternalTarget] Exec failed: Process.Start returned null");
                        return false;
                    }

                    // xdg-open / wine start may either exit quickly or stay around.
                    // If it doesn't exit quickly, assume it successfully launched the handler.
                    bool exited = p.WaitForExit(2000);

                    if (exited)
                    {
                        string stdout = SafeReadToEnd(p.StandardOutput);
                        string stderr = SafeReadToEnd(p.StandardError);

                        DebugOutput("[OpenExternalTarget] Exited=True ExitCode=" + p.ExitCode);
                        if (!string.IsNullOrWhiteSpace(stdout)) DebugOutput("[OpenExternalTarget] stdout: " + stdout.Trim());
                        if (!string.IsNullOrWhiteSpace(stderr)) DebugOutput("[OpenExternalTarget] stderr: " + stderr.Trim());

                        return p.ExitCode == 0;
                    }

                    DebugOutput("[OpenExternalTarget] Exited=False (assuming launched OK): " + fileName);
                    return true;
                }
            }
            catch (Exception ex)
            {
                DebugOutput("[OpenExternalTarget] Exec exception: " + ex);
                return false;
            }
        }

        private static string SafeReadToEnd(StreamReader sr)
        {
            try { return sr.ReadToEnd(); }
            catch { return ""; }
        }

        private static string EscapeForBashSingleArg(string s)
        {
            // Wrap with single quotes and escape existing single quotes:  abc'd -> 'abc'"'"'d'
            if (s == null) return "''";
            return "'" + s.Replace("'", "'\"'\"'") + "'";
        }

        private static string NormalizePossiblyWinePath(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return path;

            // Convert Wine Z:\path\to\file => /path/to/file
            // Wine typically maps Z: to the Unix root.
            if (path.Length >= 3 &&
                (path[0] == 'Z' || path[0] == 'z') &&
                path[1] == ':' &&
                (path[2] == '\\' || path[2] == '/'))
            {
                string unix = path.Substring(2).Replace('\\', '/'); // keep leading slash
                DebugOutput("[OpenExternalTarget] Wine path detected. Converted [" + path + "] -> [" + unix + "]");
                return unix;
            }

            // Also handle "z:\"
            if (path.Length >= 3 &&
                (path[0] == 'z') &&
                path[1] == ':' &&
                (path[2] == '\\' || path[2] == '/'))
            {
                string unix = path.Substring(2).Replace('\\', '/');
                DebugOutput("[OpenExternalTarget] Wine path detected. Converted [" + path + "] -> [" + unix + "]");
                return unix;
            }

            // Otherwise: normalize slashes and full path
            try
            {
                return Path.GetFullPath(path);
            }
            catch
            {
                return path;
            }
        }


        // ###########################################################################################

    }


    // ###########################################################################################
    // Class definitions.
    // ###########################################################################################

    // "Hardware" is read from the Excel file in level 1, "Commodore-Repair-Toolbox.xlsx".
    // It contains a list of all associated boards.
    public class Hardware
    {
        public string Name { get; set; }
        public List<Board> Boards { get; set; }
    }

    // "Board" is read from the Excel file in level 1, "Commodore-Repair-Toolbox.xlsx".
    public class Board
    {
        public string Name { get; set; }
        public string RevisionDate { get; set; }
        public string DataFile { get; set; }
        public List<BoardOverlays> Files { get; set; }
        public List<BoardComponents> Components { get; set; }
        public List<BoardComponentUserNotes> ComponentUserNotes { get; set; }
        public List<BoardLinks> BoardLinks { get; set; }
        public List<BoardLocalFiles> BoardLocalFiles { get; set; }
        public List<BoardCredits> BoardCredits { get; set; }
    }

    public class BoardComponents
    {
        public string Label { get; set; }
        public string NameTechnical { get; set; }
        public string NameFriendly { get; set; }
        public string Partnumber { get; set; }
        public string NameDisplay { get; set; }
        public string Type { get; set; }
        public string Region { get; set; }
        public string OneLiner { get; set; }
        public List<ComponentLocalFiles> LocalFiles { get; set; }
        public List<ComponentLinks> ComponentLinks { get; set; }
        public List<ComponentImages> ComponentImages { get; set; }
    }

    public class BoardComponentUserNotes
    {
        public string Label { get; set; }
        public string OneLiner { get; set; }
        public string Description { get; set; }
    }

    public class BoardOverlays
    {
        public string Name { get; set; }
        public string SchematicFileName { get; set; }
        public string HighlightColorTab { get; set; }
        public string HighlightColorList { get; set; }
        public int HighlightOpacityTab { get; set; }
        public int HighlightOpacityList { get; set; }
        public List<ComponentBounds> Components { get; set; }
    }

    public class BoardLinks
    {
        public string Category { get; set; }
        public string Name { get; set; }
        public string Url { get; set; }
    }

    public class BoardLocalFiles
    {
        public string Category { get; set; }
        public string Name { get; set; }
        public string Datafile { get; set; }
    }

    public class ComponentBounds
    {
        public string Label { get; set; }
        public List<Overlay> Overlays { get; set; }
    }

    public class Overlay
    {
        public Rectangle Bounds { get; set; }
    }

    public class ComponentLocalFiles
    {
        public string Name { get; set; }
        public string FileName { get; set; }
    }

    public class ComponentLinks
    {
        public string Name { get; set; }
        public string Url { get; set; }
    }

    public class ComponentImages
    {
        public string Region { get; set; }
        public string Pin { get; set; }
        public string Name { get; set; }
        public string Reading { get; set; }

        public string FileName { get; set; }
        public string Note { get; set; }
    }

    public class CustomPanel : Panel
    {
        public event MouseEventHandler CustomMouseWheel;
        protected override void OnMouseWheel(MouseEventArgs e)
        {
            CustomMouseWheel?.Invoke(this, e);
        }
    }

    // Add this class to represent a local file entry
    public class LocalFiles
    {
        public string File { get; set; }
        public string Checksum { get; set; }
    }

    public class BoardCredits
    {
        public string Category { get; set; }
        public string SubCategory { get; set; }
        public string Name { get; set; }
        public string Contact { get; set; }
    }

}