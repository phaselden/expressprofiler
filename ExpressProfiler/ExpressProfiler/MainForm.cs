// Forked from:
//      sample application for demonstrating Sql Server Profiling
//      writen by Locky, 2009.

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Serialization;
using EdtDbProfiler.EventComparers;

namespace EdtDbProfiler
{
    public partial class MainForm : Form
    {
        internal const string VersionString = "EDT DB Profiler v3.0";

        private class PerfInfo
        {
            internal int Count;
            internal readonly DateTime Date = DateTime.Now;
        }

        public class PerfColumn
        {
            public string  Caption;
            public int Column;
            public int Width;
            public string Format;
            public HorizontalAlignment Alignment = HorizontalAlignment.Left;
        }

        private enum ProfilingState
        {
            Stopped, 
            Profiling, 
            Paused
        }

        private RawTraceReader _reader;

        private readonly YukonLexer _lexer = new YukonLexer();
        private SqlConnection _connection;
        private readonly SqlCommand _command = new SqlCommand();
        private Thread _thread;
        private bool _needStop = true;
        private ProfilingState _profilingState ;
        private int _eventCount;
        private readonly ProfilerEvent _eventStarted = new ProfilerEvent();
        private readonly ProfilerEvent _eventStopped = new ProfilerEvent();
        private readonly ProfilerEvent _eventPaused = new ProfilerEvent();
        internal readonly List<ListViewItem> _cached = new List<ListViewItem>(1024);
		internal readonly List<ListViewItem> _cachedUnFiltered = new List<ListViewItem>(1024);
        private readonly Dictionary<string,ListViewItem> _itemBySql = new Dictionary<string, ListViewItem>();
        private string _servername = "";
        private string _username = "";
        private string _userPassword = "";
        internal int _lastPos = -1;
        internal string _lastPattern = "";
        private ListViewNF _eventsListView;
        Queue<ProfilerEvent> _events = new Queue<ProfilerEvent>(10);
        private bool _autoStart;
        private bool _dontUpdateSource;
        private Exception _profilerException;
        private readonly Queue<PerfInfo> _perfQueue = new Queue<PerfInfo>();
        private PerfInfo _first, _prev;
        internal TraceProperties.TraceSettings _currentSettings;
        private readonly List<PerfColumn> _columns = new List<PerfColumn>();
        internal bool _matchCase = false;
        internal bool _wholeWord = false;

        public MainForm()
        {
            InitializeComponent();
            tbStart.DefaultItem = tbRun;
            Text = VersionString;
            edPassword.TextBox.PasswordChar = '*';
            _servername = Properties.Settings.Default.ServerName;
            _username = Properties.Settings.Default.UserName;
            _currentSettings = GetDefaultSettings();
            ParseCommandLine();
            InitListView();
            SetDefaults();
            edServer.Text = _servername;
            edUser.Text = _username;
            edPassword.Text = _userPassword;
            tbAuth.SelectedIndex = String.IsNullOrEmpty(_username)?0:1;
            if(_autoStart) RunProfiling(false);
            UpdateButtons();
            webBrowser1.AllowNavigation = false;
        }

        private void SetDefaults()
        {
            _servername = Environment.MachineName;
            
            _currentSettings.Filters.ApplicationNameFilterCondition = TraceProperties.StringFilterCondition.NotLike;
            _currentSettings.Filters.ApplicationName = "EDT Agent";

            _currentSettings.Filters.DatabaseName = "eD%";
        }

        private TraceProperties.TraceSettings GetDefaultSettings()
        {
            try
            {
                var serializer = new XmlSerializer(typeof(TraceProperties.TraceSettings));
                using (var sr = new StringReader(Properties.Settings.Default.TraceSettings))
                {
                    return (TraceProperties.TraceSettings)serializer.Deserialize(sr);
                }
            }
            catch (Exception)
            {
                
            }
            return TraceProperties.TraceSettings.GetDefaultSettings();
        }



//DatabaseName = Filters.DatabaseName,
//LoginName = Filters.LoginName,
//HostName = Filters.HostName,
//TextData = Filters.TextData,
//ApplicationName = Filters.ApplicationName,



        private bool ParseFilterParam(string[] args, int idx)
        {
            var condition = idx + 1 < args.Length ? args[idx + 1] : "";
            var value = idx + 2 < args.Length ? args[idx + 2] : "";

            switch (args[idx].ToLower())
            {
                case "-cpu":
                    _currentSettings.Filters.CPU = Int32.Parse(value);
                    _currentSettings.Filters.CpuFilterCondition = TraceProperties.ParseIntCondition(condition);
                    break;
                case "-duration":
                    _currentSettings.Filters.Duration = Int32.Parse(value);
                    _currentSettings.Filters.DurationFilterCondition = TraceProperties.ParseIntCondition(condition);
                    break;
                case "-reads":
                    _currentSettings.Filters.Reads = Int32.Parse(value);
                    _currentSettings.Filters.ReadsFilterCondition = TraceProperties.ParseIntCondition(condition);
                    break;
                case "-writes":
                    _currentSettings.Filters.Writes = Int32.Parse(value);
                    _currentSettings.Filters.WritesFilterCondition = TraceProperties.ParseIntCondition(condition);
                    break;
                case "-spid":
                    _currentSettings.Filters.SPID = Int32.Parse(value);
                    _currentSettings.Filters.SPIDFilterCondition = TraceProperties.ParseIntCondition(condition);
                    break;
                case "-databasename":
                    _currentSettings.Filters.DatabaseName = value;
                    _currentSettings.Filters.DatabaseNameFilterCondition = TraceProperties.ParseStringCondition(condition);
                    break;
                case "-loginname":
                    _currentSettings.Filters.LoginName = value;
                    _currentSettings.Filters.LoginNameFilterCondition = TraceProperties.ParseStringCondition(condition);
                    break;
                case "-hostname":
                    _currentSettings.Filters.HostName = value;
                    _currentSettings.Filters.HostNameFilterCondition = TraceProperties.ParseStringCondition(condition);
                    break;
                case "-textdata":
                    _currentSettings.Filters.TextData = value;
                    _currentSettings.Filters.TextDataFilterCondition = TraceProperties.ParseStringCondition(condition);
                    break;
                case "-applicationname":
                    _currentSettings.Filters.ApplicationName = value;
                    _currentSettings.Filters.ApplicationNameFilterCondition = TraceProperties.ParseStringCondition(condition);
                    break;

            }
            return false;
        }

        private void ParseCommandLine()
        {
            try
            {
                string[] args = Environment.GetCommandLineArgs();
                int i = 1;
                while (i < args.Length)
                {
                    string ep = i + 1 < args.Length ? args[i + 1] : "";
                    switch (args[i].ToLower())
                    {
                        case "-s":
                        case "-server":
                            _servername = ep;
                            i++;
                            break;
                        case "-u":
                        case "-user":
                            _username = ep;
                            i++;
                            break;
                        case "-p":
                        case "-password":
                            _userPassword = ep;
                            i++;
                            break;
                        case "-m":
                        case "-maxevents":
                            int m;
                            if (!Int32.TryParse(ep, out m)) m = 1000;
                            _currentSettings.Filters.MaximumEventCount = m;
                            break;
                        case "-d":
                        case "-duration":
                            int d;
                            if (Int32.TryParse(ep, out d))
                            {
                                _currentSettings.Filters.DurationFilterCondition = TraceProperties.IntFilterCondition.GreaterThan;
                                _currentSettings.Filters.Duration = d;
                            }
                            break;
                        case "-start":
                            _autoStart = true;
                            break;
                        case "-batchcompleted":
                            _currentSettings.EventsColumns.BatchCompleted = true;
                            break;
                        case "-batchstarting":
                            _currentSettings.EventsColumns.BatchStarting = true;
                            break;
                        case "-existingconnection":
                            _currentSettings.EventsColumns.ExistingConnection = true;
                            break;
                        case "-loginlogout":
                            _currentSettings.EventsColumns.LoginLogout = true;
                            break;
                        case "-rpccompleted":
                            _currentSettings.EventsColumns.RPCCompleted = true;
                            break;
                        case "-rpcstarting":
                            _currentSettings.EventsColumns.RPCStarting = true;
                            break;
                        case "-spstmtcompleted":
                            _currentSettings.EventsColumns.SPStmtCompleted = true;
                            break;
                        case "-spstmtstarting":
                            _currentSettings.EventsColumns.SPStmtStarting = true;
                            break;
                        default:
                            if (ParseFilterParam(args, i)) i++;
                            break;
                    }
                    i++;
                }

                if (_servername.Length == 0)
                {
                    _servername = @".\sqlexpress";
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    
        private void tbStart_Click(object sender, EventArgs e)
        {

            if (!TraceProperties.AtLeastOneEventSelected(_currentSettings))
            {
                MessageBox.Show("You should select at least 1 event", "Starting trace", MessageBoxButtons.OK, MessageBoxIcon.Error);
                RunProfiling(true);
            }
            {
                RunProfiling(false);
            }
        }

        private void UpdateButtons()
        {
            tbStart.Enabled = _profilingState==ProfilingState.Stopped||_profilingState==ProfilingState.Paused;
            tbRun.Enabled = tbStart.Enabled;
            mnRun.Enabled = tbRun.Enabled;
            tbRunWithFilters.Enabled = ProfilingState.Stopped==_profilingState;
            mnRunWithFilters.Enabled = tbRunWithFilters.Enabled;
            startTraceToolStripMenuItem.Enabled = tbStart.Enabled;
            tbStop.Enabled = _profilingState==ProfilingState.Paused||_profilingState==ProfilingState.Profiling;
            stopTraceToolStripMenuItem.Enabled = tbStop.Enabled;
            tbPause.Enabled = _profilingState == ProfilingState.Profiling;
            pauseTraceToolStripMenuItem.Enabled = tbPause.Enabled;
            timer1.Enabled = _profilingState == ProfilingState.Profiling;
            edServer.Enabled = _profilingState == ProfilingState.Stopped;
            tbAuth.Enabled = _profilingState == ProfilingState.Stopped;
            edUser.Enabled = edServer.Enabled&&(tbAuth.SelectedIndex==1);
            edPassword.Enabled = edServer.Enabled && (tbAuth.SelectedIndex == 1);
        }
        
        private void InitListView()
        {
            _eventsListView = new ListViewNF
                           {
                               Dock = DockStyle.Fill,
                               Location = new System.Drawing.Point(0, 0),
                               Name = "_eventsListView",
                               Size = new System.Drawing.Size(979, 297),
                               TabIndex = 0,
                               VirtualMode = true,
                               UseCompatibleStateImageBehavior = false,
                               BorderStyle = BorderStyle.None,
                               FullRowSelect = true,
                               View = View.Details,
                               GridLines = true,
                               HideSelection = false,
                               MultiSelect = true,
                               AllowColumnReorder = false
                           };
            _eventsListView.RetrieveVirtualItem += EventsListViewRetrieveVirtualItem;
            _eventsListView.KeyDown += EventsListViewKeyDown;
            _eventsListView.ItemSelectionChanged += listView1_ItemSelectionChanged_1;
            _eventsListView.ColumnClick += EventsListViewColumnClick;
            _eventsListView.SelectedIndexChanged += EventsListViewSelectedIndexChanged;
            _eventsListView.VirtualItemsSelectionRangeChanged += EventsListViewOnVirtualItemsSelectionRangeChanged;
            _eventsListView.ContextMenuStrip = contextMenuStrip1;
            splitContainer1.Panel1.Controls.Add(_eventsListView);
            InitColumns();
            InitGridColumns();
        }

        private void InitColumns()
        {
            _columns.Clear();
            _columns.Add(new PerfColumn{ Caption = "Event Class", Column = ProfilerEventColumns.EventClass,Width = 122});
            _columns.Add(new PerfColumn { Caption = "Text Data", Column = ProfilerEventColumns.TextData, Width = 255});
            _columns.Add(new PerfColumn { Caption = "Login Name", Column = ProfilerEventColumns.LoginName, Width = 79 });
            _columns.Add(new PerfColumn { Caption = "CPU", Column = ProfilerEventColumns.CPU, Width = 40, Alignment = HorizontalAlignment.Right, Format = "#,0" });
            _columns.Add(new PerfColumn { Caption = "Reads", Column = ProfilerEventColumns.Reads, Width = 46, Alignment = HorizontalAlignment.Right, Format = "#,0" });
            _columns.Add(new PerfColumn { Caption = "Writes", Column = ProfilerEventColumns.Writes, Width = 46, Alignment = HorizontalAlignment.Right, Format = "#,0" });
            _columns.Add(new PerfColumn { Caption = "Duration, ms", Column = ProfilerEventColumns.Duration, Width = 76, Alignment = HorizontalAlignment.Right, Format = "#,0" });
            _columns.Add(new PerfColumn { Caption = "SPID", Column = ProfilerEventColumns.SPID, Width = 40, Alignment = HorizontalAlignment.Right });

            if (_currentSettings.EventsColumns.StartTime) _columns.Add(new PerfColumn { Caption = "Start time", Column = ProfilerEventColumns.StartTime, Width = 140, Format = "yyyy-MM-dd hh:mm:ss.ffff" });
            if (_currentSettings.EventsColumns.EndTime) _columns.Add(new PerfColumn { Caption = "End time", Column = ProfilerEventColumns.EndTime, Width = 90, Format = "hh:mm:ss.ffff" });
            if (_currentSettings.EventsColumns.DatabaseName) _columns.Add(new PerfColumn { Caption = "DatabaseName", Column = ProfilerEventColumns.DatabaseName, Width = 156 });
            if (_currentSettings.EventsColumns.ObjectName) _columns.Add(new PerfColumn { Caption = "Object name", Column = ProfilerEventColumns.ObjectName, Width = 70 });
            if (_currentSettings.EventsColumns.ApplicationName) _columns.Add(new PerfColumn { Caption = "Application name", Column = ProfilerEventColumns.ApplicationName, Width = 80 });
            if (_currentSettings.EventsColumns.HostName) _columns.Add(new PerfColumn { Caption = "Host name", Column = ProfilerEventColumns.HostName, Width = 70 });

            _columns.Add(new PerfColumn { Caption = "#", Column = -1, Width = 40, Alignment = HorizontalAlignment.Right});
        }

        private void InitGridColumns()
        {
            InitColumns();
            _eventsListView.BeginUpdate();
            try
            {
                _eventsListView.Columns.Clear();
                foreach (PerfColumn pc in _columns)
                {
                    var l = _eventsListView.Columns.Add(pc.Caption, pc.Width);
                    l.TextAlign = pc.Alignment;
                }
            }
            finally
            {
                _eventsListView.EndUpdate();
            }
        }

        private void EventsListViewOnVirtualItemsSelectionRangeChanged(object sender, ListViewVirtualItemsSelectionRangeChangedEventArgs listViewVirtualItemsSelectionRangeChangedEventArgs)
        {
            UpdateSourceBox();
        }

        void EventsListViewSelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateSourceBox();
        }

        void EventsListViewColumnClick(object sender, ColumnClickEventArgs e)
        {
			_eventsListView.ToggleSortOrder();
			_eventsListView.SetSortIcon(e.Column, _eventsListView.SortOrder);
			var comparer = new TextDataComparer(e.Column, _eventsListView.SortOrder);
			_cached.Sort(comparer);
			UpdateSourceBox();
			ShowSelectedEvent();
        }

        private string GetEventCaption(ProfilerEvent evt)
        {
            if (evt == _eventStarted)
            {
                return "Trace started";
            }
            if (evt == _eventPaused)
            {
                return "Trace paused";
            }
            if (evt == _eventStopped)
            {
                return "Trace stopped";
            }
            return ProfilerEvents.Names[evt.EventClass];
        }

        private string GetFormattedValue(ProfilerEvent evt,int column,string format)
        {
            return ProfilerEventColumns.Duration == column ? (evt.Duration / 1000).ToString(format) : evt.GetFormattedData(column,format);
        }

        private void NewEventArrived(ProfilerEvent evt,bool last)
        {
            //TODO do this better
            if (evt.TextData == "exec sp_reset_connection")
                return;

            var current = (_eventsListView.SelectedIndices.Count > 0) ? _cached[_eventsListView.SelectedIndices[0]] : null;
            _eventCount++;
            var caption = GetEventCaption(evt);

            var listViewItem = new ListViewItem(caption);
            var items = new string[_columns.Count];
            for (var i = 1; i < _columns.Count; i++)
            {
                var perfColumn = _columns[i];
                items[i - 1] = perfColumn.Column == -1
                    ? _eventCount.ToString("#,0")
                    : GetFormattedValue(evt, perfColumn.Column, perfColumn.Format) ?? "";
            }

            listViewItem.SubItems.AddRange(items);
            listViewItem.Tag = evt;
            _cached.Add(listViewItem);
            if (last)
            {
                _eventsListView.VirtualListSize = _cached.Count;
                _eventsListView.SelectedIndices.Clear();
                FocusListViewItem(tbScroll.Checked ? _eventsListView.Items[_cached.Count - 1] : current, tbScroll.Checked);
                _eventsListView.Invalidate(listViewItem.Bounds);
            }
        }

        internal void FocusListViewItem(ListViewItem listViewItem, bool ensure)
        {
            if (null != listViewItem)
            {
                listViewItem.Focused = true;
                listViewItem.Selected = true;
                listView1_ItemSelectionChanged_1(_eventsListView, null);
                if (ensure)
                {
                    _eventsListView.EnsureVisible(_eventsListView.Items.IndexOf(listViewItem));
                }
            }
        }

        private void ProfilerThread(Object state)
        {
            try
            {
                while (!_needStop && _reader.TraceIsActive)
                {
                    ProfilerEvent profilerEvent = _reader.Next();
                    if (profilerEvent != null)
                    {
                        lock (this)
                        {
                            _events.Enqueue(profilerEvent);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                lock (this)
                {
                    if (!_needStop && _reader.TraceIsActive)
                    {
                        _profilerException = e;
                    }
                }
            }
        }

        private SqlConnection GetConnection()
        {
            return new SqlConnection
                {
                    ConnectionString = tbAuth.SelectedIndex == 0
                        ? $@"Data Source={edServer.Text}; Initial Catalog = master; Integrated Security=SSPI; Application Name=EDT DB Profiler"
                        : $@"Data Source={edServer.Text}; Initial Catalog=master; User Id={edUser.Text}; Password='{edPassword.Text}';; Application Name=EDT DB Profiler"
                };
        }

        private void StartProfiling()
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                _perfQueue.Clear();
                _first = null;
                _prev = null;
                if (_profilingState == ProfilingState.Paused)
                {
                    ResumeProfiling();
                    return;
                }
                if (_connection != null && _connection.State == ConnectionState.Open)
                {
                    _connection.Close();
                }
                InitGridColumns();
                _eventCount = 0;
                _connection = GetConnection();
                _connection.Open();
                _reader = new RawTraceReader(_connection);

                _reader.CreateTrace();
                if (true)
                {
                    if (_currentSettings.EventsColumns.LoginLogout)
                    {
                        _reader.SetEvent(ProfilerEvents.SecurityAudit.AuditLogin,
                                       ProfilerEventColumns.TextData,
                                       ProfilerEventColumns.LoginName,
                                       ProfilerEventColumns.SPID,
                                       ProfilerEventColumns.StartTime,
                                       ProfilerEventColumns.EndTime,
                                       ProfilerEventColumns.HostName
                            );
                        _reader.SetEvent(ProfilerEvents.SecurityAudit.AuditLogout,
                                       ProfilerEventColumns.CPU,
                                       ProfilerEventColumns.Reads,
                                       ProfilerEventColumns.Writes,
                                       ProfilerEventColumns.Duration,
                                       ProfilerEventColumns.LoginName,
                                       ProfilerEventColumns.SPID,
                                       ProfilerEventColumns.StartTime,
                                       ProfilerEventColumns.EndTime,
                                       ProfilerEventColumns.ApplicationName,
                                       ProfilerEventColumns.HostName
                            );
                    }

                    if (_currentSettings.EventsColumns.ExistingConnection)
                    {
                        _reader.SetEvent(ProfilerEvents.Sessions.ExistingConnection,
                                       ProfilerEventColumns.TextData,
                                       ProfilerEventColumns.SPID,
                                       ProfilerEventColumns.StartTime,
                                       ProfilerEventColumns.EndTime,
                                       ProfilerEventColumns.ApplicationName,
                                       ProfilerEventColumns.HostName
                            );
                    }
                    if (_currentSettings.EventsColumns.BatchCompleted)
                    {
                        _reader.SetEvent(ProfilerEvents.TSQL.SQLBatchCompleted,
                                       ProfilerEventColumns.TextData,
                                       ProfilerEventColumns.LoginName,
                                       ProfilerEventColumns.CPU,
                                       ProfilerEventColumns.Reads,
                                       ProfilerEventColumns.Writes,
                                       ProfilerEventColumns.Duration,
                                       ProfilerEventColumns.SPID,
                                       ProfilerEventColumns.StartTime,
                                       ProfilerEventColumns.EndTime,
                                       ProfilerEventColumns.DatabaseName,
                                       ProfilerEventColumns.ApplicationName,
                                       ProfilerEventColumns.HostName
                            );
                    }
                    if (_currentSettings.EventsColumns.BatchStarting)
                    {
                        _reader.SetEvent(ProfilerEvents.TSQL.SQLBatchStarting,
                                       ProfilerEventColumns.TextData,
                                       ProfilerEventColumns.LoginName,
                                       ProfilerEventColumns.SPID,
                                       ProfilerEventColumns.StartTime,
                                       ProfilerEventColumns.EndTime,
                                       ProfilerEventColumns.DatabaseName,
                                       ProfilerEventColumns.ApplicationName,
                                       ProfilerEventColumns.HostName
                            );
                    }
                    if (_currentSettings.EventsColumns.RPCStarting)
                    {
                        _reader.SetEvent(ProfilerEvents.StoredProcedures.RPCStarting,
                                       ProfilerEventColumns.TextData,
                                       ProfilerEventColumns.LoginName,
                                       ProfilerEventColumns.SPID,
                                       ProfilerEventColumns.StartTime,
                                       ProfilerEventColumns.EndTime,
                                       ProfilerEventColumns.DatabaseName,
                                       ProfilerEventColumns.ObjectName,
                                       ProfilerEventColumns.ApplicationName,
                                       ProfilerEventColumns.HostName

                            );
                    }

                }
                if (_currentSettings.EventsColumns.RPCCompleted)
                {
                    _reader.SetEvent(ProfilerEvents.StoredProcedures.RPCCompleted,
                                   ProfilerEventColumns.TextData, ProfilerEventColumns.LoginName,
                                   ProfilerEventColumns.CPU, ProfilerEventColumns.Reads,
                                   ProfilerEventColumns.Writes, ProfilerEventColumns.Duration,
                                   ProfilerEventColumns.SPID
                                   , ProfilerEventColumns.StartTime, ProfilerEventColumns.EndTime
                                   , ProfilerEventColumns.DatabaseName
                                   , ProfilerEventColumns.ObjectName
                                   , ProfilerEventColumns.ApplicationName
                                   , ProfilerEventColumns.HostName
                    );
                }
                if (_currentSettings.EventsColumns.SPStmtCompleted)
                {
                    _reader.SetEvent(ProfilerEvents.StoredProcedures.SPStmtCompleted,
                                   ProfilerEventColumns.TextData, ProfilerEventColumns.LoginName,
                                   ProfilerEventColumns.CPU, ProfilerEventColumns.Reads,
                                   ProfilerEventColumns.Writes, ProfilerEventColumns.Duration,
                                   ProfilerEventColumns.SPID
                                   , ProfilerEventColumns.StartTime, ProfilerEventColumns.EndTime
                                   , ProfilerEventColumns.DatabaseName
                                   , ProfilerEventColumns.ObjectName
                                   , ProfilerEventColumns.ObjectID
                                   , ProfilerEventColumns.ApplicationName
                                   , ProfilerEventColumns.HostName
                        );
                }
                if (_currentSettings.EventsColumns.SPStmtStarting)
                {
                    _reader.SetEvent(ProfilerEvents.StoredProcedures.SPStmtStarting,
                                   ProfilerEventColumns.TextData, ProfilerEventColumns.LoginName,
                                   ProfilerEventColumns.CPU, ProfilerEventColumns.Reads,
                                   ProfilerEventColumns.Writes, ProfilerEventColumns.Duration,
                                   ProfilerEventColumns.SPID
                                   , ProfilerEventColumns.StartTime, ProfilerEventColumns.EndTime
                                   , ProfilerEventColumns.DatabaseName
                                   , ProfilerEventColumns.ObjectName
                                   , ProfilerEventColumns.ObjectID
                                   , ProfilerEventColumns.ApplicationName
                                   , ProfilerEventColumns.HostName
                        );
                }
                if (_currentSettings.EventsColumns.UserErrorMessage)
                {
                    _reader.SetEvent(ProfilerEvents.ErrorsAndWarnings.UserErrorMessage,
                                   ProfilerEventColumns.TextData,
                                   ProfilerEventColumns.LoginName,
                                   ProfilerEventColumns.CPU,
                                   ProfilerEventColumns.SPID,
                                   ProfilerEventColumns.StartTime,
                                   ProfilerEventColumns.DatabaseName,
                                   ProfilerEventColumns.ApplicationName
                                   , ProfilerEventColumns.HostName
                        );
                }
                if (_currentSettings.EventsColumns.BlockedProcessPeport)
                {
                    _reader.SetEvent(ProfilerEvents.ErrorsAndWarnings.Blockedprocessreport,
                                   ProfilerEventColumns.TextData,
                                   ProfilerEventColumns.LoginName,
                                   ProfilerEventColumns.CPU,
                                   ProfilerEventColumns.SPID,
                                   ProfilerEventColumns.StartTime,
                                   ProfilerEventColumns.DatabaseName,
                                   ProfilerEventColumns.ApplicationName
                                   , ProfilerEventColumns.HostName
                        );
                }
                if (_currentSettings.EventsColumns.SQLStmtStarting)
                {
                    _reader.SetEvent(ProfilerEvents.TSQL.SQLStmtStarting,
                                   ProfilerEventColumns.TextData, ProfilerEventColumns.LoginName,
                                   ProfilerEventColumns.CPU, ProfilerEventColumns.Reads,
                                   ProfilerEventColumns.Writes, ProfilerEventColumns.Duration,
                                   ProfilerEventColumns.SPID
                                   , ProfilerEventColumns.StartTime, ProfilerEventColumns.EndTime
                                   , ProfilerEventColumns.DatabaseName
                                   , ProfilerEventColumns.ApplicationName
                                   , ProfilerEventColumns.HostName
                        );
                }
                if (_currentSettings.EventsColumns.SQLStmtCompleted)
                {
                    _reader.SetEvent(ProfilerEvents.TSQL.SQLStmtCompleted,
                                   ProfilerEventColumns.TextData, ProfilerEventColumns.LoginName,
                                   ProfilerEventColumns.CPU, ProfilerEventColumns.Reads,
                                   ProfilerEventColumns.Writes, ProfilerEventColumns.Duration,
                                   ProfilerEventColumns.SPID
                                   , ProfilerEventColumns.StartTime, ProfilerEventColumns.EndTime
                                   , ProfilerEventColumns.DatabaseName
                                   , ProfilerEventColumns.ApplicationName
                                   , ProfilerEventColumns.HostName
                        );
                }

                if (null != _currentSettings.Filters.Duration)
                {
                    SetIntFilter(_currentSettings.Filters.Duration*1000,
                                 _currentSettings.Filters.DurationFilterCondition, ProfilerEventColumns.Duration);
                }
                SetIntFilter(_currentSettings.Filters.Reads, _currentSettings.Filters.ReadsFilterCondition,ProfilerEventColumns.Reads);
                SetIntFilter(_currentSettings.Filters.Writes, _currentSettings.Filters.WritesFilterCondition,ProfilerEventColumns.Writes);
                SetIntFilter(_currentSettings.Filters.CPU, _currentSettings.Filters.CpuFilterCondition,ProfilerEventColumns.CPU);
                SetIntFilter(_currentSettings.Filters.SPID, _currentSettings.Filters.SPIDFilterCondition, ProfilerEventColumns.SPID);

                SetStringFilter(_currentSettings.Filters.LoginName, _currentSettings.Filters.LoginNameFilterCondition,ProfilerEventColumns.LoginName);
                SetStringFilter(_currentSettings.Filters.HostName, _currentSettings.Filters.HostNameFilterCondition, ProfilerEventColumns.HostName);
                SetStringFilter(_currentSettings.Filters.DatabaseName,_currentSettings.Filters.DatabaseNameFilterCondition, ProfilerEventColumns.DatabaseName);
                SetStringFilter(_currentSettings.Filters.TextData, _currentSettings.Filters.TextDataFilterCondition,ProfilerEventColumns.TextData);
                SetStringFilter(_currentSettings.Filters.ApplicationName, _currentSettings.Filters.ApplicationNameFilterCondition, ProfilerEventColumns.ApplicationName);


                _command.Connection = _connection;
                _command.CommandTimeout = 0;
                _reader.SetFilter(ProfilerEventColumns.ApplicationName, LogicalOperators.AND, ComparisonOperators.NotLike,
                                "EDT DB Profiler");
                _cached.Clear();
                _events.Clear();
                _itemBySql.Clear();
                _eventsListView.VirtualListSize = 0;
                StartProfilerThread();
                _servername = edServer.Text;
                _username = edUser.Text;
                SaveDefaultSettings();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                UpdateButtons();
                Cursor = Cursors.Default;
            }
        }

	    private void SaveDefaultSettings()
	    {
		    Properties.Settings.Default.ServerName = _servername;
		    Properties.Settings.Default.UserName = tbAuth.SelectedIndex == 0 ? "" : _username;
		    Properties.Settings.Default.Save();
	    }

	    private void SetIntFilter(int? value, TraceProperties.IntFilterCondition condition, int column)
        {
            var com = new[] { ComparisonOperators.Equal, ComparisonOperators.NotEqual, ComparisonOperators.GreaterThan, ComparisonOperators.LessThan };
            if (null != value)
            {
                long? v = value;
                _reader.SetFilter(column, LogicalOperators.AND, com[(int)condition], v);
            }
        }

        private void SetStringFilter(string value,TraceProperties.StringFilterCondition condition,int column)
        {
            if (!String.IsNullOrEmpty(value))
            {
                _reader.SetFilter(column, LogicalOperators.AND
                    , condition == TraceProperties.StringFilterCondition.Like ? ComparisonOperators.Like : ComparisonOperators.NotLike
                    , value);
            }
        }

        private void StartProfilerThread()
        { 
            if (_reader != null)
            {
                _reader.Close();
            }
            _reader.StartTrace();
            _thread = new Thread(ProfilerThread) {IsBackground = true, Priority = ThreadPriority.Lowest};
            _needStop = false;
            _profilingState = ProfilingState.Profiling;
            NewEventArrived(_eventStarted,true);
            _thread.Start();
        }

        private void ResumeProfiling()
        {
            StartProfilerThread();
            UpdateButtons();
        }

        private void tbStop_Click(object sender, EventArgs e)
        {
            StopProfiling();
        }

        private void StopProfiling()
        {
            tbStop.Enabled = false;
            using (var sqlConnection = GetConnection())
            {
                sqlConnection.Open();
                _reader.StopTrace(sqlConnection);
                _reader.CloseTrace(sqlConnection);
                sqlConnection.Close();
            }
            _needStop = true;
            if (_thread.IsAlive)
            {
                _thread.Abort();
            }
            _thread.Join();
            _profilingState = ProfilingState.Stopped;
            NewEventArrived(_eventStopped,true);
            UpdateButtons();
        }

        private void listView1_ItemSelectionChanged_1(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            UpdateSourceBox();
        }

        private void UpdateSourceBox()
        {
            if (_dontUpdateSource) 
                return;

            var sb = new StringBuilder();
            foreach (int i in _eventsListView.SelectedIndices)
            {
                ListViewItem lv = _cached[i];
                if (lv.SubItems[1].Text != "")
                {
                    //sb.AppendFormat("{0}\r\ngo\r\n", lv.SubItems[1].Text);
                    sb.AppendLine(lv.SubItems[1].Text);
                }
            }

            _lexer.FillRichEdit(reTextData, sb.ToString());

            var res = SqlFormatter.ParseSql(sb.ToString());
            _lexer.FillRichEdit(richTextBox1, res.Sql);
            
            DisplayHtml(SqlFormatter.Format(sb.ToString()));
        }

        private void DisplayHtml(string html)
        {
            webBrowser1.Navigate("about:blank");
            webBrowser1.Document?.OpenNew(false);
            webBrowser1.Document?.Write(html);
            webBrowser1.Refresh();
        }
        
        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (_profilingState == ProfilingState.Paused || _profilingState == ProfilingState.Profiling)
            {
                StopProfiling();
                // if (MessageBox.Show("There are traces still running. Are you sure you want to stop profiling and close the application?","ExpressProfiler",
                //     MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.Yes)
                // {
                //     StopProfiling();
                // }
                // else
                // {
                //     e.Cancel = true;
                // }
            }
        }

        private void EventsListViewRetrieveVirtualItem(object sender, RetrieveVirtualItemEventArgs e)
        {
			e.Item = _cached[e.ItemIndex];
        }

        private void tbPause_Click(object sender, EventArgs e)
        {
            PauseProfiling();
        }

        private void PauseProfiling()
        {
            using (SqlConnection cn = GetConnection())
            {
                cn.Open();
                _reader.StopTrace(cn);
                cn.Close();
            }
            _profilingState = ProfilingState.Paused;
            NewEventArrived(_eventPaused,true);
            UpdateButtons();
        }

        internal void SelectAllEvents(bool select)
        {
            lock (_cached)
            {
                _eventsListView.BeginUpdate();
                _dontUpdateSource = true;
                try
                {
                    foreach (var listViewItem in _cached)
                    {
                        listViewItem.Selected = select;
                    }
                }
                finally
                {
                    _dontUpdateSource = false;
                    UpdateSourceBox();
                    _eventsListView.EndUpdate();
                }
            }
        }

        private void EventsListViewKeyDown(object sender, KeyEventArgs e)
        {
        }

        private void MainForm_Load(object sender, EventArgs e)
        {

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            Queue<ProfilerEvent> saved;
            Exception exc;
            lock (this)
            {
                saved = _events;
                _events = new Queue<ProfilerEvent>(10);
                exc = _profilerException;
                _profilerException = null;
            }
            if (null != exc)
            {
                using (var threadExceptionDialog = new ThreadExceptionDialog(exc))
                {
                    threadExceptionDialog.ShowDialog();
                }
            }

            lock (_cached)
            {
                while (0 != saved.Count)
                {
                    NewEventArrived(saved.Dequeue(), 0 == saved.Count);
                }
                if (_cached.Count > _currentSettings.Filters.MaximumEventCount)
                {
                    while (_cached.Count > _currentSettings.Filters.MaximumEventCount)
                    {
                        _cached.RemoveAt(0);
                    }
                    _eventsListView.VirtualListSize = _cached.Count;
                    _eventsListView.Invalidate();
                }

                if (null == _prev || DateTime.Now.Subtract(_prev.Date).TotalSeconds >= 1)
                {
                    var currentPerfInfo = new PerfInfo { Count = _eventCount };
                    if (_perfQueue.Count >= 60)
                    {
                        _first = _perfQueue.Dequeue();
                    }
                    if (null == _first) _first = currentPerfInfo;
                    if (null == _prev) _prev = currentPerfInfo;

                    var now = DateTime.Now;
                    double d1 = now.Subtract(_prev.Date).TotalSeconds;
                    double d2 = now.Subtract(_first.Date).TotalSeconds;
                    slEPS.Text = String.Format("{0} / {1} EPS(last/avg for {2} second(s))",
                        Math.Abs(d1 - 0) > 0.001 
                            ? ((currentPerfInfo.Count - _prev.Count)/d1).ToString("#,0.00") 
                            : "",
                        Math.Abs(d2 - 0) > 0.001 
                            ? ((currentPerfInfo.Count - _first.Count) / d2).ToString("#,0.00") 
                            : "", 
                        d2 .ToString("0"));

                    _perfQueue.Enqueue(currentPerfInfo);
                    _prev = currentPerfInfo;
                }
            }
        }

        private void tbAuth_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateButtons();
        }

        private void ClearTrace()
        {
            lock (_eventsListView)
            {
                _cached.Clear();
                _itemBySql.Clear();
                _eventsListView.VirtualListSize = 0;
                listView1_ItemSelectionChanged_1(_eventsListView, null);
                _eventsListView.Invalidate();
            }
        }

        private void tbClear_Click(object sender, EventArgs e)
        {
            ClearTrace();
        }

        private void NewAttribute(XmlNode node, string name, string value)
        {
            XmlAttribute attr = node.OwnerDocument.CreateAttribute(name);
            attr.Value = value;
            node.Attributes.Append(attr);
        }
        private void NewAttribute(XmlNode node, string name, string value, string namespaceURI)
        {
            XmlAttribute attr = node.OwnerDocument.CreateAttribute("ss", name, namespaceURI);
            attr.Value = value;
            node.Attributes.Append(attr);
        }

        private void copyAllToClipboardToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CopyEventsToClipboard(false);
        }

        private void CopyEventsToClipboard(bool copySelected)
        {
            var doc = new XmlDocument();
            XmlNode root = doc.CreateElement("events");
            lock (_cached)
            {
                if (copySelected)
                {
                    foreach (int i in _eventsListView.SelectedIndices)
                    {
                        CreateEventRow((ProfilerEvent)(_cached[i]).Tag, root);
                    }
                }
                else
                {
                    foreach (var i in _cached)
                    {
                        CreateEventRow((ProfilerEvent)i.Tag, root);
                    }
                }
            }
            doc.AppendChild(root);
            doc.PreserveWhitespace = true;
            using (StringWriter writer = new StringWriter())
            {
                XmlTextWriter textWriter = new XmlTextWriter(writer) {Formatting = Formatting.Indented};
                doc.Save(textWriter);
                Clipboard.SetText(writer.ToString());
            }
            MessageBox.Show("Event(s) data copied to clipboard", "Information", MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
        }

        private void CreateEventRow(ProfilerEvent evt, XmlNode root)
        {
            XmlNode row = root.OwnerDocument.CreateElement("event");
            NewAttribute(row, "EventClass", evt.EventClass.ToString(CultureInfo.InvariantCulture));
            NewAttribute(row, "CPU", evt.CPU.ToString(CultureInfo.InvariantCulture));
            NewAttribute(row, "Reads", evt.Reads.ToString(CultureInfo.InvariantCulture));
            NewAttribute(row, "Writes", evt.Writes.ToString(CultureInfo.InvariantCulture));
            NewAttribute(row, "Duration", evt.Duration.ToString(CultureInfo.InvariantCulture));
            NewAttribute(row, "SPID", evt.SPID.ToString(CultureInfo.InvariantCulture));
            NewAttribute(row, "LoginName", evt.LoginName);
            NewAttribute(row, "DatabaseName", evt.DatabaseName);
            NewAttribute(row, "ObjectName", evt.ObjectName);
            NewAttribute(row, "ApplicationName", evt.ApplicationName);
            NewAttribute(row, "HostName", evt.HostName);
            row.InnerText = evt.TextData;
            root.AppendChild(row);
        }

        private void copySelectedToClipboardToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CopyEventsToClipboard(true);
        }

        private void clearTraceWindowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ClearTrace();
        }

        private void extractAllEventsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CopyEventsToClipboard(false);
        }

        private void extractSelectedEventsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CopyEventsToClipboard(true);
        }


        private void pauseTraceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PauseProfiling();
        }

        private void stopTraceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            StopProfiling();
        }

        private void findToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DoFind();
        }

        private void DoFind()
        {
            if (_profilingState == ProfilingState.Profiling)
            {
                MessageBox.Show("You cannot find when trace is running", "ExpressProfiler", MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
                return;
            }
            using (FindForm f = new FindForm(this))
            {
                f.TopMost = this.TopMost;
                f.ShowDialog();
            }
        }

        private void selectAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (_eventsListView.Focused && (_profilingState!=ProfilingState.Profiling))
            {
                SelectAllEvents(true);
            }
            else if (reTextData.Focused)
            {
                reTextData.SelectAll();
            }
        }

		//internal void PerformFind(bool forwards)
		//{
		//    if(String.IsNullOrEmpty(_lastPattern)) return;

		//    if (forwards)
		//    {
		//        for (int i = _lastPos = _eventsListView.Items.IndexOf(_eventsListView.FocusedItem) + 1; i < _cached.Count; i++)
		//        {
		//            if (FindText(i))
		//            {
		//                return;
		//            }
		//        }
		//    }
		//    else
		//    {
		//        for (int i = _lastPos = _eventsListView.Items.IndexOf(_eventsListView.FocusedItem) - 1; i > 0; i--)
		//        {
		//            if (FindText(i))
		//            {
		//                return;
		//            }
		//        }
		//    }
		//    MessageBox.Show(String.Format("Failed to find \"{0}\". Searched to the end of data. ", _lastPattern), "ExpressProfiler", MessageBoxButtons.OK, MessageBoxIcon.Information);
		//}


		internal void PerformFind(bool forwards, bool wrapAround)
		{
			if (String.IsNullOrEmpty(_lastPattern)) return;
			int lastPos = _eventsListView.Items.IndexOf(_eventsListView.FocusedItem);
			if (forwards)
			{
				for (var i = lastPos + 1; i < _cached.Count; i++)
				{
					if (FindText(i))
					{
						return;
					}
				}
				if (wrapAround)
				{
					for (var i = 0; i < lastPos; i++)
					{
						if (FindText(i))
						{
							return;
						}
					}
				}
			}
			else
			{
				for (var i = lastPos - 1; i > 0; i--)
				{
					if (FindText(i))
					{
						return;
					}
				}
				if (wrapAround)
				{
					for (int i = _cached.Count; i > lastPos; i--)
					{
						if (FindText(i))
						{
							return;
						}
					}
				}
			}
			MessageBox.Show(String.Format("Failed to find \"{0}\". Searched to the end of data. ", _lastPattern), 
                "ExpressProfiler", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}
        
		private void ShowSelectedEvent()
		{
			var focusedIndex = _eventsListView.Items.IndexOf(_eventsListView.FocusedItem);
			if (focusedIndex > -1 && (focusedIndex < _cached.Count))
			{
				var listViewItem = _cached[focusedIndex];
				ProfilerEvent evt = (ProfilerEvent) listViewItem.Tag;

				listViewItem.Focused = true;
				_lastPos = focusedIndex;
				SelectAllEvents(false);
				FocusListViewItem(listViewItem, true);
			}
		}
        
        private bool FindText(int i)
        {
            var listViewItem = _cached[i];
            ProfilerEvent evt = (ProfilerEvent) listViewItem.Tag;
            var pattern = (_wholeWord ? "\\b" + _lastPattern + "\\b" : _lastPattern);
            if (Regex.IsMatch(evt.TextData, pattern, (_matchCase ? RegexOptions.None : RegexOptions.IgnoreCase)))
            {
                listViewItem.Focused = true;
                _lastPos = i;
                SelectAllEvents(false);
                FocusListViewItem(listViewItem, true);
                return true;
            }

            return false;
        }

        private void findNextToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (_profilingState == ProfilingState.Profiling)
            {
                MessageBox.Show("You cannot find when trace is running", "ExpressProfiler", MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
                return;
            }
            PerformFind(true, false);
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        internal void RunProfiling(bool showfilters)
        {
            if (showfilters)
            {
                TraceProperties.TraceSettings ts = _currentSettings.GetCopy();
                using (TraceProperties frm = new TraceProperties())
                {
                    frm.SetSettings(ts);
                    if (DialogResult.OK != frm.ShowDialog()) return;
                    _currentSettings = frm.m_currentsettings.GetCopy();
                }
            }
            StartProfiling();
        }

        private void tbRunWithFilters_Click(object sender, EventArgs e)
        {
            RunProfiling(true);
        }

        private void copyToXlsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CopyForExcel();
        }

        private void CopyForExcel()
        {

            var doc = new XmlDocument();
            XmlProcessingInstruction pi = doc.CreateProcessingInstruction("mso-application", "progid='Excel.Sheet'");
            doc.AppendChild(pi); 
            const string urn = "urn:schemas-microsoft-com:office:spreadsheet";
            XmlNode root = doc.CreateElement("ss","Workbook",urn);
            NewAttribute(root, "xmlns:ss", urn);
            doc.AppendChild(root);

            XmlNode styles = doc.CreateElement("ss","Styles", urn);
            root.AppendChild(styles);
            XmlNode style = doc.CreateElement("ss","Style", urn);
            styles.AppendChild(style);
            NewAttribute(style,"ID","s62",urn);
            XmlNode font = doc.CreateElement("ss","Font",urn);
            style.AppendChild(font);
            NewAttribute(font, "Bold", "1", urn);
            
            XmlNode worksheet = doc.CreateElement("ss", "Worksheet", urn);
            root.AppendChild(worksheet);
            NewAttribute(worksheet, "Name", "Sql Trace", urn);
            XmlNode table = doc.CreateElement("ss", "Table", urn);
            worksheet.AppendChild(table);
            NewAttribute(table, "ExpandedColumnCount",_columns.Count.ToString(CultureInfo.InvariantCulture),urn);

            foreach (ColumnHeader lv in _eventsListView.Columns)
            {
                XmlNode r = doc.CreateElement("ss","Column", urn);
                NewAttribute(r, "AutoFitWidth","0",urn);
                NewAttribute(r, "Width", lv.Width.ToString(CultureInfo.InvariantCulture), urn);
                table.AppendChild(r);
            }

            XmlNode row = doc.CreateElement("ss","Row", urn);
            table.AppendChild(row);
            foreach (ColumnHeader lv in _eventsListView.Columns)
            {
                XmlNode cell = doc.CreateElement("ss","Cell", urn);
                row.AppendChild(cell);
                NewAttribute(cell, "StyleID","s62",urn);
                XmlNode data = doc.CreateElement("ss","Data", urn);
                cell.AppendChild(data);
                NewAttribute(data, "Type","String",urn);
                data.InnerText = lv.Text;
            }

            lock (_cached)
            {
				long rowNumber = 1;
                foreach (ListViewItem lvi in _cached)
                {
                    row = doc.CreateElement("ss", "Row", urn);
                    table.AppendChild(row);
                    for (int i = 0; i < _columns.Count; i++)
                    {
                        PerfColumn pc = _columns[i];
                        if(pc.Column!=-1)
                        {
							XmlNode cell = doc.CreateElement("ss", "Cell", urn);
							row.AppendChild(cell);
							XmlNode data = doc.CreateElement("ss", "Data", urn);
							cell.AppendChild(data);
								string dataType;
								switch (ProfilerEventColumns.ProfilerColumnDataTypes[pc.Column])
								{
										case ProfilerColumnDataType.Int:
										case ProfilerColumnDataType.Long:
											dataType = "Number";
										break;
										case ProfilerColumnDataType.DateTime:
											dataType = "String";
										break;
									default:
											dataType = "String";
										break;
								}
							if (ProfilerEventColumns.EventClass == pc.Column) dataType = "String";
							NewAttribute(data, "Type", dataType, urn);
							if (ProfilerEventColumns.EventClass == pc.Column)
							{
								data.InnerText = GetEventCaption(((ProfilerEvent) (lvi.Tag)));
							}
							else
							{
								data.InnerText = pc.Column == -1
													 ? ""
													 : GetFormattedValue((ProfilerEvent)(lvi.Tag),pc.Column,ProfilerEventColumns.ProfilerColumnDataTypes[pc.Column]==ProfilerColumnDataType.DateTime?pc.Format:"") ??
													   "";
							}
						}
						else
						{
							//The export of the sequence number '#' is handled here.
							XmlNode cell = doc.CreateElement("ss", "Cell", urn);
							row.AppendChild(cell);
							XmlNode data = doc.CreateElement("ss", "Data", urn);
							cell.AppendChild(data);
							const string dataType = "Number";
							NewAttribute(data, "Type", dataType, urn);
							data.InnerText = rowNumber.ToString();
						}
                    }
					rowNumber++;
                }
            }
            using (StringWriter writer = new StringWriter())
            {
                var textWriter = new XmlTextWriter(writer) { Formatting = Formatting.Indented,Namespaces = true};
                doc.Save(textWriter);
                var xml = writer.ToString();
                MemoryStream xmlStream = new MemoryStream();
                xmlStream.Write(System.Text.Encoding.UTF8.GetBytes(xml), 0, xml.Length);
                Clipboard.SetData("XML Spreadsheet", xmlStream);
            }
            MessageBox.Show("Event(s) data copied to clipboard", "Information", MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
        }

		private void mnAbout_Click(object sender, EventArgs e)
		{
			var aboutMsgOrig = String.Format("{0} nhttps://expressprofiler.codeplex.com/ \n Filter Icon: http://www.softicons.com/toolbar-icons/iconza-light-blue-icons-by-turbomilk/filter-icon", VersionString);

			var aboutMsg = new StringBuilder();
			aboutMsg.AppendLine(VersionString + "\nhttps://expressprofiler.codeplex.com/");
			aboutMsg.AppendLine();
			aboutMsg.AppendLine("Filter Icon Downloaded From:");
			aboutMsg.AppendLine("    http://www.softicons.com/toolbar-icons/iconza-light-blue-icons-by-turbomilk/filter-icon");
			aboutMsg.AppendLine("    By Author Turbomilk:  	http://turbomilk.com/");
			aboutMsg.AppendLine("    Used under Creative Commons License: http://creativecommons.org/licenses/by/3.0/");
		
			MessageBox.Show(aboutMsg.ToString(), "About", MessageBoxButtons.OK,
							MessageBoxIcon.Information);
		}

        private void tbStayOnTop_Click(object sender, EventArgs e)
        {
            SetStayOnTop();
        }

        private void SetStayOnTop()
        {
            tbStayOnTop.Checked = !tbStayOnTop.Checked;
            this.TopMost = tbStayOnTop.Checked;
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            SetTransparent();
        }

        private void SetTransparent()
        {
            tbTransparent.Checked = !tbTransparent.Checked;
            this.Opacity = tbTransparent.Checked ? 0.50 : 1;
        }

        private void stayOnTopToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SetStayOnTop();
        }

        private void transparentToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SetTransparent();
        }

        private void deleteSelectedToolStripMenuItem_Click(object sender, EventArgs e)
        {
            for (int i = _eventsListView.SelectedIndices.Count-1; i >= 0; i--)
            {
                _cached.RemoveAt(_eventsListView.SelectedIndices[i]);
            }
            _eventsListView.VirtualListSize = _cached.Count;
            _eventsListView.SelectedIndices.Clear();
        }

        private void keepSelectedToolStripMenuItem_Click(object sender, EventArgs e)
        {
            for (int i = _cached.Count - 1; i >= 0; i--)
            {
                if (!_eventsListView.SelectedIndices.Contains(i))
                {
                    _cached.RemoveAt(i);
                }
            }
            _eventsListView.VirtualListSize = _cached.Count;
            _eventsListView.SelectedIndices.Clear();
        }


		private void SaveToExcelXmlFile()
		{
            var doc = new XmlDocument();
			XmlProcessingInstruction pi = doc.CreateProcessingInstruction("mso-application", "progid='Excel.Sheet'");
			doc.AppendChild(pi);
			const string urn = "urn:schemas-microsoft-com:office:spreadsheet";
			XmlNode root = doc.CreateElement("ss", "Workbook", urn);
			NewAttribute(root, "xmlns:ss", urn);
			doc.AppendChild(root);

			XmlNode styles = doc.CreateElement("ss", "Styles", urn);
			root.AppendChild(styles);
			XmlNode style = doc.CreateElement("ss", "Style", urn);
			styles.AppendChild(style);
			NewAttribute(style, "ID", "s62", urn);
			XmlNode font = doc.CreateElement("ss", "Font", urn);
			style.AppendChild(font);
			NewAttribute(font, "Bold", "1", urn);
            
			XmlNode worksheet = doc.CreateElement("ss", "Worksheet", urn);
			root.AppendChild(worksheet);
			NewAttribute(worksheet, "Name", "Sql Trace", urn);
			XmlNode table = doc.CreateElement("ss", "Table", urn);
			worksheet.AppendChild(table);
			NewAttribute(table, "ExpandedColumnCount", _columns.Count.ToString(CultureInfo.InvariantCulture), urn);

			foreach (ColumnHeader lv in _eventsListView.Columns)
			{
				XmlNode r = doc.CreateElement("ss", "Column", urn);
				NewAttribute(r, "AutoFitWidth", "0", urn);
				NewAttribute(r, "Width", lv.Width.ToString(CultureInfo.InvariantCulture), urn);
				table.AppendChild(r);
			}

			XmlNode row = doc.CreateElement("ss", "Row", urn);
			table.AppendChild(row);
			foreach (ColumnHeader lv in _eventsListView.Columns)
			{
				XmlNode cell = doc.CreateElement("ss", "Cell", urn);
				row.AppendChild(cell);
				NewAttribute(cell, "StyleID", "s62", urn);
				XmlNode data = doc.CreateElement("ss", "Data", urn);
				cell.AppendChild(data);
				NewAttribute(data, "Type", "String", urn);
				data.InnerText = lv.Text;
			}

			lock (_cached)
			{
				long rowNumber = 1;
				foreach (ListViewItem lvi in _cached)
				{
					row = doc.CreateElement("ss", "Row", urn);
					table.AppendChild(row);
					for (int i = 0; i < _columns.Count; i++)
					{
						PerfColumn pc = _columns[i];
						if (pc.Column != -1)
						{
							XmlNode cell = doc.CreateElement("ss", "Cell", urn);
							row.AppendChild(cell);
							XmlNode data = doc.CreateElement("ss", "Data", urn);
							cell.AppendChild(data);
							string dataType;
							switch (ProfilerEventColumns.ProfilerColumnDataTypes[pc.Column])
							{
								case ProfilerColumnDataType.Int:
								case ProfilerColumnDataType.Long:
									dataType = "Number";
									break;
								case ProfilerColumnDataType.DateTime:
									dataType = "String";
									break;
								default:
									dataType = "String";
									break;
							}
							if (ProfilerEventColumns.EventClass == pc.Column) dataType = "String";
							NewAttribute(data, "Type", dataType, urn);
							if (ProfilerEventColumns.EventClass == pc.Column)
							{
								data.InnerText = GetEventCaption(((ProfilerEvent)(lvi.Tag)));
							}
							else
							{
								data.InnerText = pc.Column == -1
													 ? ""
													 : GetFormattedValue((ProfilerEvent)(lvi.Tag), pc.Column, ProfilerEventColumns.ProfilerColumnDataTypes[pc.Column] == ProfilerColumnDataType.DateTime ? pc.Format : "") ??
													   "";
							}
						}
						else
						{
							//The export of the sequence number '#' is handled here.
							XmlNode cell = doc.CreateElement("ss", "Cell", urn);
							row.AppendChild(cell);
							XmlNode data = doc.CreateElement("ss", "Data", urn);
							cell.AppendChild(data);
							const string dataType = "Number";
							NewAttribute(data, "Type", dataType, urn);
							data.InnerText = rowNumber.ToString();
						}
					}
					rowNumber++;
				}
			}

			var saveFileDialog = new SaveFileDialog();
			saveFileDialog.Filter = "Excel XML|*.xml";
			saveFileDialog.Title = "Save the Excel XML FIle";
			saveFileDialog.ShowDialog();

			if (!string.IsNullOrEmpty(saveFileDialog.FileName))
			{
				using (StringWriter writer = new StringWriter())
				{
					var textWriter = new XmlTextWriter(writer)
					{
						Formatting = Formatting.Indented,
						Namespaces = true
					};
					doc.Save(textWriter);
					var xml = writer.ToString();
					MemoryStream xmlStream = new MemoryStream();
					xmlStream.Write(System.Text.Encoding.UTF8.GetBytes(xml), 0, xml.Length);
					xmlStream.Position = 0;
					FileStream fs = new FileStream(saveFileDialog.FileName, FileMode.Create, FileAccess.Write);
					xmlStream.WriteTo(fs);
					fs.Close();
					xmlStream.Close();
				}
				MessageBox.Show(string.Format("File saved to: {0}", saveFileDialog.FileName), "Information", MessageBoxButtons.OK,
					MessageBoxIcon.Information);
			}
		}
        
	    private void SetFilterEvents()
	    {
		    if (_cachedUnFiltered.Count == 0)
		    {
			    _eventsListView.SelectedIndices.Clear();
			    TraceProperties.TraceSettings ts = _currentSettings.GetCopy();
			    using (TraceProperties frm = new TraceProperties())
			    {
				    frm.SetSettings(ts);
				    if (DialogResult.OK != frm.ShowDialog()) return;
				    ts = frm.m_currentsettings.GetCopy();

				    _cachedUnFiltered.AddRange(_cached);
				    _cached.Clear();
				    foreach (ListViewItem lvi in _cachedUnFiltered)
				    {
					    if (frm.IsIncluded(lvi) && _cached.Count < ts.Filters.MaximumEventCount)
					    {
						    _cached.Add(lvi);
					    }
				    }
			    }

			    _eventsListView.VirtualListSize = _cached.Count;
			    UpdateSourceBox();
			    ShowSelectedEvent();
		    }
	    }
        
	    private void ClearFilterEvents()
	    {
		    if (_cachedUnFiltered.Count > 0)
		    {
			    _cached.Clear();
			    _cached.AddRange(_cachedUnFiltered);
			    _cachedUnFiltered.Clear();
			    _eventsListView.VirtualListSize = _cached.Count;
			    _eventsListView.SelectedIndices.Clear();
			    UpdateSourceBox();
			    ShowSelectedEvent();
		    }
	    }
        
		private void saveAllEventsToExcelXmlFileToolStripMenuItem_Click(object sender, EventArgs e)
		{
			SaveToExcelXmlFile();
		}

		/// <summary>
		/// Persist the server string when it changes.
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void edServer_TextChanged(object sender, EventArgs e)
		{
			_servername = edServer.Text;
			SaveDefaultSettings();
		}
        
		/// <summary>
		/// Persist the user name string when it changes.
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void edUser_TextChanged(object sender, EventArgs e)
		{
			_username = edUser.Text;
			SaveDefaultSettings();
		}
        
		private void filterCapturedEventsToolStripMenuItem_Click(object sender, EventArgs e)
		{
			SetFilterEvents();
		}

		private void clearCapturedFiltersToolStripMenuItem_Click(object sender, EventArgs e)
		{
			ClearFilterEvents();
		}

        private void tbFilterEvents_Click(object sender, EventArgs e)
		{
			ToolStripButton filterButton = (ToolStripButton)sender;
			if (filterButton.Checked)
			{
				SetFilterEvents();
			}
			else
			{
				ClearFilterEvents();
			}
		}
    }
}
