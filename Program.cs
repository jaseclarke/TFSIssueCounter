using CommandLine;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Server;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace TfsIssueCounter
{
    [Verb("stats",HelpText = "Produce stats on what issues were opened or closed.")]
    public class StatsDisplayOptions
    {
        [Option('o', "output-file", Required = false, Default = "", HelpText = "Output file to be generated.")]
        public string OutputFilePath { get; set; }

        [Option('c',"closed-states",Required = true, HelpText = "Resolved States for Issues.")]
        public IEnumerable<string> ClosedStates { get; set; }

        [Option('s', "start-date", Required = true,  HelpText = "Earliest Date to Process.")]
        public DateTime StartDate { get; set; }

        [Option("tfs", Required = false, Default = "", HelpText = "Earliest Date to Process.")]
        public string TFSServer { get; set; }

        [Option("project", Required = false, Default = "", HelpText = "Earliest Date to Process.")]
        public string TFSProject { get; set; }

        public string ClosedStateClause
        {
            get
            {
                var quoted = ClosedStates.Select(item => $"'{item}'");
                return $"({string.Join(",", quoted)})";
            }
        }


        [Option('w', "work-item-kinds", Required = true, HelpText = "Resolved States for Issues.")]
        public IEnumerable<string> WorkItemKinds { get; set; }

        public string WorkItemKindClause
        {
            get
            {
                var quoted = WorkItemKinds.Select(item => $"'{item}'");
                return $"({string.Join(",",quoted)})";
            }
        }

    }

    [Verb("states", HelpText = "Shows State transitions for one issue.")]
    public class IssueDisplayOptions
    {
        [Option('i', "work-item-id", Required = true,  HelpText = "Issue Number.")]
        public int WorkItemId { get; set; }

        [Option("tfs", Required = false, Default = "", HelpText = "Earliest Date to Process.")]
        public string TFSServer { get; set; }

        [Option("project", Required = false, Default = "", HelpText = "Earliest Date to Process.")]
        public string TFSProject { get; set; }
    }

    public class TFSBugCounter
    {
        private TfsTeamProjectCollection _tfs = null;
        private ProjectInfo _project = null;
        public List<IssueState> _states = new List<IssueState>();

        private IEnumerable<string> ClosedStates { get; set; }
        private IEnumerable<string> WorkItemKinds { get; set; }

        public string OutputFilePath { get; set; }

        public bool Connected { get; set; }

        private string _tfsServerUrl { get; set; }
        private string _tfsProject { get; set; }

        private int _tfsIssueId { get; set; }

        public TFSBugCounter(StatsDisplayOptions options)
        {
            //_options = options;
            ClosedStates = options.ClosedStates;
            WorkItemKinds = options.WorkItemKinds;
            OutputFilePath = options.OutputFilePath.Trim();
            _tfsServerUrl = options.TFSServer;
            _tfsProject = options.TFSProject;

            Connected = false;
        }
        public TFSBugCounter(IssueDisplayOptions options)
        {
            _tfsServerUrl = options.TFSServer;
            _tfsProject = options.TFSProject;
            _tfsIssueId = options.WorkItemId;

            Connected = false;
        }

        public class IssueState
        {
            public int ID { get; set; }
            public DateTime ClosedDate { get; set; }
            public string FinalState { get; set; }
            public DateTime CreatedDate { get; set; }
        }

        public string WorkItemKindClause
        {
            get
            {
                var quoted = WorkItemKinds.Select(item => $"'{item}'");
                return $"({string.Join(",", quoted)})";
            }
        }

        public string ClosedStateClause
        {
            get
            {
                var quoted = ClosedStates.Select(item => $"'{item}'");
                return $"({string.Join(",", quoted)})";
            }
        }

        public void Connect()
        {
            // Use the TFS Project Picker if no command line options are selected
            if (string.IsNullOrWhiteSpace(_tfsServerUrl))
            {
                using (TeamProjectPicker tpp = new TeamProjectPicker())
                {
                    if (tpp.ShowDialog() == DialogResult.OK)
                    {
                        _tfs = tpp.SelectedTeamProjectCollection;
                        _tfs.EnsureAuthenticated();

                        _project = tpp.SelectedProjects[0];

                        Console.WriteLine($"Connected to TFS Project: {_project.Name}");
                    }

                    Connected = _project != null;
                }
            }
            else // otherwise connect using the command line parameters.
            {
                _tfs = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri(_tfsServerUrl));

                _tfs.EnsureAuthenticated();

                var structureService = _tfs.GetService<ICommonStructureService>();
                _project = structureService.GetProjectFromName(_tfsProject);

                Connected = _project != null;
            }
        }

        public int CountAllOpenItemsAtDate(DateTime date)
        {
            string query = $"SELECT [System.CreatedDate], [System.State] FROM WorkItems WHERE [System.TeamProject] = '{_project.Name}' AND [System.State] NOT IN {ClosedStateClause} AND [System.WorkItemType] IN {WorkItemKindClause} AND[System.CreatedDate] <= '{date.ToString("dd-MMM-yyyy")}'";

            var store = _tfs.GetService<WorkItemStore>();

            var wic = store.Query(query);

            int count = wic.Count;

            Console.WriteLine($"Total Number of Open Items on {date:dd-MMM-yyyy} = {count}");

            return count;
        }

        public void GenerateStatistics(DateTime date, string outputFile)
        {
            int openCount = CountAllOpenItemsAtDate(date);

            HashSet<DateTime> allDates = new HashSet<DateTime>();

            allDates.UnionWith(_states.Select(state => state.CreatedDate));
            allDates.UnionWith(_states.Where(state => state.ClosedDate > DateTime.MinValue).Select(state => state.ClosedDate));

            using (var sw = (string.IsNullOrWhiteSpace(outputFile) ? new StreamWriter(Console.OpenStandardOutput()) : new StreamWriter(outputFile)))
            {
                foreach (var issueDate in allDates.ToList().OrderBy(d => d))
                {
                    int openedCount = _states.Where(s => s.CreatedDate == issueDate).Count();
                    int closedCount = _states.Where(s => s.ClosedDate == issueDate).Count();

                    openCount += openedCount;
                    openCount -= closedCount;

                    sw.WriteLine($"{issueDate.ToShortDateString()},{openedCount},{closedCount},{openCount}");
                }
            }
        }

        public void DisplayStateTransitions(int workItemId)
        {
            var store = _tfs.GetService<WorkItemStore>();

            var wi = store.GetWorkItem(workItemId);

            foreach (Revision rev in wi.Revisions)
            {
                string newState = rev.Fields["System.State"].Value.ToString();
                string originalState = rev.Fields["System.State"].OriginalValue.ToString();
                if (newState != originalState)
                {
                    DateTime revisionDate = (DateTime)rev.Fields["System.ChangedDate"].Value;
                    Console.WriteLine($"{wi.Id}, State Change: {rev.Fields["System.State"].OriginalValue} => {rev.Fields["System.State"].Value} on {revisionDate.ToShortDateString()}");
                }
            }
        }

        public void FindAllItemsNewerThan(DateTime startDate)
        {
            _states.Clear();

            string query = $"SELECT [System.CreatedDate], [System.State] FROM WorkItems WHERE [System.TeamProject] = '{_project.Name}' AND [System.State] NOT IN ('Removed') AND [System.WorkItemType] IN ('Product Backlog Item','Bug') AND [System.CreatedDate] > '{startDate.ToString("dd-MMM-yyyy")}'";

            var store = _tfs.GetService<WorkItemStore>();

            var wic = store.Query(query);

            foreach (WorkItem wi in wic)
            {
                IssueState state = new IssueState { CreatedDate = wi.CreatedDate.Date, FinalState = wi.State, ClosedDate = DateTime.MinValue };

                if (ClosedStates.Contains(wi.State))
                {
                    DateTime lastTransitionToDone = DateTime.MinValue;
                    foreach (Revision rev in wi.Revisions)
                    {
                        string newState = rev.Fields["System.State"].Value.ToString();
                        string originalState = rev.Fields["System.State"].OriginalValue.ToString();
                        if (newState != originalState && newState == state.FinalState)
                        {
                            DateTime revisionDate = (DateTime)rev.Fields["System.ChangedDate"].Value;
                            Console.WriteLine($"{wi.Id}, State Change: {rev.Fields["System.State"].OriginalValue} => {rev.Fields["System.State"].Value} on {revisionDate.ToShortDateString()}");
                            lastTransitionToDone = revisionDate.Date;
                        }
                    }
                    if (lastTransitionToDone > DateTime.MinValue)
                    {
                        state.ClosedDate = lastTransitionToDone;
                    }
                }
                _states.Add(state);
            }
        }
    }
    class Program
    {

        static void Main(string[] args)
        {
            StatsDisplayOptions options = null;
            IssueDisplayOptions issueOptions = null;
            Parser.Default.ParseArguments<StatsDisplayOptions, IssueDisplayOptions>(args).WithParsed<StatsDisplayOptions>(opts => options = opts).WithParsed<IssueDisplayOptions>(opts => issueOptions = opts);

            if (options != null)
            {
                Console.WriteLine($"Closed Work Item States : {options.ClosedStateClause}");
                Console.WriteLine($"Work Item Kinds : {options.WorkItemKindClause}");
                Console.WriteLine($"Start Date : {options.StartDate.ToShortDateString()}");

                var bc = new TFSBugCounter(options);
                bc.Connect();

                if (bc.Connected)
                {
                    DateTime countDate = options.StartDate.Date;
                    bc.CountAllOpenItemsAtDate(DateTime.Now);
                    bc.FindAllItemsNewerThan(countDate);
                    bc.GenerateStatistics(countDate, options.OutputFilePath);
                }
            }
            else if (issueOptions != null)
            {
                Console.WriteLine($"State Transitions for Issue {issueOptions.WorkItemId}");
                var bc = new TFSBugCounter(issueOptions);
                bc.Connect();
                bc.DisplayStateTransitions(issueOptions.WorkItemId);
            }
        }
    }
}
