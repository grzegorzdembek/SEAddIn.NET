namespace SolidEdgeAdd_In.Ribbons
{
    public sealed class RibbonController : IDisposable, ISEAddInEventsEx
    {
        private readonly SolidEdgeAddIn _addIn;
        private readonly List<Ribbon> _ribbons = new();
        private bool _disposed = false;
        private readonly Dictionary<IConnectionPoint, int> _connectionPoints = new();

        internal RibbonController(SolidEdgeAddIn addIn)
        {
            _addIn = addIn;
            AdviseSink<ISEAddInEventsEx>(_addIn.AddInEx);
        }

        public void Add(Ribbon ribbon, Guid environmentCategory, bool firstTime)
        {
            if (ribbon == null) throw new ArgumentNullException("ribbon");
            ribbon.EnvironmentCategory = environmentCategory;

            foreach (var tab in ribbon.Tabs)
            {
                foreach (var group in tab.Groups)
                {
                    var controlsInGroup = group.Controls.ToList();
                    int controlCount = controlsInGroup.Count;

                    if (controlCount == 0) continue;

                    Array commandNames = Array.CreateInstance(typeof(string), controlCount);
                    Array commandIDs = Array.CreateInstance(typeof(int), controlCount);

                    for (int i = 0; i < controlCount; i++)
                    {
                        var control = controlsInGroup[i];
                        commandNames.SetValue($"{_addIn.Guid}_{control.CommandId}\n{control.Label}\n{control.SuperTip}\n{control.ScreenTip}", i);
                        commandIDs.SetValue(control.CommandId, i);
                    }

                    try
                    {
                        _addIn.AddInEx?.SetAddInInfoEx(
                                _addIn.NativeResourcesDllPath,
                                environmentCategory.ToString("B"),
                                $"{tab.Name}\n{group.Name}",
                                -1, 
                                -1, -1, -1,
                                controlCount,
                                ref commandNames,
                                ref commandIDs);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"{ex.Message}");
                    }
                }
            }
            _ribbons.Add(ribbon);
        }

        public IEnumerable<Ribbon> Ribbons => _ribbons.AsEnumerable();

        #region ISEAddInEventsEx Implementation

        public void OnCommand(int CommandID)
        {
            try
            {
                foreach (var ribbon in _ribbons)
                {
                    var control = ribbon.GetControl<RibbonControl>(CommandID);
                    if (control != null)
                    {
                        control.DoClick();
                        return;
                    }
                }
            }
            catch (Exception ex) 
            { 
                MessageBox.Show(ex.Message); 
            }
        }

        public void OnCommandHelp(int hFrameWnd, int HelpCommandID, int CommandID){}

        public void OnCommandUpdateUI(int CommandID, ref int CommandFlags, out string MenuItemText, ref int BitmapID)
        {
            MenuItemText = null;
            CommandFlags = 1; 
        }

        public void OnCommandOnLineHelp(int HelpCommandID, int CommandID, out string HelpURL){HelpURL = null;}

        #endregion

        #region Connection Point Handling

        private void AdviseSink<T>(object container)
        {
            try
            {
                IConnectionPointContainer cpc = (IConnectionPointContainer)container;
                IConnectionPoint cp = null;
                Guid iid = typeof(T).GUID;
                cpc.FindConnectionPoint(ref iid, out cp);

                if (cp != null)
                {
                    int cookie;
                    cp.Advise(this, out cookie);
                    _connectionPoints.Add(cp, cookie);
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void UnadviseAllSinks()
        {
            foreach (var kvp in _connectionPoints)
            {
                try
                {
                    kvp.Key.Unadvise(kvp.Value);
                    Marshal.ReleaseComObject(kvp.Key); 
                }
                catch { }
            }
            _connectionPoints.Clear();
        }

        #endregion

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            UnadviseAllSinks();

            foreach (var ribbon in _ribbons)
            {
                ribbon.Dispose();
            }
            _ribbons.Clear();
        }
    }
}