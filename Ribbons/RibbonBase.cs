namespace SolidEdgeAdd_In.Ribbons
{
    public abstract class Ribbon : IDisposable
    {
        private Guid _environmentCategory;

        private readonly List<RibbonTab> _tabs = new ();

        public SolidEdgeFramework.Application Application { get; set; } 

        public RibbonTab AddTab(string name)
        {
            var ribbonTab = new RibbonTab(name);
            _tabs.Add(ribbonTab);
            return ribbonTab;
        }

        public T GetControl<T>(int commandId) where T : RibbonControl
        {
            return Controls.OfType<T>().FirstOrDefault(x => x.CommandId == commandId);
        }

        public IEnumerable<RibbonButton> Buttons => Controls.OfType<RibbonButton>();

        public IEnumerable<RibbonControl> Controls => Tabs.SelectMany(tab => tab.Controls);

        public IEnumerable<RibbonTab> Tabs => _tabs.AsEnumerable();

        public Guid EnvironmentCategory
        {
            get 
            { 
                return _environmentCategory; 
            }
            internal set 
            { 
                _environmentCategory = value; 
            }
        }

        public virtual void OnControlClick(RibbonControl control) 
        {
            
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing) 
        {
            
        }
    }

    [Serializable]
    public delegate void RibbonControlClickEventHandler(RibbonControl control);

    public abstract class RibbonControl
    {
        public event RibbonControlClickEventHandler Click;

        private readonly int _commandId;
        private string _label;
        private string _screenTip;
        private string _superTip;

        internal RibbonControl(int commandId)
        {
            _commandId = commandId;
        }

        public int CommandId => _commandId;
        public string Label 
        { 
            get 
            { 
                return _label; 
            } 
            set 
            { 
                _label = value; 
            } 
        }

        public string ScreenTip 
        { 
            get 
            { 
                return _screenTip; 
            } 
            set 
            { 
                _screenTip = value; 
            } 
        }
        public string SuperTip 
        { 
            get 
            { 
                return _superTip; 
            } 
            set 
            { 
                _superTip = value; 
            } 
        }

        internal virtual void DoClick()
        {
            Click?.Invoke(this);
        }
    }

    public class RibbonButton : RibbonControl
    {
        public RibbonButton(int id) : base(id) 
        {
            
        }

        public RibbonButtonSize Size 
        {   get; set; } = RibbonButtonSize.Normal;

        public string DropDownGroup 
        { 
            get; 
            set; 
        }
        public int ImageId 
        { get; set; } = -1;
    }

    public class RibbonGroup
    {
        private readonly string _name;
        private readonly List<RibbonControl> _controls = new();

        internal RibbonGroup(string name)
        {
            _name = name;
        }

        public void AddControl(RibbonControl control)
        {
            _controls.Add(control);
        }

        public string Name => _name;
        public IEnumerable<RibbonControl> Controls => _controls.AsEnumerable();
    }

    public class RibbonTab
    {
        private readonly string _name;
        private readonly List<RibbonGroup> _groups = new();

        internal RibbonTab(string name)
        {
            _name = name;
        }

        public RibbonGroup AddGroup(string name)
        {
            var ribbonGroup = new RibbonGroup(name);
            _groups.Add(ribbonGroup);
            return ribbonGroup;
        }

        public string Name => _name;
        public IEnumerable<RibbonGroup> Groups => _groups.AsEnumerable();
        public IEnumerable<RibbonControl> Controls => Groups.SelectMany(group => group.Controls);
    }

    public enum RibbonButtonSize
    {
        Normal,
        Large
    }
}