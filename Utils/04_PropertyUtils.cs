namespace SolidEdgeAdd_In.Utils
{
    public class FileData
    {
        public string Name { get; set; }
        public string Type { get; set; }
        public int Status { get; set; }
        public string Thickness { get; set; }
        public string Material { get; set; }
        public string MaterialName { get; set; }
        public int Count { get; set; }
        public string DxfDate { get; set; }
        public int OccurrenceCount { get; set; }
        public string SizeX { get; set; }
        public string SizeY { get; set; }
        public string Title { get; set; }
        public string Color { get; set; }
        public string Finish { get; set; }
    }

    public class PropertyUtils : IDisposable
    {
        private SeFilePropertySets _filePropertySets = null;
        private SePropertySets _docPropertySets = null;

        private readonly bool _isFileMode;
        private bool _disposed = false;

        private static Dictionary<string, string> _materialTranslations = null;
        private static readonly object _cacheLock = new();

        public PropertyUtils(string filePath, bool readOnly = false)
        {
            _isFileMode = true;
            _filePropertySets = new SeFilePropertySets();
            _filePropertySets.Open(filePath, readOnly);
        }

        public PropertyUtils(SeDocument document)
        {
            _isFileMode = false;
            _docPropertySets = (SePropertySets)document.Properties;
        }

        public string Color => GetPropertyString(Constants.Properties.CustomSet, Constants.Properties.Color);
        public string Finish => GetPropertyString(Constants.Properties.CustomSet, Constants.Properties.Finish);
        public string TitleEng => GetPropertyString(Constants.Properties.SummarySet, Constants.Properties.TitleEng);
        public string TitlePl => GetPropertyString(Constants.Properties.SummarySet, Constants.Properties.TitlePl);

        public string Type
        {
            get => GetPropertyString(Constants.Properties.CustomSet, Constants.Properties.Type);
            set => SetProperty(Constants.Properties.CustomSet, Constants.Properties.Type, value);
        }

        public int Status
        {
            get
            {
                object val = _isFileMode ? GetCustomFileProperty(Constants.Properties.ExtendedSummarySet, Constants.Properties.Status) : GetCustomDocProperty(Constants.Properties.ExtendedSummarySet, Constants.Properties.Status);
                return val != null ? (int)val : -1;
            }
        }

        public string Thickness
        {
            get
            {
                object rawValue = _isFileMode ? GetCustomFileProperty(Constants.Properties.CustomSet, Constants.Properties.Thickness) : GetCustomDocProperty(Constants.Properties.CustomSet, Constants.Properties.Thickness);
                if (rawValue != null)
                {
                    string thickness = rawValue.ToString().Replace("mm", "").Replace(" ", "").Trim();
                    thickness = thickness.Replace('.', ',');
                    if (thickness.Contains(",")) { thickness = thickness.TrimEnd('0').TrimEnd(','); }
                    return thickness.Replace(',', '_');
                }
                return null;
            }
        }

        public string Material
        {
            get
            {
                object rawValue = _isFileMode ? GetCustomFileProperty(Constants.Properties.MechanicalModeling, Constants.Properties.Material) : GetCustomDocProperty(Constants.Properties.MechanicalModeling, Constants.Properties.Material);
                if (rawValue != null)
                {
                    EnsureMaterialsLoaded();
                    string material = rawValue.ToString();
                    if (_materialTranslations.TryGetValue(material, out string translatedMaterial)) { return translatedMaterial; }
                }
                return null;
            }
        }

        public string MaterialName => GetPropertyString(Constants.Properties.CustomSet, Constants.Properties.MaterialName);

        public int Count
        {
            get
            {
                object rawValue = _isFileMode ? GetCustomFileProperty(Constants.Properties.CustomSet, Constants.Properties.Count) : GetCustomDocProperty(Constants.Properties.CustomSet, Constants.Properties.Count);
                return (rawValue != null && int.TryParse(rawValue.ToString(), out int count)) ? count : 0;
            }
            set => SetProperty(Constants.Properties.CustomSet, Constants.Properties.Count, value);
        }

        public string DxfDate => GetPropertyString(Constants.Properties.CustomSet, Constants.Properties.DxfDate);
        public string SizeX => GetPropertyString(Constants.Properties.CustomSet, Constants.Properties.SizeX);
        public string SizeY => GetPropertyString(Constants.Properties.CustomSet, Constants.Properties.SizeY);

        public bool HasType => !string.IsNullOrEmpty(Type);
        public bool HasStatus => Status >= 0;
        public bool HasThickness => !string.IsNullOrEmpty(Thickness);
        public bool HasMaterial => !string.IsNullOrEmpty(Material);
        public bool HasCount => Count > 0;
        public bool HasDxfDate => !string.IsNullOrEmpty(DxfDate);
        public bool IsStatusAvailable => Status == 0;

        public bool IsTypeA => Type == Constants.PartTypes.Assembly;
        public bool IsTypeB => Type == Constants.PartTypes.SheetMetal;
        public bool IsTypeC => Type == Constants.PartTypes.Part;
        public bool IsTypeK => Type == Constants.PartTypes.Steelmaking;
        public bool IsTypeH => Type == Constants.PartTypes.Commercial;
        public bool IsTypeN => Type == Constants.PartTypes.Standard;

        public void UpdateDxfDate() { SetProperty(Constants.Properties.CustomSet, Constants.Properties.DxfDate, DateTime.Now.ToString("yyyy-MM-dd-HH-mm")); }
        public void ClearDxfDate() { SetProperty(Constants.Properties.CustomSet, Constants.Properties.DxfDate, String.Empty); }
        public void SetSheetMaterial() { SetProperty(Constants.Properties.CustomSet, Constants.Properties.MaterialName, "Blachy"); }

        private string GetPropertyString(string setName, string propName)
        {
            object rawValue = _isFileMode ? GetCustomFileProperty(setName, propName) : GetCustomDocProperty(setName, propName);
            return rawValue?.ToString();
        }

        private void SetProperty(string setName, string propName, object value)
        {
            if (_isFileMode) { SetCustomFileProperty(setName, propName, value); }
            else { SetCustomDocProperty(setName, propName, value); }
        }

        private object GetCustomFileProperty(string setName, string propName)
        {
            SeFileProperties properties = null;
            SeFileProperty property = null;
            try
            {
                properties = (SeFileProperties)_filePropertySets[setName];
                property = (SeFileProperty)properties[propName];
                return property.Value;
            }
            catch { return null; }
            finally { Helpers.ReleaseCom(ref property); Helpers.ReleaseCom(ref properties); }
        }

        private void SetCustomFileProperty(string setName, string propName, object value)
        {
            SeFileProperties properties = null;
            SeFileProperty property = null;
            try
            {
                properties = (SeFileProperties)_filePropertySets[setName];
                try { property = (SeFileProperty)properties[propName]; property.Value = value; }
                catch { property = (SeFileProperty)properties.Add(propName, value); }
                _filePropertySets.Save();
            }
            catch { /* */ }
            finally { Helpers.ReleaseCom(ref property); Helpers.ReleaseCom(ref properties); }
        }

        private object GetCustomDocProperty(string setName, string propName)
        {
            SeProperties properties = null;
            SeProperty property = null;
            try
            {
                properties = (SeProperties)_docPropertySets.Item(setName);
                property = (SeProperty)properties.Item(propName);
                dynamic dynProperty = property;
                return dynProperty.Value;
            }
            catch { return null; }
            finally { Helpers.ReleaseCom(ref property); Helpers.ReleaseCom(ref properties); }
        }

        private void SetCustomDocProperty(string setName, string propName, object value)
        {
            SeProperties properties = null;
            SeProperty property = null;
            try
            {
                properties = (SeProperties)_docPropertySets.Item(setName);
                for (int i = 1; i <= properties.Count; i++)
                {
                    SeProperty tempProp = null;
                    try
                    {
                        tempProp = (SeProperty)properties.Item(i);
                        dynamic dynProp = tempProp;
                        if (dynProp.Name == propName) { tempProp.Delete(); break; }
                    }
                    finally { Helpers.ReleaseCom(ref tempProp); }
                }
                property = (SeProperty)properties.Add(propName, value);
            }
            finally { Helpers.ReleaseCom(ref property); Helpers.ReleaseCom(ref properties); }
        }

        private static void EnsureMaterialsLoaded()
        {
            if (_materialTranslations != null) { return; }

            lock (_cacheLock)
            {
                if (_materialTranslations != null) { return; }
                _materialTranslations = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                string dllPath = Assembly.GetExecutingAssembly().Location;
                string basePath = Path.GetDirectoryName(dllPath);
                string materialsFile = Path.Combine(basePath, "materialy.txt");
                if (File.Exists(materialsFile))
                {
                    foreach (var line in File.ReadLines(materialsFile))
                    {
                        if (!string.IsNullOrWhiteSpace(line) && line.Contains(">"))
                        {
                            string[] parts = line.Split('>');
                            if (parts.Length == 2) { _materialTranslations[parts[0].Trim()] = parts[1].Trim(); }
                        }
                    }
                }
            }
        }

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed) return;
            if (disposing)
            {
                if (_isFileMode && _filePropertySets != null)
                {
                    try { _filePropertySets.Close(); } catch { }
                    Helpers.ReleaseCom(ref _filePropertySets);
                }
                else if (!_isFileMode && _docPropertySets != null)
                {
                    Helpers.ReleaseCom(ref _docPropertySets);
                }
            }
            _filePropertySets = null;
            _docPropertySets = null;
            _disposed = true;
        }

        public void Dispose() { Dispose(true); GC.SuppressFinalize(this); }
        ~PropertyUtils() { Dispose(false); }
    }
}