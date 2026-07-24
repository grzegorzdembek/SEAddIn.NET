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
        public bool NeedsDxf { get; set; }
    }

    public class PropertyProvider : IDisposable
    {
        private SeFilePropertySets _filePropertySets = null; private SePropertySets _docPropertySets = null;
        private readonly bool _isFileMode; private bool _disposed = false;
        private static Dictionary<string, string> _materialTranslations = null; private static readonly object _cacheLock = new();

        public PropertyProvider(string filePath, bool readOnly = false) { _isFileMode = true; _filePropertySets = new SeFilePropertySets(); _filePropertySets.Open(filePath, readOnly); }

        public PropertyProvider(SeDocument document) { _isFileMode = false; _docPropertySets = (SePropertySets)document.Properties; }

        public string Type { get => GetPropertyString("Custom", "Typ"); set => SetProperty("Custom", "Typ", value); }

        public int Status { get { object val = _isFileMode ? GetCustomFileProperty("ExtendedSummaryInformation", "Status") : GetCustomDocProperty("ExtendedSummaryInformation", "Status"); return val != null ? (int)val : -1; } }

        public string Thickness
        {
            get
            {
                object rawValue = _isFileMode ? GetCustomFileProperty("Custom", "Grubość materiału") : GetCustomDocProperty("Custom", "Grubość materiału");

                if (rawValue != null)
                {
                    string thickness = rawValue.ToString().Replace("mm", "").Replace(" ", "").Trim(); thickness = thickness.Replace('.', ',');

                    if (thickness.Contains(",")) thickness = thickness.TrimEnd('0').TrimEnd(',');

                    return thickness.Replace(',', '_');
                }

                return null;
            }
        }

        public string Material
        {
            get
            {
                object rawValue = _isFileMode ? GetCustomFileProperty("MechanicalModeling", "Material") : GetCustomDocProperty("MechanicalModeling", "Material");

                if (rawValue != null)
                {
                    EnsureMaterialsLoaded(); string material = rawValue.ToString();

                    if (_materialTranslations.TryGetValue(material, out string translatedMaterial)) return translatedMaterial;
                }

                return null;
            }
        }

        public string MaterialName => GetPropertyString("Custom", "material_nazwa");

        public int Count { get { object rawValue = _isFileMode ? GetCustomFileProperty("Custom", "Ilość") : GetCustomDocProperty("Custom", "Ilość"); return (rawValue != null && int.TryParse(rawValue.ToString(), out int count)) ? count : 0; } set => SetProperty("Custom", "Ilość", value); }

        public string DxfDate => GetPropertyString("Custom", "DXF");
        public string SizeX => GetPropertyString("Custom", "Model_Rozwinięcia_RozmiarArkuszaX");
        public string SizeY => GetPropertyString("Custom", "Model_Rozwinięcia_RozmiarArkuszaY");

        public bool HasType => !string.IsNullOrEmpty(Type); public bool HasStatus => Status >= 0; public bool HasThickness => !string.IsNullOrEmpty(Thickness);
        public bool HasMaterial => !string.IsNullOrEmpty(Material); public bool HasCount => Count > 0; public bool HasDxfDate => !string.IsNullOrEmpty(DxfDate);
        public bool IsStatusAvailable => Status == 0;

        public bool IsTypeA => Type == "A"; public bool IsTypeB => Type == "B"; public bool IsTypeC => Type == "C";
        public bool IsTypeK => Type == "K"; public bool IsTypeH => Type == "H"; public bool IsTypeN => Type == "N";

        public void UpdateDxfDate() { SetProperty("Custom", "DXF", DateTime.Now.ToString("yyyy-MM-dd")); }
        public void ClearDxfDate() { SetProperty("Custom", "DXF", String.Empty); }
        public void SetSheetMaterial() { SetProperty("Custom", "material_nazwa", "Blachy"); }

        private string GetPropertyString(string setName, string propName) { object rawValue = _isFileMode ? GetCustomFileProperty(setName, propName) : GetCustomDocProperty(setName, propName); return rawValue?.ToString(); }

        private void SetProperty(string setName, string propName, object value) { if (_isFileMode) SetCustomFileProperty(setName, propName, value); else SetCustomDocProperty(setName, propName, value); }

        private object GetCustomFileProperty(string setName, string propName)
        {
            SeFileProperties properties = null; SeFileProperty property = null;

            try { properties = (SeFileProperties)_filePropertySets[setName]; property = (SeFileProperty)properties[propName]; return property.Value; }
            catch { return null; }
            finally { CoreUtils.ReleaseCom(ref property); CoreUtils.ReleaseCom(ref properties); }
        }

        private void SetCustomFileProperty(string setName, string propName, object value)
        {
            SeFileProperties properties = null; SeFileProperty property = null;

            try
            {
                properties = (SeFileProperties)_filePropertySets[setName];

                try { property = (SeFileProperty)properties[propName]; property.Value = value; } catch { property = (SeFileProperty)properties.Add(propName, value); }

                _filePropertySets.Save();
            }
            catch { }
            finally { CoreUtils.ReleaseCom(ref property); CoreUtils.ReleaseCom(ref properties); }
        }

        private object GetCustomDocProperty(string setName, string propName)
        {
            SeProperties properties = null; SeProperty property = null;

            try { properties = (SeProperties)_docPropertySets.Item(setName); property = (SeProperty)properties.Item(propName); dynamic dynProperty = property; return dynProperty.Value; }
            catch { return null; }
            finally { CoreUtils.ReleaseCom(ref property); CoreUtils.ReleaseCom(ref properties); }
        }

        private void SetCustomDocProperty(string setName, string propName, object value)
        {
            SeProperties properties = null; SeProperty property = null;

            try
            {
                properties = (SeProperties)_docPropertySets.Item(setName);

                for (int i = 1; i <= properties.Count; i++)
                {
                    SeProperty tempProp = null;

                    try { tempProp = (SeProperty)properties.Item(i); dynamic dynProp = tempProp; if (dynProp.Name == propName) { tempProp.Delete(); break; } }
                    finally { CoreUtils.ReleaseCom(ref tempProp); }
                }

                property = (SeProperty)properties.Add(propName, value);
            }
            catch { }
            finally { CoreUtils.ReleaseCom(ref property); CoreUtils.ReleaseCom(ref properties); }
        }

        private static void EnsureMaterialsLoaded()
        {
            if (_materialTranslations != null) return;

            lock (_cacheLock)
            {
                if (_materialTranslations != null) return;

                _materialTranslations = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                try
                {
                    string dllPath = Assembly.GetExecutingAssembly().Location; string basePath = Path.GetDirectoryName(dllPath); string materialsFile = Path.Combine(basePath, "materialy.txt");

                    if (File.Exists(materialsFile))
                    {
                        foreach (var line in File.ReadLines(materialsFile))
                        {
                            if (!string.IsNullOrWhiteSpace(line) && line.Contains(">")) { string[] parts = line.Split('>'); if (parts.Length == 2) { _materialTranslations[parts[0].Trim()] = parts[1].Trim(); } }
                        }
                    }
                }
                catch { }
            }
        }

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed) return;

            if (disposing)
            {
                if (_isFileMode && _filePropertySets != null) { try { _filePropertySets.Close(); } catch { } CoreUtils.ReleaseCom(ref _filePropertySets); }
                else if (!_isFileMode && _docPropertySets != null) { CoreUtils.ReleaseCom(ref _docPropertySets); }
            }

            _filePropertySets = null; _docPropertySets = null; _disposed = true;
        }

        public void Dispose() { Dispose(true); GC.SuppressFinalize(this); }

        ~PropertyProvider() { Dispose(false); }
    }
}