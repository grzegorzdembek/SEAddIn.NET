using System;
using System.Collections.Generic;
using System.IO;

namespace SolidEdgeAdd_In.Utils
{
    public class PropertyProvider
    {
        public static string GetType(string filePath)
        {
            SolidEdgeFileProperties.PropertySets propertySets = null;
            SolidEdgeFileProperties.Properties properties = null;
            SolidEdgeFileProperties.Property property = null;

            try
            {
                propertySets = new SolidEdgeFileProperties.PropertySets();
                propertySets.Open(filePath, false);
                properties = (SolidEdgeFileProperties.Properties)propertySets["Custom"];
                property = (SolidEdgeFileProperties.Property)properties["Typ"];
                string type = (string)property.Value;
                return type;
            }
            catch { return null; }
            finally
            {
                propertySets?.Close();
                CoreUtils.ReleaseCom(ref property);
                CoreUtils.ReleaseCom(ref properties);
                CoreUtils.ReleaseCom(ref propertySets);
            }
        }

        public static int GetStatus(string filePath)
        {
            SolidEdgeFileProperties.PropertySets propertySets = null;
            SolidEdgeFileProperties.Properties properties = null;
            SolidEdgeFileProperties.Property property = null;
            try
            {
                propertySets = new SolidEdgeFileProperties.PropertySets();
                propertySets.Open(filePath, false);
                properties = (SolidEdgeFileProperties.Properties)propertySets["ExtendedSummaryInformation"];
                property = (SolidEdgeFileProperties.Property)properties["Status"];
                return (int)property.Value;
            }
            catch { return -1; }
            finally
            {
                propertySets?.Close();
                CoreUtils.ReleaseCom(ref property);
                CoreUtils.ReleaseCom(ref properties);
                CoreUtils.ReleaseCom(ref propertySets);
            }
        }

        public static string GetThickness(string filePath)
        {
            SolidEdgeFileProperties.PropertySets propertySets = null;
            SolidEdgeFileProperties.Properties properties = null;
            SolidEdgeFileProperties.Property property = null;
            try
            {
                propertySets = new SolidEdgeFileProperties.PropertySets();
                propertySets.Open(filePath, false);
                properties = (SolidEdgeFileProperties.Properties)propertySets["Custom"];
                property = (SolidEdgeFileProperties.Property)properties["Grubość materiału"];
                string thickness = (string)property.Value;
                string thicknessNew = null;

                if (thickness != null)
                {
                    thickness = thickness.Replace("mm", "").Replace(" ", "").Trim();
                    thickness = thickness.Replace('.', ',');
                    if (thickness.Contains(","))
                    {
                        thickness = thickness.TrimEnd('0').TrimEnd(',');
                    }
                    thicknessNew = thickness.Replace(',', '_');
                }

                return thicknessNew;
            }
            catch { return null; }
            finally
            {
                propertySets?.Close();
                CoreUtils.ReleaseCom(ref property);
                CoreUtils.ReleaseCom(ref properties);
                CoreUtils.ReleaseCom(ref propertySets);
            }
        }

        public static string GetThickness(SolidEdgeFramework.SolidEdgeDocument document)
        {
            SolidEdgeFramework.PropertySets propertySets = null;
            SolidEdgeFramework.Properties properties = null;
            SolidEdgeFramework.Property property = null;
            try
            {
                propertySets = (SolidEdgeFramework.PropertySets)document.Properties;
                properties = (SolidEdgeFramework.Properties)propertySets.Item("Custom");
                property = (SolidEdgeFramework.Property)properties.Item("Grubość materiału");

                dynamic dynProperty = property;
                object rawValue = dynProperty.Value;

                if (rawValue != null)
                {
                    string thickness = rawValue.ToString().Replace("mm", "").Replace(" ", "").Trim();
                    thickness = thickness.Replace('.', ',');
                    if (thickness.Contains(","))
                    {
                        thickness = thickness.TrimEnd('0').TrimEnd(',');
                    }
                    return thickness.Replace(',', '_');
                }

                return null;
            }
            catch { return null; }
            finally
            {
                CoreUtils.ReleaseCom(ref property);
                CoreUtils.ReleaseCom(ref properties);
                CoreUtils.ReleaseCom(ref propertySets);
            }
        }

        public static string GetMaterial(string filePath)
        {
            SolidEdgeFileProperties.PropertySets propertySets = null;
            SolidEdgeFileProperties.Properties properties = null;
            SolidEdgeFileProperties.Property property = null;

            string dllPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string basePath = System.IO.Path.GetDirectoryName(dllPath);
            string materials = System.IO.Path.Combine(basePath, "materialy.txt");
            List<string> tablicaOld = new List<string>();
            List<string> tablicaNew = new List<string>();
            try
            {
                propertySets = new SolidEdgeFileProperties.PropertySets();
                propertySets.Open(filePath, false);
                properties = (SolidEdgeFileProperties.Properties)propertySets["MechanicalModeling"];
                property = (SolidEdgeFileProperties.Property)properties["Material"];

                foreach (var line in File.ReadLines(materials))
                {
                    if (!string.IsNullOrWhiteSpace(line) && line.Contains(">"))
                    {
                        string[] parts = line.Split('>');
                        if (parts.Length == 2)
                        {
                            tablicaOld.Add(parts[0].Trim());
                            tablicaNew.Add(parts[1].Trim());
                        }
                    }
                }

                string material = (string)property.Value;
                string materialNew = null;
                if (material != null)
                {
                    for (int i = 0; i < tablicaOld.Count; i++)
                    {
                        if (tablicaOld[i] == material) materialNew = tablicaNew[i];
                    }
                }

                return materialNew;
            }
            catch { return null; }
            finally
            {
                propertySets?.Close();
                CoreUtils.ReleaseCom(ref property);
                CoreUtils.ReleaseCom(ref properties);
                CoreUtils.ReleaseCom(ref propertySets);
            }
        }

        public static string GetMaterial(SolidEdgeFramework.SolidEdgeDocument document)
        {
            SolidEdgeFramework.PropertySets propertySets = null;
            SolidEdgeFramework.Properties properties = null;
            SolidEdgeFramework.Property property = null;

            string dllPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string basePath = System.IO.Path.GetDirectoryName(dllPath);
            string materials = System.IO.Path.Combine(basePath, "materialy.txt");
            List<string> tablicaOld = new List<string>();
            List<string> tablicaNew = new List<string>();

            try
            {
                propertySets = (SolidEdgeFramework.PropertySets)document.Properties;
                properties = (SolidEdgeFramework.Properties)propertySets.Item("MechanicalModeling");
                property = (SolidEdgeFramework.Property)properties.Item("Material");

                foreach (var line in File.ReadLines(materials))
                {
                    if (!string.IsNullOrWhiteSpace(line) && line.Contains(">"))
                    {
                        string[] parts = line.Split('>');
                        if (parts.Length == 2)
                        {
                            tablicaOld.Add(parts[0].Trim());
                            tablicaNew.Add(parts[1].Trim());
                        }
                    }
                }

                dynamic dynProperty = property;
                object rawValue = dynProperty.Value;
                string materialNew = null;

                if (rawValue != null)
                {
                    string material = rawValue.ToString();
                    for (int i = 0; i < tablicaOld.Count; i++)
                    {
                        if (tablicaOld[i] == material)
                        {
                            materialNew = tablicaNew[i];
                        }
                    }
                }

                return materialNew;
            }
            catch { return null; }
            finally
            {
                CoreUtils.ReleaseCom(ref property);
                CoreUtils.ReleaseCom(ref properties);
                CoreUtils.ReleaseCom(ref propertySets);
            }
        }

        public static string GetMaterialName(SolidEdgeFramework.SolidEdgeDocument document)
        {
            SolidEdgeFramework.PropertySets propertySets = null;
            SolidEdgeFramework.Properties properties = null;
            SolidEdgeFramework.Property property = null;
            try
            {
                propertySets = (SolidEdgeFramework.PropertySets)document.Properties;
                properties = (SolidEdgeFramework.Properties)propertySets.Item("Custom");
                property = (SolidEdgeFramework.Property)properties.Item("material_nazwa");

                dynamic dynProperty = property;
                object rawValue = dynProperty.Value;

                if (rawValue != null)
                {
                    return rawValue.ToString();
                }

                return null;
            }
            catch { return null; }
            finally
            {
                CoreUtils.ReleaseCom(ref property);
                CoreUtils.ReleaseCom(ref properties);
                CoreUtils.ReleaseCom(ref propertySets);
            }
        }

        public static int GetCount(string filePath)
        {
            SolidEdgeFileProperties.PropertySets propertySets = null;
            SolidEdgeFileProperties.Properties properties = null;
            SolidEdgeFileProperties.Property property = null;

            try
            {
                propertySets = new SolidEdgeFileProperties.PropertySets();
                propertySets.Open(filePath, false);
                properties = (SolidEdgeFileProperties.Properties)propertySets["Custom"];
                property = (SolidEdgeFileProperties.Property)properties["Ilość"];
                string stringValue = property.Value?.ToString();
                if (int.TryParse(stringValue, out int count))
                {
                    return count; 
                }
                return 0; 
            }
            catch { return 0; }
            finally
            {
                propertySets?.Close();
                CoreUtils.ReleaseCom(ref property);
                CoreUtils.ReleaseCom(ref properties);
                CoreUtils.ReleaseCom(ref propertySets);
            }
        }

        public static int GetCount(SolidEdgeFramework.SolidEdgeDocument document)
        {
            SolidEdgeFramework.PropertySets propertySets = null;
            SolidEdgeFramework.Properties properties = null;
            SolidEdgeFramework.Property property = null;
            try
            {
                propertySets = (SolidEdgeFramework.PropertySets)document.Properties;
                properties = (SolidEdgeFramework.Properties)propertySets.Item("Custom");
                property = (SolidEdgeFramework.Property)properties.Item("Ilość");

                dynamic dynProperty = property;
                object rawValue = dynProperty.Value;

                if (rawValue != null && int.TryParse(rawValue.ToString(), out int value))
                {
                    return value;
                }
                return 0;
            }
            catch { return 0; }
            finally
            {
                CoreUtils.ReleaseCom(ref property);
                CoreUtils.ReleaseCom(ref properties);
                CoreUtils.ReleaseCom(ref propertySets);
            }
        }

        public static string GetDxfDate(string filePath)
        {
            SolidEdgeFileProperties.PropertySets propertySets = null;
            SolidEdgeFileProperties.Properties properties = null;
            SolidEdgeFileProperties.Property property = null;

            try
            {
                propertySets = new SolidEdgeFileProperties.PropertySets();
                propertySets.Open(filePath, false);
                properties = (SolidEdgeFileProperties.Properties)propertySets["Custom"];
                property = (SolidEdgeFileProperties.Property)properties["DXF"];
                return (string)property.Value;
            }
            catch { return null; }
            finally
            {
                propertySets?.Close();
                CoreUtils.ReleaseCom(ref property);
                CoreUtils.ReleaseCom(ref properties);
                CoreUtils.ReleaseCom(ref propertySets);
            }
        }

        public static string GetDxfDate(SolidEdgeFramework.SolidEdgeDocument document)
        {
            SolidEdgeFramework.PropertySets propertySets = null;
            SolidEdgeFramework.Properties properties = null;
            SolidEdgeFramework.Property property = null;
            try
            {
                propertySets = (SolidEdgeFramework.PropertySets)document.Properties;
                properties = (SolidEdgeFramework.Properties)propertySets.Item("Custom");
                property = (SolidEdgeFramework.Property)properties.Item("DXF");

                dynamic dynProperty = property;
                object rawValue = dynProperty.Value;

                if (rawValue != null)
                {
                    return rawValue.ToString();
                }

                return null;
            }
            catch { return null; }
            finally
            {
                CoreUtils.ReleaseCom(ref property);
                CoreUtils.ReleaseCom(ref properties);
                CoreUtils.ReleaseCom(ref propertySets);
            }
        }

        public static bool HasType(string filePath)
        {
            string type = GetType(filePath);
            return !string.IsNullOrEmpty(type);
        }

        public static bool HasStatus(string filePath)
        {
            int status = GetStatus(filePath);
            return status >= 0;
        }

        public static bool HasThickness(string filePath)
        {
            string thickness = GetThickness(filePath);
            return !string.IsNullOrEmpty(thickness);
        }

        public static bool HasMaterial(string filePath)
        {
            string material = GetMaterial(filePath);
            return !string.IsNullOrEmpty(material);
        }

        public static bool HasCount(string filePath)
        {
            int count = GetCount(filePath);
            return count > 0;
        }

        public static bool HasCount(SolidEdgeFramework.SolidEdgeDocument document)
        {
            int count = GetCount(document);
            return count > 0;
        }

        public static bool HasDxfDate(string filePath)
        {
            string date = GetDxfDate(filePath);
            return !string.IsNullOrEmpty(date);
        }

        public static bool HasDxfDate(SolidEdgeFramework.SolidEdgeDocument document)
        {
            string date = GetDxfDate(document);
            return !string.IsNullOrEmpty(date);
        }

        public static bool IsTypeA
            (string filePath)
        {
            string type = GetType(filePath);
            return type == "A";
        }

        public static bool IsTypeB(string filePath)
        {
            string type = GetType(filePath);
            return type == "B";
        }

        public static bool IsTypeC(string filePath)
        {
            string type = GetType(filePath);
            return type == "C";
        }

        public static bool IsTypeK(string filePath)
        {
            string type = GetType(filePath);
            return type == "K";
        }

        public static bool IsTypeH(string filePath)
        {
            string type = GetType(filePath);
            return type == "H";
        }

        public static bool IsTypeN(string filePath)
        {
            string type = GetType(filePath);
            return type == "N";
        }

        public static bool IsStatusAvailable(string filePath)
        {
            int status = GetStatus(filePath);
            return status == 0;
        }

        public static bool SetCount(string filePath, int count)
        {
            SolidEdgeFileProperties.PropertySets propertySets = null;
            SolidEdgeFileProperties.Properties properties = null;
            SolidEdgeFileProperties.Property property = null;
            bool isSuccess = false;

            try
            {
                propertySets = new SolidEdgeFileProperties.PropertySets();
                propertySets.Open(filePath, false);
                properties = (SolidEdgeFileProperties.Properties)propertySets["Custom"];

                try
                {
                    property = (SolidEdgeFileProperties.Property)properties["Ilość"];
                    property.Value = count;
                }
                catch { property = (SolidEdgeFileProperties.Property)properties.Add("Ilość", count); }

                propertySets.Save();
                isSuccess = true;

            }
            catch { isSuccess = false; }
            finally
            {
                
                propertySets?.Close();
                CoreUtils.ReleaseCom(ref property);
                CoreUtils.ReleaseCom(ref properties);
                CoreUtils.ReleaseCom(ref propertySets);
            }

            return isSuccess;
        }

        public static void SetCount(SolidEdgeFramework.SolidEdgeDocument document, int count)
        {
            SolidEdgeFramework.PropertySets propertySets = null;
            SolidEdgeFramework.Properties properties = null;
            SolidEdgeFramework.Property property = null;
            try
            {
                propertySets = (SolidEdgeFramework.PropertySets)document.Properties;
                properties = (SolidEdgeFramework.Properties)propertySets.Item("Custom");

                for (int i = 1; i <= properties.Count; i++)
                {
                    SolidEdgeFramework.Property tempProp = (SolidEdgeFramework.Property)properties.Item(i);
                    dynamic dynProp = tempProp;
                    if (dynProp.Name == "Ilość")
                    {
                        tempProp.Delete();
                        CoreUtils.ReleaseCom(ref tempProp);
                        break;
                    }
                    CoreUtils.ReleaseCom(ref tempProp);
                }
                property = (SolidEdgeFramework.Property)properties.Add("Ilość", count);
            }
            catch { }
            finally
            {
                CoreUtils.ReleaseCom(ref property);
                CoreUtils.ReleaseCom(ref properties);
                CoreUtils.ReleaseCom(ref propertySets);
            }
        }

        public static void SetDxfDate(string filePath)
        {
            SolidEdgeFileProperties.PropertySets propertySets = null;
            SolidEdgeFileProperties.Properties properties = null;
            SolidEdgeFileProperties.Property property = null;

            try
            {
                propertySets = new SolidEdgeFileProperties.PropertySets();
                propertySets.Open(filePath, false);
                properties = (SolidEdgeFileProperties.Properties)propertySets["Custom"];

                try
                {
                    property = (SolidEdgeFileProperties.Property)properties["DXF"];
                    property.Value = DateTime.Now.ToString("yyyy-MM-dd_HH-mm");
                }
                catch { property = (SolidEdgeFileProperties.Property)properties.Add("DXF", DateTime.Now.ToString("yyyy-MM-dd_HH-mm")); }
                propertySets.Save();
            }
            finally
            {
                propertySets?.Close();
                CoreUtils.ReleaseCom(ref property);
                CoreUtils.ReleaseCom(ref properties);
                CoreUtils.ReleaseCom(ref propertySets);
            }
        }

        public static void SetDxfDate(SolidEdgeFramework.SolidEdgeDocument document)
        {
            SolidEdgeFramework.PropertySets propertySets = null;
            SolidEdgeFramework.Properties properties = null;
            SolidEdgeFramework.Property property = null;
            try
            {
                propertySets = (SolidEdgeFramework.PropertySets)document.Properties;
                properties = (SolidEdgeFramework.Properties)propertySets.Item("Custom");

                for (int i = 1; i <= properties.Count; i++)
                {
                    SolidEdgeFramework.Property tempProp = (SolidEdgeFramework.Property)properties.Item(i);
                    dynamic dynProp = tempProp;
                    if (dynProp.Name == "DXF")
                    {
                        tempProp.Delete();
                        CoreUtils.ReleaseCom(ref tempProp);
                        break;
                    }
                    CoreUtils.ReleaseCom(ref tempProp);
                }

                property = (SolidEdgeFramework.Property)properties.Add("DXF", DateTime.Now.ToString("yyyy-MM-dd_HH-mm"));
            }
            catch { }
            finally
            {
                CoreUtils.ReleaseCom(ref property);
                CoreUtils.ReleaseCom(ref properties);
                CoreUtils.ReleaseCom(ref propertySets);
            }
        }

        public static void SetType(SolidEdgeFramework.SolidEdgeDocument document, string type)
        {
            SolidEdgeFramework.PropertySets propertySets = null;
            SolidEdgeFramework.Properties properties = null;
            SolidEdgeFramework.Property property = null;
            try
            {
                propertySets = (SolidEdgeFramework.PropertySets)document.Properties;
                properties = (SolidEdgeFramework.Properties)propertySets.Item("Custom");

                for (int i = 1; i <= properties.Count; i++)
                {
                    SolidEdgeFramework.Property tempProp = (SolidEdgeFramework.Property)properties.Item(i);
                    dynamic dynProp = tempProp;
                    if (dynProp.Name == "Typ")
                    {
                        tempProp.Delete();
                        CoreUtils.ReleaseCom(ref tempProp);
                        break;
                    }
                    CoreUtils.ReleaseCom(ref tempProp);
                }

                property = (SolidEdgeFramework.Property)properties.Add("Typ", type);
            }
            catch { }
            finally
            {
                CoreUtils.ReleaseCom(ref property);
                CoreUtils.ReleaseCom(ref properties);
                CoreUtils.ReleaseCom(ref propertySets);
            }
        }

        public static void SetSheetMaterial(SolidEdgeFramework.SolidEdgeDocument document)
        {
            SolidEdgeFramework.PropertySets propertySets = null;
            SolidEdgeFramework.Properties properties = null;
            SolidEdgeFramework.Property property = null;
            try
            {
                propertySets = (SolidEdgeFramework.PropertySets)document.Properties;
                properties = (SolidEdgeFramework.Properties)propertySets.Item("Custom");

                for (int i = 1; i <= properties.Count; i++)
                {
                    SolidEdgeFramework.Property tempProp = (SolidEdgeFramework.Property)properties.Item(i);
                    dynamic dynProp = tempProp;
                    if (dynProp.Name == "material_nazwa")
                    {
                        tempProp.Delete();
                        CoreUtils.ReleaseCom(ref tempProp);
                        break;
                    }
                    CoreUtils.ReleaseCom(ref tempProp);
                }

                property = (SolidEdgeFramework.Property)properties.Add("material_nazwa", "Blachy");
            }
            catch { }
            finally
            {
                CoreUtils.ReleaseCom(ref property);
                CoreUtils.ReleaseCom(ref properties);
                CoreUtils.ReleaseCom(ref propertySets);
            }
        }
    }
}
