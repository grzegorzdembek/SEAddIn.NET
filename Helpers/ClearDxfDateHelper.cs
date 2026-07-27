using SolidEdgeAdd_In.Utils;

namespace SolidEdgeAdd_In.Helpers
{
    public class ClearDxfDateHelper
    {
        public static void ProcessClearing(SeAssembly assembly)
        {
            SeOccurrences occurrences = null;

            try
            {
                occurrences = assembly.Occurrences; int count = occurrences.Count;

                for (int i = 1; i <= count; i++)
                {
                    SeOccurrence occurrence = null; SeDocument doc = null; 

                    try
                    {
                        occurrence = (SeOccurrence)occurrences.Item(i); if (occurrence.IncludeInBom == false) continue;

                        doc = (SeDocument)occurrence.OccurrenceDocument;

                        string path = null; try { path = doc.FullName; } catch { continue; } if (string.IsNullOrEmpty(path)) continue;

                        using var properties = new PropertyProvider(doc);

                        if (doc is SePart || doc is SeSheetMetal)
                        {
                            if (!properties.IsTypeB) continue; if (!properties.HasMaterial) continue; if (!properties.HasThickness) continue; if (!properties.IsStatusAvailable) continue;

                            if (properties.HasDxfDate) properties.ClearDxfDate();
                        }
                        else if (doc is SeAssembly asmDoc)
                        {
                            bool isTypeA = false; isTypeA = properties.IsTypeA; 

                            if (!isTypeA) continue;

                            ProcessClearing(asmDoc);
                        }
                    }
                    catch { continue; }
                    finally { CoreUtils.ReleaseCom(ref doc); CoreUtils.ReleaseCom(ref occurrence); }
                }
            }
            finally { CoreUtils.ReleaseCom(ref occurrences); }
        }
    }
}