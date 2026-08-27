using SolidEdgeAdd_In.Utils;

namespace SolidEdgeAdd_In.Processors
{
    public class RenamePartNumberProcessor : IDisposable
    {
        private readonly SeAssembly _assembly;
        private readonly SeApp _application;

        private string _assemblyPath;
        private string _projectDirectory;

        private List<SeOccurrence> _occurrencesToProcess;

        readonly StringBuilder _feedback;

        public RenamePartNumberProcessor(SeAssembly assembly)
        {
            _assembly = assembly;
            _application = _assembly.Application;
            _feedback = new StringBuilder();
        }

        public bool Initialize()
        {
            _assemblyPath = _assembly.FullName;
            _projectDirectory = Path.GetDirectoryName(_assemblyPath);

            if (!IsLoaded_SelectedSet())
            {
                return false;
            }

            return true;
        }

        public void Process()
        {
            bool isReplaceAll = DialogUtils.IsReplaceAll();
            _feedback.AppendLine("==========================================   Co zostało zrobione:   ==========================================");
            _feedback.AppendLine(" ");

            HashSet<string> processedPaths = new(StringComparer.OrdinalIgnoreCase);
            List<(string DftPath, string OldPath, string NewPath, bool GenDoc)> draftsToUpdate = new();

            _feedback.AppendLine("--------------------------------------------------");
            _feedback.AppendLine("1.Tworzenie kopii oraz zmiany w drzewie złożenia: ");

            int i = 0;
            foreach (SeOccurrence occurrence in _occurrencesToProcess)
            {             
                SeOccurrence tempOcc = occurrence;
                try
                {
                    string occurrencePath = occurrence.OccurrenceFileName;
                    if (processedPaths.Contains(occurrencePath))
                    {
                        continue;
                    }
                    processedPaths.Add(occurrencePath);

                    i++;
                    _feedback.AppendLine(" ");
                    _feedback.AppendLine($"     1.{i}:");

                    string occurrenceName = Path.GetFileNameWithoutExtension(occurrencePath);
                    string occurrenceExtension = Path.GetExtension(occurrencePath);

                    if (!File.Exists(occurrencePath))
                    {
                        _feedback.AppendLine($"     [POMINIĘTO] {occurrenceName} -> Brak pliku w folderze projektu.");
                        continue;
                    }

                    (bool isConfirmed, string newPartNumber) = DialogUtils.GetNewPartNumber(occurrenceName);

                    if (!isConfirmed)
                    {
                        _feedback.AppendLine($"     [POMINIĘTO] - {occurrenceName} -> Anulowano przez użytkownika w procesie podania nowej nazwy.");
                        continue;
                    }

                    if (newPartNumber == occurrenceName)
                    {
                        _feedback.AppendLine($"     [POMINIĘTO] - {occurrenceName} -> Nie ma czego podmieniać, bo nowa nazwa jest taka sama jak stara.");
                        continue;
                    }

                    if (string.IsNullOrWhiteSpace(newPartNumber))
                    {
                        _feedback.AppendLine($"     [POMINIĘTO] - {occurrenceName} -> Brak podania nazwy.");
                        continue;
                    }

                    string newOccurrencePath = Path.Combine(_projectDirectory, newPartNumber + occurrenceExtension);
                    string dftPath = Path.Combine(_projectDirectory, occurrenceName + ".dft");
                    string newDftPath = Path.Combine(_projectDirectory, newPartNumber + ".dft");

                    if (!File.Exists(dftPath))
                    {
                        if (!Helpers.IsMessageAccepted($"   Brakuje rysunku (.dft) dla tej części - {occurrenceName} w folderze projektu"))
                        {
                            _feedback.AppendLine($"     [POMINIĘTO] - {occurrenceName} -> Anulowano przez użytkownika, bo brakuje rysunku (.dft).");
                            continue;
                        }
                        _feedback.AppendLine($"     Brak rysunku (.dft). Nie można wykonać kopii.");
                    }
                    else
                    {
                        File.Copy(dftPath, newDftPath, true);
                        _feedback.AppendLine($"     Utworzono kopię pliku (.dft). Ścieżka: {newDftPath}.");
                    }

                    File.Copy(occurrencePath, newOccurrencePath, true);
                    _feedback.AppendLine($"     Utworzono kopię pliku ({occurrenceExtension}). Ścieżka: {newOccurrencePath}.");

                    occurrence.Replace(newOccurrencePath, isReplaceAll, Missing.Value);
                    _feedback.AppendLine($"     Zmieniono wystąpienie w drzewie złożenia: {occurrenceName} ---> {newPartNumber}.");

                    string oldPdfPath = Path.Combine(_projectDirectory, occurrenceName + ".pdf");
                    string oldDxfPath = Path.Combine(_projectDirectory, occurrenceName + ".dxf");
                    bool isGenerateNewDocumentation = false;
                    if (File.Exists(oldPdfPath) || File.Exists(oldDxfPath))
                    {
                        isGenerateNewDocumentation = DialogUtils.IsGenerateNewDocumentation();
                    }

                    if (File.Exists(newDftPath))
                    {
                        draftsToUpdate.Add((newDftPath, occurrencePath, newOccurrencePath, isGenerateNewDocumentation));
                    }
                }
                catch (Exception ex)
                {
                    _feedback.AppendLine($"     [BŁĄD] {ex.Message}");
                    continue;
                }
                finally
                {
                    Helpers.ReleaseCom(ref tempOcc);
                }
            }

            _feedback.AppendLine(" ");
            _feedback.AppendLine("----------------------------------------------------------------------------");
            _feedback.AppendLine("2. Aktualizacja rysunków oraz generowanie nowej dokumentacji (.pdf i .dxf): ");

            int j = 0;
            foreach (var (dftPath, oldPath, newPath, isGenDoc) in draftsToUpdate)
            {
                j++;
                _feedback.AppendLine(" ");
                _feedback.AppendLine($"     2.{j}:");
                UpdateDrawingLinks(dftPath, oldPath, newPath, isGenDoc);
            }

            DisplayFeedback();
        }

        private void UpdateDrawingLinks(string dftPath, string oldFilePath, string newFilePath, bool isGenerateNewDocumentation)
        {
            SeDocuments documents = null;
            SeDraft draft = null;
            SeModelLinks modelLinks = null;

            try
            {
                documents = _application.Documents;
                draft = documents.Open(dftPath);

                if (draft == null)
                {
                    return;
                }

                modelLinks = draft.ModelLinks;

                if (modelLinks != null && modelLinks.Count > 0)
                {
                    for (int i = 1; i <= modelLinks.Count; i++)
                    {
                        SeModelLink modelLink = null;
                        try
                        {
                            modelLink = modelLinks.Item(i);

                            if (!modelLink.FileName.Equals(oldFilePath, StringComparison.OrdinalIgnoreCase))
                            {
                                continue;
                            }

                            modelLink.ChangeSource(newFilePath);
                        }
                        finally
                        {
                            Helpers.ReleaseCom(ref modelLink);
                        }
                    }
                }

                draft.Save();

                _feedback.AppendLine($"     Zaktualizowano rysunek (.dft). Ścieżka: {dftPath}.");

                if (!isGenerateNewDocumentation)
                {
                    return;
                }

                string newPdfPath = Path.ChangeExtension(dftPath, ".pdf");
                draft.SaveAs(newPdfPath, 14);
                _feedback.AppendLine($"     Wyeksportowano (.pdf). Ścieżka: {newPdfPath}.");

                string newDxfPath = Path.ChangeExtension(dftPath, ".dxf");
                draft.SaveAs(newDxfPath, 5);
                _feedback.AppendLine($"     Wyeksportowano (.dxf). Ścieżka: {newDxfPath}.");
            }
            finally
            {
                Helpers.ReleaseCom(ref modelLinks);

                try
                {
                    draft?.Close(false);
                }
                catch { }

                Helpers.ReleaseCom(ref draft);
                Helpers.ReleaseCom(ref documents);
            }
        }

        public bool IsLoaded_SelectedSet()
        {
            _occurrencesToProcess = new List<SeOccurrence>();
            SelectSet selectSet = null;

            try
            {
                selectSet = _assembly.SelectSet;

                if (selectSet.Count == 0)
                {
                    MessageBox.Show("Brak zaznaczonych elementów.");
                    return false;
                }

                for (int i = 1; i <= selectSet.Count; i++)
                {
                    object selectedItem = null;
                    try
                    {
                        selectedItem = selectSet.Item(i);

                        if (!(selectedItem is SeOccurrence occurrence))
                        {
                            Helpers.ReleaseCom(ref selectedItem);
                            continue;
                        }

                        _occurrencesToProcess.Add(occurrence);
                    }
                    catch
                    {
                        Helpers.ReleaseCom(ref selectedItem);
                    }
                }

                if (_occurrencesToProcess.Count == 0)
                {
                    MessageBox.Show("Zaznaczono elementy, ale nie są wystąpienia w złożeniu.");
                    return false;
                }

                return true;
            }
            finally
            {
                Helpers.ReleaseCom(ref selectSet);
            }
        }

        private void DisplayFeedback()
        {
            if (_feedback.Length == 0)
            {
                return;
            }

            using Form form = new()
            {
                Text = "Raport",
                Width = 1200,
                Height = 600,
                StartPosition = FormStartPosition.CenterScreen
            };

            TextBox textBox = new ()
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill,
                Text = _feedback.ToString(),
                Font = new Font("Consolas", 10)
            };

            form.Controls.Add(textBox);
            form.ShowDialog();
        }

        public void Dispose()
        {       
            if (_occurrencesToProcess != null)
            {
                foreach (var occ in _occurrencesToProcess)
                {
                    SeOccurrence tempOcc = occ;
                    if (tempOcc != null)
                    {
                        Helpers.ReleaseCom(ref tempOcc);
                    }
                }
                _occurrencesToProcess.Clear();
            }
        }
    }
}
