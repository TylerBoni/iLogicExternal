using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Collections.Concurrent;
using System.Threading;
using Autodesk.iLogic.Interfaces;
using Inventor;
using Microsoft.Win32;
using System.Linq;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace iLogicExternal
{
    /// <summary>
    /// Bridge class that implements the iLogic bridge functionality
    /// </summary>
    public class iLogicBridge
    {
        private bool isOpen = false;
        private dynamic iLogicAuto;
        private Inventor.Application inventorApp;
        private FileSystemWatcher watcher;
        private const string localILogicFolderName = "ilogic"; // Local folder name for iLogic rules
        private ConcurrentDictionary<string, bool> processingFiles = new ConcurrentDictionary<string, bool>();
        private bool isExportingRules = false;
        private readonly SynchronizationContext uiContext;
        private readonly object watcherLock = new object();
        private readonly ConcurrentDictionary<string, FileSystemWatcher> documentWatchers = new ConcurrentDictionary<string, FileSystemWatcher>();
        private bool isShuttingDown = false;
        private readonly ConcurrentDictionary<string, List<string>> ignorePatterns = new ConcurrentDictionary<string, List<string>>();
        private const string ignoreFileName = ".ilogicignore";
        private readonly ParameterManager parameterManager;

        public iLogicBridge(Inventor.Application app)
        {
            inventorApp = app;
            isOpen = true; // Since we're getting app from AddIn, it's already open
            uiContext = SynchronizationContext.Current ?? new SynchronizationContext();
            parameterManager = new ParameterManager();
        }

        public void Start()
        {
            try
            {
                // Get iLogic
                iLogicAuto = GetiLogicAddIn(inventorApp);
                if (iLogicAuto != null)
                {
                    iLogicAuto.CallingFromOutside = false;

                    // Initial active document setup - do this on the UI thread
                    OpenDocumentRules();

                    // Set up document open event instead of activation event
                    inventorApp.ApplicationEvents.OnOpenDocument += ApplicationEvents_OnOpenDocument;
                    inventorApp.ApplicationEvents.OnCloseDocument += ApplicationEvents_OnCloseDocument;
                    inventorApp.ApplicationEvents.OnSaveDocument += ApplicationEvents_OnSaveDocument;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in iLogic Bridge: {ex.Message}");
            }
        }



        public void Stop()
        {
            try
            {
                isShuttingDown = true;

                // Clean up event handlers
                if (inventorApp != null)
                {
                    inventorApp.ApplicationEvents.OnOpenDocument -= ApplicationEvents_OnOpenDocument;
                    inventorApp.ApplicationEvents.OnCloseDocument -= ApplicationEvents_OnCloseDocument;
                    inventorApp.ApplicationEvents.OnSaveDocument -= ApplicationEvents_OnSaveDocument;
                }

                // Clean up all watchers
                foreach (var watcherEntry in documentWatchers)
                {
                    try
                    {
                        var docWatcher = watcherEntry.Value;
                        if (docWatcher != null)
                        {
                            docWatcher.EnableRaisingEvents = false;
                            docWatcher.Changed -= OnChanged;
                            docWatcher.Created -= OnCreated;
                            docWatcher.Deleted -= OnDeleted;
                            docWatcher.Renamed -= OnRenamed;
                            docWatcher.Dispose();
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error disposing watcher: {ex.Message}");
                    }
                }
                documentWatchers.Clear();

                // Release references
                iLogicAuto = null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error stopping iLogic Bridge: {ex.Message}");
            }
        }

        private void ApplicationEvents_OnSaveDocument(_Document DocumentObject, EventTimingEnum BeforeOrAfter, NameValueMap Context, out HandlingCodeEnum HandlingCode)
        {
            HandlingCode = HandlingCodeEnum.kEventNotHandled;

            if (BeforeOrAfter == EventTimingEnum.kAfter)
            {
                // Log information about document saving
                System.Diagnostics.Debug.WriteLine($"Saving document: {System.IO.Path.GetFileName(DocumentObject.FullDocumentName)}");

                // Use the UI synchronization context to handle document operations
                uiContext.Post(_ =>
                {
                    try
                    {
                        ExportDocumentParameters(DocumentObject);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error exporting parameters on save: {ex.Message}");
                    }
                }, null);
            }
        }

        private void ApplicationEvents_OnOpenDocument(_Document DocumentObject, string FullDocumentName, EventTimingEnum BeforeOrAfter, NameValueMap Context, out HandlingCodeEnum HandlingCode)
        {
            HandlingCode = HandlingCodeEnum.kEventNotHandled;

            if (BeforeOrAfter == EventTimingEnum.kAfter)
            {
                // Log information about document opening
                System.Diagnostics.Debug.WriteLine($"Opening document: {System.IO.Path.GetFileName(FullDocumentName)}");

                // Use the UI synchronization context to handle document operations
                uiContext.Post(_ =>
                {
                    OpenDocumentRules(DocumentObject);
                    ImportDocumentParameters(DocumentObject);
                }, null);
            }
        }

        private void ApplicationEvents_OnCloseDocument(_Document DocumentObject, string FullDocumentName, EventTimingEnum BeforeOrAfter, NameValueMap Context, out HandlingCodeEnum HandlingCode)
        {
            HandlingCode = HandlingCodeEnum.kEventNotHandled;

            try
            {
                // We need to capture all the necessary information BEFORE the document is closed
                if (BeforeOrAfter == EventTimingEnum.kBefore)
                {
                    // Capture the information we need before the document is disposed
                    string docId = DocumentObject.InternalName.ToString();

                    // Use the UI synchronization context to clean up after the document
                    uiContext.Post(_ =>
                    {
                        try
                        {
                            CleanupDocumentResources(docId);
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error in document cleanup: {ex.Message}");
                        }
                    }, null);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in OnCloseDocument: {ex.Message}");
            }
        }

        private void CleanupDocumentResources(string docId)
        {
            if (isShuttingDown) return;

            try
            {
                // Remove watcher for this document
                if (documentWatchers.TryRemove(docId, out FileSystemWatcher docWatcher))
                {
                    try
                    {
                        docWatcher.EnableRaisingEvents = false;
                        docWatcher.Changed -= OnChanged;
                        docWatcher.Created -= OnCreated;
                        docWatcher.Deleted -= OnDeleted;
                        docWatcher.Renamed -= OnRenamed;
                        docWatcher.Dispose();
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error disposing document watcher: {ex.Message}");
                    }
                }

                // Log that the watcher was removed but we're keeping files for Git tracking
                System.Diagnostics.Debug.WriteLine($"Document watcher removed. Files preserved for Git tracking.");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in CleanupDocumentResources: {ex.Message}");
            }
        }

        /// <summary>
        /// Exports parameters from the document to the iLogic folder
        /// </summary>
        /// <param name="documentObject">The document to export parameters from</param>
        private void ExportDocumentParameters(_Document documentObject)
        {
            if (isShuttingDown) return;

            try
            {
                if (documentObject == null)
                {
                    System.Diagnostics.Debug.WriteLine("Cannot export parameters: Document is null");
                    return;
                }

                // Get the document folder
                string docFolder = System.IO.Path.GetDirectoryName(documentObject.FullDocumentName);

                // Look for .ilogicignore file to determine the iLogic folder
                string ignoreFilePath = FindIgnoreFile(docFolder);
                if (string.IsNullOrEmpty(ignoreFilePath))
                {
                    System.Diagnostics.Debug.WriteLine("No .ilogicignore file found. Skipping parameter export.");
                    return;
                }

                // Parse ignore file to check if transfer is enabled
                List<string> patterns = ParseIgnoreFile(ignoreFilePath, out bool shouldTransfer);
                if (!shouldTransfer)
                {
                    System.Diagnostics.Debug.WriteLine("Parameter transfer disabled in .ilogicignore. Skipping parameter export.");
                    return;
                }

                // Get the parent folder of the ignore file to determine the iLogic folder location
                string ignoreFileFolder = System.IO.Path.GetDirectoryName(ignoreFilePath);
                string ilogicLocalFolder = System.IO.Path.Combine(ignoreFileFolder, localILogicFolderName);

                // Create a document-specific subfolder for parameters
                string docName = System.IO.Path.GetFileNameWithoutExtension(documentObject.FullDocumentName);
                string docExtension = System.IO.Path.GetExtension(documentObject.FullDocumentName).TrimStart('.');
                string docFolderName = $"{docName}_{docExtension}";
                string docSpecificFolder = System.IO.Path.Combine(ilogicLocalFolder, docFolderName);

                // Create the subfolder if it doesn't exist
                if (!System.IO.Directory.Exists(docSpecificFolder))
                {
                    System.IO.Directory.CreateDirectory(docSpecificFolder);
                    System.Diagnostics.Debug.WriteLine($"Created document subfolder: {docSpecificFolder}");
                }

                // Export parameters to the document-specific folder
                parameterManager.ExportParameters(documentObject, docSpecificFolder);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error exporting document parameters: {ex.Message}");
            }
        }

        private void OpenDocumentRules(Inventor._Document documentObject = null)
        {
            if (isShuttingDown) return;

            // Create Inventor progress bar
            Inventor.ProgressBar progressBar = null;
            int totalSteps = 4; // Increased to include ignore file check

            try
            {
                // Initialize progress bar
                if (inventorApp != null)
                {
                    progressBar = inventorApp.CreateProgressBar(false, totalSteps, "iLogic Rules");
                    progressBar.Message = "Processing rules...";
                    // First step
                    progressBar.UpdateProgress();
                }

                // Update the active document and export rules
                try
                {
                    if (inventorApp != null && iLogicAuto != null)
                    {
                        Inventor.Document doc = documentObject ?? inventorApp.ActiveDocument;
                        if (doc != null)
                        {
                            string docName = System.IO.Path.GetFileName(doc.FullDocumentName);
                            System.Diagnostics.Debug.WriteLine($"Setting up rules for {docName}...");

                            if (progressBar != null)
                            {
                                progressBar.Message = $"Checking for ignore file";
                                progressBar.UpdateProgress();
                            }

                            // Get the document folder and check for .ilogicignore file
                            string docFolder = System.IO.Path.GetDirectoryName(doc.FullDocumentName);
                            bool shouldTransfer = false; // Default to not transferring
                            List<string> patterns = new List<string>();
                            string ilogicLocalFolder = null;
                            string ignoreFilePath = null;

                            // Look for .ilogicignore in the document folder or any parent folder
                            ignoreFilePath = FindIgnoreFile(docFolder);

                            if (!string.IsNullOrEmpty(ignoreFilePath))
                            {
                                string ignoreFileFolder = System.IO.Path.GetDirectoryName(ignoreFilePath);

                                // Parse the ignore file
                                patterns = ParseIgnoreFile(ignoreFilePath, out shouldTransfer);
                                ignorePatterns[doc.InternalName.ToString()] = patterns;

                                // Create or ensure the local ilogic folder exists in the same folder as the .ilogicignore file
                                ilogicLocalFolder = System.IO.Path.Combine(ignoreFileFolder, localILogicFolderName);
                                System.IO.Directory.CreateDirectory(ilogicLocalFolder);

                                System.Diagnostics.Debug.WriteLine($"Found .ilogicignore file at: {ignoreFilePath}");
                                System.Diagnostics.Debug.WriteLine($"Transfer enabled: {shouldTransfer}");
                                System.Diagnostics.Debug.WriteLine($"Using local ilogic folder: {ilogicLocalFolder}");
                                if (patterns.Count > 0)
                                {
                                    System.Diagnostics.Debug.WriteLine($"Ignoring patterns: {string.Join(", ", patterns)}");
                                }
                            }
                            else
                            {
                                System.Diagnostics.Debug.WriteLine($"No .ilogicignore file found. Not tracking iLogic files.");
                                return; // Exit early if no ignore file found
                            }

                            // If transfer is disabled in the ignore file, skip file watcher and export
                            if (!shouldTransfer)
                            {
                                System.Diagnostics.Debug.WriteLine($"Transfer disabled for {docName} in .ilogicignore file");
                                return;
                            }

                            if (progressBar != null)
                            {
                                progressBar.Message = $"Exporting rules for {docName}";
                                progressBar.UpdateProgress();
                            }

                            System.Diagnostics.Debug.WriteLine($"Exporting rules for {docName}...");

                            // Export rules to the appropriate folder (this will also set up document-specific watcher)
                            ExportRulesToFolder(doc, ilogicLocalFolder);

                            // Final progress update
                            if (progressBar != null)
                            {
                                progressBar.Message = "Completed";
                                // Final step
                                progressBar.UpdateProgress();
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error updating active document: {ex.Message}");
                    if (progressBar != null)
                    {
                        progressBar.Message = $"Error: {ex.Message}";
                    }
                }
            }
            finally
            {
                // Clean up progress bar
                if (progressBar != null)
                {
                    progressBar.Close();
                    progressBar = null;
                }

                // Log completion
                System.Diagnostics.Debug.WriteLine("Completed processing rules");
            }
        }

        /// <summary>
        /// Searches for a .ilogicignore file in the current folder and parent folders
        /// </summary>
        /// <param name="startFolder">Folder to start searching from</param>
        /// <returns>Full path to the .ilogicignore file if found, otherwise null</returns>
        private string FindIgnoreFile(string startFolder)
        {
            if (string.IsNullOrEmpty(startFolder))
                return null;

            string currentFolder = startFolder;
            int maxDepth = 10; // Limit how far up we search to prevent infinite loops

            for (int i = 0; i < maxDepth && !string.IsNullOrEmpty(currentFolder); i++)
            {
                string ignoreFilePath = System.IO.Path.Combine(currentFolder, ignoreFileName);
                if (System.IO.File.Exists(ignoreFilePath))
                {
                    return ignoreFilePath;
                }

                // Move up to parent folder
                try
                {
                    DirectoryInfo parentDir = System.IO.Directory.GetParent(currentFolder);
                    if (parentDir == null)
                        break;

                    currentFolder = parentDir.FullName;
                }
                catch
                {
                    break; // Exit if we can't get parent folder
                }
            }

            return null;
        }

        private List<string> ParseIgnoreFile(string filePath, out bool shouldTransfer)
        {
            List<string> patterns = new List<string>();
            shouldTransfer = true;

            try
            {
                string[] lines = System.IO.File.ReadAllLines(filePath);

                foreach (string line in lines)
                {
                    string trimmedLine = line.Trim();

                    // Skip empty lines and comments
                    if (string.IsNullOrWhiteSpace(trimmedLine) || trimmedLine.StartsWith("#"))
                        continue;

                    // Check for special directive
                    if (trimmedLine.Equals("@disable-transfer", StringComparison.OrdinalIgnoreCase))
                    {
                        shouldTransfer = false;
                        continue;
                    }

                    // Add pattern to the list
                    patterns.Add(trimmedLine);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error parsing ignore file: {ex.Message}");
            }

            return patterns;
        }

        private bool ShouldIgnoreRule(string docId, string ruleName)
        {
            if (!ignorePatterns.TryGetValue(docId, out List<string> patterns) || patterns.Count == 0)
                return false;

            foreach (string pattern in patterns)
            {
                try
                {
                    // Use wildcard matching
                    if (IsWildcardMatch(ruleName, pattern))
                        return true;
                }
                catch
                {
                    // If regex fails, try direct comparison
                    if (ruleName.Equals(pattern, StringComparison.OrdinalIgnoreCase))
                        return true;
                }
            }

            return false;
        }

        private bool IsWildcardMatch(string input, string pattern)
        {
            // Convert wildcard pattern to regex
            string regexPattern = "^" + Regex.Escape(pattern)
                .Replace("\\*", ".*")
                .Replace("\\?", ".") + "$";

            return Regex.IsMatch(input, regexPattern, RegexOptions.IgnoreCase);
        }

        private void ExportRulesToFolder(Inventor.Document doc, string targetFolder)
        {
            // Prevent recursive calls or overlapping exports
            if (isExportingRules || isShuttingDown)
                return;

            try
            {
                isExportingRules = true;

                if (doc != null && iLogicAuto != null)
                {
                    string docId = doc.InternalName.ToString();
                    // Export current document's rules
                    dynamic rules = iLogicAuto.get_Rules(doc);
                    if (rules == null)
                    {
                        System.Diagnostics.Debug.WriteLine("No rules found in the document.");
                        return;
                    }

                    // Create a document-specific subfolder for the rules
                    string docName = System.IO.Path.GetFileNameWithoutExtension(doc.FullDocumentName);
                    string docExtension = System.IO.Path.GetExtension(doc.FullDocumentName).TrimStart('.');
                    string docFolderName = $"{docName}_{docExtension}";
                    string docSpecificFolder = System.IO.Path.Combine(targetFolder, docFolderName);

                    // Create the subfolder if it doesn't exist
                    if (!System.IO.Directory.Exists(docSpecificFolder))
                    {
                        System.IO.Directory.CreateDirectory(docSpecificFolder);
                        System.Diagnostics.Debug.WriteLine($"Created document subfolder: {docSpecificFolder}");
                    }

                    foreach (iLogicRule r in rules)
                    {
                        try
                        {
                            // Check if rule should be ignored
                            if (ShouldIgnoreRule(docId, r.Name))
                            {
                                System.Diagnostics.Debug.WriteLine($"Ignoring rule {r.Name} as specified in .ilogicignore");
                                continue;
                            }

                            System.IO.File.WriteAllText(System.IO.Path.Combine(docSpecificFolder, r.Name + ".vb"), r.Text);
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error exporting rule {r.Name}: {ex.Message}");
                        }
                    }

                    // Update the file watcher to monitor the document-specific subfolder
                    CreateFileWatcherForDocument(doc, docSpecificFolder);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error exporting rules: {ex.Message}");
            }
            finally
            {
                isExportingRules = false;
            }
        }

        private Object GetiLogicAddIn(Inventor.Application app)
        {
            string iLogicGUID = "{3BDD8D79-2179-4B11-8A5A-257B1C0263AC}";
            Inventor.ApplicationAddIn iLogicAddIn = app.ApplicationAddIns.ItemById[iLogicGUID];
            iLogicAddIn.Activate();
            return iLogicAddIn.Automation;
        }

        private void CreateFileWatcherForDocument(Inventor.Document doc, string targetFolder)
        {
            if (isShuttingDown) return;

            try
            {
                // Get document ID for watcher tracking
                string docId = doc.InternalName.ToString();

                // Don't create a new watcher if one exists for this document
                if (documentWatchers.ContainsKey(docId))
                {
                    // If the folder has changed, update the watcher path
                    FileSystemWatcher existingWatcher = documentWatchers[docId];
                    if (existingWatcher != null && existingWatcher.Path != targetFolder)
                    {
                        try
                        {
                            existingWatcher.EnableRaisingEvents = false;
                            existingWatcher.Path = targetFolder;
                            existingWatcher.EnableRaisingEvents = true;
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error updating watcher path: {ex.Message}");
                            // If we can't update, remove and create a new one
                            if (documentWatchers.TryRemove(docId, out FileSystemWatcher oldWatcher))
                            {
                                try
                                {
                                    oldWatcher.EnableRaisingEvents = false;
                                    oldWatcher.Dispose();
                                }
                                catch { }
                            }
                        }
                    }
                    return;
                }

                // Create a new FileSystemWatcher and set its properties
                FileSystemWatcher newWatcher = new FileSystemWatcher
                {
                    Path = targetFolder,
                    NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.CreationTime,
                    Filter = "*.vb",
                    EnableRaisingEvents = false, // Start disabled until we add handlers
                    IncludeSubdirectories = false // We're watching specific document subfolder only
                };

                // Add event handlers
                newWatcher.Changed += OnChanged;
                newWatcher.Created += OnCreated;
                newWatcher.Deleted += OnDeleted;
                newWatcher.Renamed += OnRenamed;

                // Store document reference with the watcher
                if (documentWatchers.TryAdd(docId, newWatcher))
                {
                    // Begin watching only after successfully adding to dictionary
                    newWatcher.EnableRaisingEvents = true;
                    System.Diagnostics.Debug.WriteLine($"Created file watcher for document {System.IO.Path.GetFileName(doc.FullDocumentName)} at path: {targetFolder}");
                }
                else
                {
                    // If we couldn't add it to the dictionary, dispose it
                    newWatcher.Dispose();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error creating file watcher: {ex.Message}");
            }
        }

        private void OnChanged(object source, FileSystemEventArgs e)
        {
            if (isShuttingDown) return;
            ProcessFileChange(e.FullPath, FileChangeType.Changed);
        }

        private void OnCreated(object source, FileSystemEventArgs e)
        {
            if (isShuttingDown) return;
            ProcessFileChange(e.FullPath, FileChangeType.Created);
        }

        private void OnDeleted(object source, FileSystemEventArgs e)
        {
            if (isShuttingDown) return;
            ProcessFileChange(e.FullPath, FileChangeType.Deleted);
        }

        private void OnRenamed(object source, RenamedEventArgs e)
        {
            if (isShuttingDown) return;

            // Skip temp files
            if (e.OldName + "~" == e.Name)
            {
                System.Diagnostics.Debug.WriteLine("**VIM SWAP FILE, NOT A RENAME**");
                return;
            }

            ProcessFileChange(e.FullPath, FileChangeType.Renamed, e.OldFullPath);
        }

        private enum FileChangeType
        {
            Changed,
            Created,
            Deleted,
            Renamed
        }

        private void ProcessFileChange(string fullPath, FileChangeType changeType, string oldFullPath = null)
        {
            if (isShuttingDown) return;

            // Prevent processing the same file multiple times concurrently
            string key = fullPath + changeType.ToString();
            if (!processingFiles.TryAdd(key, true))
                return;

            try
            {
                // Use the UI sync context to ensure we're on the right thread for Inventor API calls
                uiContext.Post(_ =>
                {
                    try
                    {
                        if (isShuttingDown || inventorApp == null || iLogicAuto == null)
                        {
                            bool temp;
                            processingFiles.TryRemove(key, out temp);
                            return;
                        }

                        Inventor.Document doc = null;

                        // Determine which document the rule belongs to based on the subfolder
                        string ruleFolder = System.IO.Path.GetDirectoryName(fullPath);
                        string docFolderName = System.IO.Path.GetFileName(ruleFolder);

                        try
                        {
                            // First try by matching the folder name to document name with extension format
                            // Format should be "documentname_extension" (e.g., "Part1_ipt")
                            foreach (Inventor.Document openDoc in inventorApp.Documents)
                            {
                                string openDocName = System.IO.Path.GetFileNameWithoutExtension(openDoc.FullDocumentName);
                                string openDocExt = System.IO.Path.GetExtension(openDoc.FullDocumentName).TrimStart('.');
                                string expectedFolderName = $"{openDocName}_{openDocExt}";

                                if (expectedFolderName.Equals(docFolderName, StringComparison.OrdinalIgnoreCase))
                                {
                                    doc = openDoc;
                                    break;
                                }
                            }

                            // If no matching document, use active document as fallback
                            if (doc == null)
                            {
                                doc = inventorApp.ActiveDocument;
                                System.Diagnostics.Debug.WriteLine($"No document matching subfolder '{docFolderName}' found, using active document as fallback.");
                            }

                            if (doc == null)
                            {
                                bool temp;
                                processingFiles.TryRemove(key, out temp);
                                System.Diagnostics.Debug.WriteLine("No active document available for processing rule change.");
                                return;
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error finding document for rule: {ex.Message}");
                            bool temp;
                            processingFiles.TryRemove(key, out temp);
                            return;
                        }

                        string fileName = System.IO.Path.GetFileName(fullPath);
                        string ruleName = System.IO.Path.GetFileNameWithoutExtension(fileName);
                        string docId = doc.InternalName.ToString();

                        // Check if rule should be ignored
                        if (ShouldIgnoreRule(docId, ruleName))
                        {
                            System.Diagnostics.Debug.WriteLine($"Ignoring change to rule {ruleName} as specified in .ilogicignore");
                            bool temp;
                            processingFiles.TryRemove(key, out temp);
                            return;
                        }

                        try
                        {
                            switch (changeType)
                            {
                                case FileChangeType.Changed:
                                    try
                                    {
                                        dynamic rule = iLogicAuto.GetRule(doc, ruleName);
                                        if (rule != null)
                                        {
                                            System.Diagnostics.Debug.WriteLine($"Updating Rule {ruleName} in document {System.IO.Path.GetFileName(doc.FullDocumentName)}...");
                                            string newText = System.IO.File.ReadAllText(fullPath);
                                            rule.Text = newText;
                                        }
                                        else
                                        {
                                            System.Diagnostics.Debug.WriteLine($"**FAILED TO WRITE TO {ruleName} ON CHANGE: RULE DOESNT EXIST**");
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"Error updating rule {ruleName}: {ex.Message}");
                                    }
                                    break;

                                case FileChangeType.Created:
                                    try
                                    {
                                        dynamic existingRule = iLogicAuto.GetRule(doc, ruleName);
                                        if (existingRule == null)
                                        {
                                            System.Diagnostics.Debug.WriteLine($"Creating Rule {ruleName} in document {System.IO.Path.GetFileName(doc.FullDocumentName)}...");
                                            string newText = System.IO.File.ReadAllText(fullPath);
                                            dynamic newRule = iLogicAuto.AddRule(doc, ruleName, "");
                                            newRule.Text = newText;
                                        }
                                        else
                                        {
                                            System.Diagnostics.Debug.WriteLine($"**RULE ALREADY EXISTS: {ruleName}**");
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"Error creating rule {ruleName}: {ex.Message}");
                                    }
                                    break;

                                case FileChangeType.Deleted:
                                    try
                                    {
                                        dynamic ruleToDelete = iLogicAuto.GetRule(doc, ruleName);
                                        if (ruleToDelete != null)
                                        {
                                            System.Diagnostics.Debug.WriteLine($"WARNING: Removing Rule {ruleName} from document {System.IO.Path.GetFileName(doc.FullDocumentName)}...");
                                            iLogicAuto.DeleteRule(doc, ruleName);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"Error deleting rule {ruleName}: {ex.Message}");
                                    }
                                    break;

                                case FileChangeType.Renamed:
                                    try
                                    {
                                        string oldFileName = System.IO.Path.GetFileName(oldFullPath);
                                        string oldRuleName = System.IO.Path.GetFileNameWithoutExtension(oldFileName);

                                        dynamic oldRule = iLogicAuto.GetRule(doc, oldRuleName);
                                        if (oldRule != null)
                                        {
                                            System.Diagnostics.Debug.WriteLine($"Renaming Rule {oldRuleName} To {ruleName} in document {System.IO.Path.GetFileName(doc.FullDocumentName)}...");
                                            string ruleText = oldRule.Text;
                                            iLogicAuto.DeleteRule(doc, oldRuleName);
                                            dynamic renamedRule = iLogicAuto.AddRule(doc, ruleName, "");
                                            renamedRule.Text = ruleText;
                                        }
                                        else
                                        {
                                            System.Diagnostics.Debug.WriteLine($"**OLD RULE {oldRuleName} DOES NOT EXIST**");
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"Error renaming rule: {ex.Message}");
                                    }
                                    break;
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error processing {changeType} for {ruleName}: {ex.Message}");
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error processing file change: {ex.Message}");
                    }
                    finally
                    {
                        // Remove from processing dictionary when done
                        bool temp;
                        processingFiles.TryRemove(key, out temp);
                    }
                }, null);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error initiating file change processing: {ex.Message}");
                bool temp;
                processingFiles.TryRemove(key, out temp);
            }
        }

        /// <summary>
        /// Exports parameters from the document to the iLogic folder on document open
        /// </summary>
        /// <param name="documentObject">The document to export parameters from</param>
        private void ImportDocumentParameters(_Document documentObject)
        {
            if (isShuttingDown) return;

            try
            {
                if (documentObject == null)
                {
                    System.Diagnostics.Debug.WriteLine("Cannot export parameters: Document is null");
                    return;
                }

                // Get the document folder
                string docFolder = System.IO.Path.GetDirectoryName(documentObject.FullDocumentName);
                
                // Look for .ilogicignore file to determine the iLogic folder
                string ignoreFilePath = FindIgnoreFile(docFolder);
                if (string.IsNullOrEmpty(ignoreFilePath))
                {
                    System.Diagnostics.Debug.WriteLine("No .ilogicignore file found. Skipping parameter export.");
                    return;
                }

                // Parse ignore file to check if transfer is enabled
                List<string> patterns = ParseIgnoreFile(ignoreFilePath, out bool shouldTransfer);
                if (!shouldTransfer)
                {
                    System.Diagnostics.Debug.WriteLine("Parameter transfer disabled in .ilogicignore. Skipping parameter export.");
                    return;
                }

                // Get the parent folder of the ignore file to determine the iLogic folder location
                string ignoreFileFolder = System.IO.Path.GetDirectoryName(ignoreFilePath);
                string ilogicLocalFolder = System.IO.Path.Combine(ignoreFileFolder, localILogicFolderName);
                
                // Get the document-specific subfolder for parameters
                string docName = System.IO.Path.GetFileNameWithoutExtension(documentObject.FullDocumentName);
                string docExtension = System.IO.Path.GetExtension(documentObject.FullDocumentName).TrimStart('.');
                string docFolderName = $"{docName}_{docExtension}";
                string docSpecificFolder = System.IO.Path.Combine(ilogicLocalFolder, docFolderName);

                // Create the subfolder if it doesn't exist
                if (!System.IO.Directory.Exists(docSpecificFolder))
                {
                    System.IO.Directory.CreateDirectory(docSpecificFolder);
                    System.Diagnostics.Debug.WriteLine($"Created document subfolder: {docSpecificFolder}");
                }

                // Export parameters to the document-specific folder
                System.Diagnostics.Debug.WriteLine("Exporting parameters on document open...");
                parameterManager.ExportParameters(documentObject, docSpecificFolder);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error exporting document parameters: {ex.Message}");
            }
        }
    }
}