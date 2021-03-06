﻿using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Outlook;
using System.Diagnostics;
using System.Reflection;
using olinteroplib.ExtensionMethods;
using System.Text;
using static olinteroplib.Tracer;

namespace olinteroplib
{
    struct IMAPIPropertyTags {
        private static readonly string hidden = @"http://schemas.microsoft.com/mapi/proptag/0x10F5000B";
        private static readonly string subfolders = @"http://schemas.microsoft.com/mapi/proptag/0x360A000B";
        public static string PR_ATTR_HIDDEN { get { return hidden; } }
        public static string PR_SUBFOLDERS { get { return subfolders; } }
    }

    public static class Tracer {
        internal static TraceSource TraceOutput = new TraceSource("olinteroplib");
        public static Switch TracerLevel = TraceOutput.Switch;
        public static TraceListenerCollection TracerListeners = TraceOutput.Listeners;
    }

    public static class Methods {
        public static List<Folder> EnumerateFolders(Folder startFolder, bool includeHiddenFolders, uint maxDepth = 1) {
            if (maxDepth == 0)
                return new List<Folder>(0);

            List<Folder> list = new List<Folder>();
            if (!startFolder.HasSubfolders()) {
                TraceOutput.TraceEvent(TraceEventType.Verbose,0,"No subfolders found in folder specified");
                return new List<Folder>(0);
            }
            TraceOutput.TraceEvent(TraceEventType.Verbose, 0, $"Found {startFolder.Folders.Count} folders in {startFolder.Name}");

            foreach (Folder folder in startFolder.Folders) {
                if (folder.IsHidden() && !includeHiddenFolders) 
                    continue;

                list.Add(folder);
                TraceOutput.TraceEvent(TraceEventType.Information, 0, $"\t{folder.Name}");

                if (folder.Folders.Count > 0) {
                    list.AddRange(EnumerateFolders(folder, includeHiddenFolders, maxDepth--));
                }
            }
            return list;
        }
        public static void DisableVisiblePrintUserProp(UserProperty prop) {
            long printablePropertyFlag = 0x4; // PDO_PRINT_SAVEAS
            string printablePropertyCode = "[DispID=107]";
            Type customPropertyType = prop.GetType();

            // Get current flags.
            object rawFlags =
                customPropertyType.InvokeMember(printablePropertyCode, BindingFlags.GetProperty, null, prop, null);
            long flags = long.Parse(rawFlags.ToString());

            // Remove printable flag.
            flags &= ~printablePropertyFlag;

            object[] newParameters = new object[] { flags };

            // Set current flags.
            customPropertyType.InvokeMember(printablePropertyCode, BindingFlags.SetProperty, null, prop, newParameters);
        }
    }

    public static class CategoryParser {
        public static List<string> ConvertToList(string categoryString) {
            List<string> categoryList = new List<string>();
            if (categoryString == null) {
                categoryList.Add(categoryString);
            }
            else {
                string[] c = categoryString.Split(',');
                foreach (string s in c) {
                    categoryList.Add(s);
                }
            }
            return categoryList;
        }
        public static string ConvertToString(List<string> categoryList) {
            string categories = String.Empty;
            foreach (string s in categoryList) {
                if (String.IsNullOrEmpty(categories))
                    categories = s;
                else
                    categories = String.Join(",", categories, s);
            }
            return categories;
        }
    }

    namespace ExtensionMethods {
        public static class ExtensionMethods {
            private static int HasIMAPIProperty(this Folder folder, string propertyTag) {
                try {
                    bool hasProp = (bool)folder.PropertyAccessor.GetProperty(propertyTag);
                    if (hasProp)
                        return 1;
                    else
                        return 0;
                }
                catch (System.Exception e) {
                    TraceOutput.TraceInformation(e.Message);
                    return -1;
                }
            }

            public static bool HasSubfolders(this Folder folder) {
                bool b;
                if (folder.HasIMAPIProperty(IMAPIPropertyTags.PR_SUBFOLDERS) == 1)
                    b = true;
                else
                    b = false;
                TraceOutput.TraceEvent(TraceEventType.Verbose, 0, $"{folder.Name} has subfolders: {b}");
                return b;
            }
            public static bool IsHidden(this Folder folder) {
                bool b;
                if (folder.HasIMAPIProperty(IMAPIPropertyTags.PR_ATTR_HIDDEN) == 1)
                    b = true;
                else
                    b = false;
               TraceOutput.TraceEvent(TraceEventType.Verbose, 0, $"{folder.Name} is hidden: {b}");
                return b;
            }
            public static Folder Parent(this Folder folder) {
                return (Folder)folder.Parent;
            }
            public static Folder GetFolder(this Folder folder, string folderName) {
                Folder f = null;
                try {
                    f = (Folder)folder.Folders[folderName];
                }
                catch (System.Runtime.InteropServices.COMException) {
                    foreach (Folder subfolder in folder.Folders) {
                        if (folderName.ToLower() == subfolder.Name.ToLower())
                            f = subfolder;
                    }
                    if (f == null) {
                        Trace.WriteLine($"No folder exists with that name in {folder.Name}", folderName);
                        return null;
                    }
                }
                return f;
            }
            public static Folder GetFolder(this Folders folders, string folderName) {
                Folder f = null;
                try {
                    f = (Folder)folders[folderName];
                }
                catch (System.Runtime.InteropServices.COMException) {
                    foreach (Folder subfolder in folders) {
                        if (folderName.ToLower() == subfolder.Name.ToLower())
                            f = subfolder;
                    }
                    if (f == null) {
                        Trace.WriteLine($"No folder \"{folderName}\" exists in {((Folder)folders.Parent).Name} subfolders");
                        return null;
                    }
                }
                return f;
            }

            public static Folder GetFolderByPath(this NameSpace session, string path) {
                StringBuilder msg = new StringBuilder();
                msg.Append($"Get folder from path input '{path}': ");

                if (!path.Contains("/")) {
                    try {
                        Folder target = ((Folder)session.DefaultStore.GetRootFolder()).GetFolder(path);
                        return target;
                    }
                    catch {
                        msg.Append($"Failed - Invalid folder/path");
                        TraceOutput.TraceEvent(TraceEventType.Verbose, 0, msg.ToString());
                        throw new System.IO.DirectoryNotFoundException();
                    }
                }
                char slashChar = '/';
                Folder root = session.DefaultStore.GetRootFolder() as Folder;

                if (path.StartsWith("/") | path.EndsWith("/"))
                    path = path.Trim(slashChar);

                string[] folders = path.Split(slashChar);

                Folder folder = root;
                try {
                    if (folder != null) {
                        for (int i = 0; i <= folders.GetUpperBound(0); i++) {
                            Folders subFolders = folder.Folders;
                            folder = subFolders.GetFolder(folders[i]);
                            if (folder == null) {
                                msg.Append($"Failed - Folder not found at path");
                                TraceOutput.TraceEvent(TraceEventType.Information, 0, msg.ToString());
                                throw new System.IO.DirectoryNotFoundException();
                            }
                        }
                    }
                }
                catch (System.Exception e) {
                    msg.Append($"Failed - {e.Message}");
                    TraceOutput.TraceEvent(TraceEventType.Information, 0, msg.ToString());
                    throw new System.IO.DirectoryNotFoundException();
                }
                msg.Append($"Success");
                TraceOutput.TraceEvent(TraceEventType.Information, 0, msg.ToString());
                return folder;
            }
            public static Folder CreateFolderAtPath(this NameSpace session, string path, string name = null) {
                string parentPath = path;
                if (name == null) {
                    parentPath = path.Substring(0, path.LastIndexOf("/"));
                    name = path.Substring(path.LastIndexOf("/"));
                }
                try {
                    Folder newFolder = (Folder)GetFolderByPath(session, parentPath).Folders.Add("name");
                    return newFolder;
                }
                catch (System.IO.DirectoryNotFoundException) {
                    throw new System.IO.DirectoryNotFoundException(
                        $"Directory not found, could not create folder at {path}");
                }
            }

            public static void RemoveFolderUserProperty(this Folder folder, string name) {
                if (folder.UserDefinedProperties.Count == 0)
                    return;

                int index = 1;
                foreach (UserDefinedProperty prop in folder.UserDefinedProperties) {
                    if (prop.Name == name) {
                        folder.UserDefinedProperties.Remove(index);
                    }
                    else
                        index++;
                }
            }
            public static void RemoveMailItemUserProperty(this MailItem item, string name) {
                if (item.UserProperties.Count == 0)
                    return;

                int index = 1;
                foreach (UserProperty prop in item.UserProperties) {
                    if (prop.Name == name) {
                        item.UserProperties.Remove(index);
                    }
                    else
                        index++;
                }
            }
            public static void RemoveUserProperty_FolderAndItems(this Folder folder, string name) {
                if (folder.UserDefinedProperties.Count == 0)
                    return;

                foreach (MailItem item in folder.Items) {
                    if (item.UserProperties.Find(name) != null)
                        item.RemoveMailItemUserProperty(name);
                }
                folder.RemoveFolderUserProperty(name);
            }

            public static void AddCategory(this MailItem item, Category category) {
                List<string> categories = CategoryParser.ConvertToList(item.Categories);
                if (!categories.Contains(category.Name))
                    categories.Add(category.Name);
                item.Categories = CategoryParser.ConvertToString(categories);
            }
            public static void AddCategory(this MailItem item, string category) {
                List<string> categories = CategoryParser.ConvertToList(item.Categories);
                if (!categories.Contains(category))
                    categories.Add(category);
                item.Categories = CategoryParser.ConvertToString(categories);
            }
            public static void RemoveCategory(this MailItem item, string category) {
                List<string> categories = CategoryParser.ConvertToList(item.Categories);
                if (categories.Contains(category))
                    categories.Remove(category);
                item.Categories = CategoryParser.ConvertToString(categories);
            }
            public static void RemoveCategory(this MailItem item, Category category) {
                List<string> categories = CategoryParser.ConvertToList(item.Categories);
                if (categories.Contains(category.Name))
                    categories.Remove(category.Name);
                item.Categories = CategoryParser.ConvertToString(categories);
            }
            public static void RemoveAllCategories(this MailItem item) {
                item.Categories = "";
                item.Save();
            }
            public static bool CategoryExists(this NameSpace session, string name) {
                try {
                    Category category =
                            session.Categories[name];
                    if (category != null) {
                        return true;
                    }
                    else {
                        return false;
                    }
                }
                catch { return false; }
            }


        }
    }
}
