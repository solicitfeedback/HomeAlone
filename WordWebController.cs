using System.Globalization;
using Microsoft.AspNetCore;
using System.Net.Http;
using System.Threading.Tasks;
using WopiHost.Discovery;
using WopiHost.Discovery.Enumerations;
using WopiHost.Url;
using WopiHost.Abstractions;
using WopiHost.FileSystemProvider;
using WopiHost.Web.Models;
using Microsoft.Extensions.Options;
using System;
using System.IO;

using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Bson;
using Newtonsoft.Json.Linq;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using Microsoft.Office.Interop.Word;
using Microsoft.Vbe.Interop;
using Microsoft.Office.Core;
using System.Web.Http;
using System.Net;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Routing.Constraints;
using Microsoft.AspNetCore.Identity;
using Microsoft.Extensions.DependencyInjection;
using DocumentFormat.OpenXml.Office;
using Microsoft.AspNetCore.Authentication.OAuth;
using System.Linq;
using WopiHost.Core.Controllers;
using WopiHost.Core.Models;
using System.Web.Http.Results;
using Microsoft.AspNetCore.Diagnostics;
using System.Xml.Schema;
using Spire.Doc;
using Document = Spire.Doc.Document;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using Aspose.Words;
using Aspose.Words.Settings;
using Aspose.Words.Saving;
using System.Threading;
using Microsoft.AspNetCore.StaticFiles;
using System.Text.RegularExpressions;
using DocumentProperty = Spire.Doc.DocumentProperty;
using DocumentFormat.OpenXml.Wordprocessing;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using OpenXmlPowerTools;
//using WebSupergoo.WordGlue3;
//using GemBox.Document;
//using System.Web.Routing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
//using NUnit.Framework;
using OpenXMLTemplates;
using OpenXMLTemplates.ControlReplacers;
using OpenXMLTemplates.Documents;
using OpenXMLTemplates.Utils;
using OpenXMLTemplates.Variables;
using OpenXMLTemplates.Engine;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using Break = DocumentFormat.OpenXml.Wordprocessing.Break;
using TextInput = DocumentFormat.OpenXml.Wordprocessing.TextInput;
using System.Collections;
using System.Reflection;
using DocumentFormat.OpenXml.VariantTypes;
using System.Security.Cryptography;
using Aspose.Words.Loading;
using Microsoft.AspNetCore.Cors;
using StackExchange.Redis;
using Microsoft.Extensions.Caching.Distributed;
using Microsoft.Extensions.Caching.Memory;
using System.Xml.Linq;
using System.Xml;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Hosting;

namespace WopiHost.Web.Controllers
{




    public enum DocState : int
    {
        UNDEFINED = 0,
        BEFORE_EDIT = 1,
        EDIT_STARTED = 2,
        CHANGES_PENDING = 3,
        CHANGES_RESUMED = 4,
        CHANGES_DISCARD = 5,
        SAVE_PENDING = 6,
        EDIT_COMPLETED = 7

    }



    public enum ErrorArea : int
    {
        DEFAULT = 0,
        JSON = 0,
        VIEW = 1
    }

    public enum Severity : int
    {
        DEFAULT = 3,
        THREE = 3,
        TWO = 2,
        ONE = 1
    }

    public enum WordWebAction : int
    {
        DEFAULT = 0,
        RUNTIME = 1,
        VIEW = 2,
        EDIT = 3,
        SAVE = 4,
        UNSAVED = 5,
        DISCARD = 6
    }

    class DisableRemoteResourcesHandler : IResourceLoadingCallback
    {
        public ResourceLoadingAction ResourceLoading(ResourceLoadingArgs args)
        {
            return IsLocalResource(args.OriginalUri)
                ? ResourceLoadingAction.Default
                : ResourceLoadingAction.Skip;
        }

        private static bool IsLocalResource(string fileName)
        {
            DirectoryInfo dirInfo;
            try
            {
                var dirName = Path.GetDirectoryName(fileName);
                if (string.IsNullOrEmpty(dirName))
                    return false;
                dirInfo = new DirectoryInfo(dirName);
            }
            catch
            {
                return false;
            }

            foreach (DriveInfo d in DriveInfo.GetDrives())
            {
                if (string.Compare(dirInfo.Root.FullName, d.Name, StringComparison.OrdinalIgnoreCase) == 0)
                    return d.DriveType != DriveType.Network;
            }

            return false;
        }
    }


    class FileWatcher
    {
        System.IO.FileSystemWatcher watcher;
        public DateTime timestamp;
        public string path;
        public string key;
        public string user;
        public string attachmentID;
        public string version;
        public WordWebController wwcon;
        Regex expr = new Regex(@"(?<user>[^_]+)_(?<attachmentID>[^_]+)_*(?<version>[^\.]*)\.*");
        private readonly DateTime pointInTime = Convert.ToDateTime("1970-01-01");

        public FileWatcher(string path, DateTime dt, WordWebController wwcon)
        {

            watcher = new System.IO.FileSystemWatcher();
            watcher.Path = Path.GetDirectoryName(path);
            watcher.EnableRaisingEvents = true;
            watcher.Filter = Path.GetFileName(path);
            watcher.NotifyFilter = System.IO.NotifyFilters.LastWrite;
            watcher.Changed += new System.IO.FileSystemEventHandler(FileChanged);
            key = Path.GetFileNameWithoutExtension(path);
            var results = expr.Matches(key);            
            foreach (Match match in results)
            {
                user = match.Groups["user"].Value;
                attachmentID = match.Groups["attachmentID"].Value;
                version = match.Groups["version"].Value;
            }
            this.path = path;
            this.wwcon = wwcon;
            this.timestamp = System.IO.File.GetLastWriteTime(path); 
        }

        private void FileChanged(object sender, FileSystemEventArgs e)
        {
            //if (!IsFileReady(e.FullPath)) return; //first notification the file is arriving

            //The file has completed arrived, so lets process it
            //DoWorkOnFile(e.FullPath);
            if (IsFileReady(this.path))
            {
                this.wwcon.updateSessions(key, DocState.SAVE_PENDING, DateTime.Now, version, true);
                return;
            }

        }

        private bool IsFileReady(string path)
        {
            //One exception per file rather than several like in the polling pattern
            try
            {
                //If we can't open the file, it's still copying
                //using (var file = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.Read))

                var lastWritten = System.IO.File.GetLastWriteTime(this.path);
                if (lastWritten.CompareTo(this.timestamp) > 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (IOException)
            {
                return false;
            }
        }

        public void Dispose()
        {
            // avoiding resource leak
            watcher.Changed -= new System.IO.FileSystemEventHandler(FileChanged);
            this.watcher.Dispose();
        }
    }

    public class Attachment
    {
        [JsonPropertyName("mimeType")]
        public string mimeType { get; set; }

        [JsonPropertyName("contentsBlob")]
        public string fileContent { get; set; }

        [JsonPropertyName("attachVersionVrsNo")]
        public string versionNo { get; set; }

    }

    public class WopiLock
    {
        public string lockId { get; set; }        
    }


    public static class PasswordHelper
    {
        private static string encryptionKey = "/B?E(H+MbQeThWmZq4t7w9z$C&F)J@NcRfUjXn2r5u8x/A%D*G-KaPdSgVkYp3s6v9y$B&E(H+MbQeThWmZq4t7w!z%C*F-J@NcRfUjXn2r5u8x/A?D(G+KbPeSgVkYp3s6v9y$B&E)H@McQfTjWmZq4t7w!z%C*F-JaNdRgUkXp2r5u8x/A?D(G+KbPeShVmYq3t6v9y$B&E)H@McQfTjWnZr4u7x!z%C*F-JaNdRgUkXp2s5v8y/B?D(G+KbPe";
        public static string Encrypt(string clearText)
        {
            
            byte[] clearBytes = Encoding.Unicode.GetBytes(clearText);
            using (Aes encryptor = Aes.Create())
            {
                Rfc2898DeriveBytes pdb = new Rfc2898DeriveBytes(PasswordHelper.encryptionKey, new byte[] { 0x49, 0x76, 0x61, 0x6e, 0x20, 0x4d, 0x65, 0x64, 0x76, 0x65, 0x64, 0x65, 0x76 });
                encryptor.Key = pdb.GetBytes(32);
                encryptor.IV = pdb.GetBytes(16);
                using (MemoryStream ms = new MemoryStream())
                {
                    using (CryptoStream cs = new CryptoStream(ms, encryptor.CreateEncryptor(), CryptoStreamMode.Write))
                    {
                        cs.Write(clearBytes, 0, clearBytes.Length);
                        cs.Close();
                    }
                    clearText = Convert.ToBase64String(ms.ToArray());
                }
            }
            return clearText;
        }

        public static string Decrypt(string cipherText)
        {
            
            cipherText = cipherText.Replace(" ", "+");
            byte[] cipherBytes = Convert.FromBase64String(cipherText);
            using (Aes encryptor = Aes.Create())
            {
                Rfc2898DeriveBytes pdb = new Rfc2898DeriveBytes(PasswordHelper.encryptionKey, new byte[] { 0x49, 0x76, 0x61, 0x6e, 0x20, 0x4d, 0x65, 0x64, 0x76, 0x65, 0x64, 0x65, 0x76 });
                encryptor.Key = pdb.GetBytes(32);
                encryptor.IV = pdb.GetBytes(16);
                using (MemoryStream ms = new MemoryStream())
                {
                    using (CryptoStream cs = new CryptoStream(ms, encryptor.CreateDecryptor(), CryptoStreamMode.Write))
                    {
                        cs.Write(cipherBytes, 0, cipherBytes.Length);
                        cs.Close();
                    }
                    cipherText = Encoding.Unicode.GetString(ms.ToArray());
                }
            }
            return cipherText;
        }
    }

    public class DocLock
    {
        public IOptionsSnapshot<WopiOptions> WopiOptions { get; set; }
        public Dictionary<string, Dictionary<string, string>> lockedBy { get; set; }        

        public DocLock()
        {
            Dictionary<string, Dictionary<string, string>> dictionary = new Dictionary<string, Dictionary<string, string>>();
            
            lockedBy = dictionary;            
            
        }




        /* List implementation */
        /*
        public ulong getDocVersion(string user)
        {
            if (String.IsNullOrEmpty(user)) return (ulong) 0;
            foreach (var alock in lockedBy)
            {
                    if (alock.userId == user)
                        return alock.getVersion();
            }
            return (ulong)0;
        }

        
        public bool setDocVersion(string user, string version)
        {
            if (String.IsNullOrEmpty(user)) return false;
            foreach (var alock in lockedBy)
            {
                if (alock.userId == user)
                {
                    alock.version = version;
                    return true;
                }
                return false;
            }
            return false;
        }

        public string getDocVersionString(string user)
        {
            if (String.IsNullOrEmpty(user)) return "0";
            foreach (var alock in lockedBy)
            {
                    if (alock.userId == user)
                        return alock.getVersionString();
            }
            return "0";
        } */

        public ulong getDocVersion(string docId, string user)
        {
            if (String.IsNullOrEmpty(docId)) return (ulong)0;
            if (String.IsNullOrEmpty(user)) user = WopiOptions.Value.AnonymousUserId;
            user = user.ToUpper();
            //if (userID.ToUpper() == "UNAUTHENTICATED")
            //{
            //    userID = WopiOptions.Value.AnonymousUserId;
            //}
            if (lockedBy.ContainsKey(docId))
            {
                var uvr = lockedBy[docId];
                if (uvr.ContainsKey(user))
                {
                    return Convert.ToUInt64(uvr[user]);
                }
                return (ulong)0;
            }
            return (ulong)0;
        }

        public bool setDocVersion(string docId, string user, string version, WordWebException wwex)
        {
            if (String.IsNullOrEmpty(docId) || String.IsNullOrEmpty(version))
                return false;

            if (String.IsNullOrEmpty(user)) user = WopiOptions.Value.AnonymousUserId;
            user = user.ToUpper();
            //if (userID.ToUpper() == "UNAUTHENTICATED")
            //{
            //    userID = WopiOptions.Value.AnonymousUserId;
            //}

            if (getDocVersionString(docId, user) != "0")
            {
                if (lockedBy.ContainsKey(docId))
                {
                    Dictionary<string, string> uvr = lockedBy[docId];
                    if (uvr != null)
                    {
                        if (uvr.ContainsKey(user))
                        {
                            string oldver = uvr[user];

                            if (String.IsNullOrEmpty(oldver)) return false;

                            ulong oldverId = Convert.ToUInt64(oldver);
                            ulong newverId = Convert.ToUInt64(version);
                            if (newverId >= oldverId)
                            {
                                uvr[user] = version;
                            }
                            else
                            {
                                if (wwex == null) wwex = new WordWebException();
                                wwex.errMessage.code = "400";
                                wwex.errMessage.area = ErrorArea.JSON;
                                wwex.errMessage.message = "Version newer than ${version} has been already exists ${oldver}";
                                wwex.errMessage.level = "warning";
                                return false;
                            }
                        }
                        return true;
                    }
                    else
                    {
                        uvr = new Dictionary<string, string>();
                        uvr.Add(user, version);
                    }
                    return true;
                }
                else
                {
                    Dictionary<string, string> userver = new Dictionary<string, string>();
                    userver.Add(user, version);
                    lockedBy.Add(docId, userver);
                    return true;
                }
            }
            else
            {
                //getDocVersionString(docId, user) == "0"                wwex.errMessage.code = "404";
                if (lockedBy.ContainsKey(docId))
                {
                    Dictionary<string, string> userver = lockedBy[docId];
                    if (userver != null)
                    {
                        if (userver.ContainsKey(user))
                        {
                            userver[user] = version;
                        }
                        else
                        {
                            userver.Add(user, version);
                        }
                    }
                    else
                    {
                        userver = new Dictionary<string, string>();
                        userver.Add(user, version);
                    }
                }
                else
                {
                    Dictionary<string, string> userver = new Dictionary<string, string>();
                    userver.Add(user, version);
                    lockedBy.Add(docId, userver);
                }
                return true;
                /*if (wwex == null) wwex = new WordWebException();
                wwex.errMessage.area = ErrorArea.JSON;
                wwex.errMessage.message = "Unable to locate the document version";
                wwex.errMessage.level = "warning";
                return false;*/
            }
        }


        public bool rmDocVersion(string docId, string user, string version, WordWebException wwex)
        {
            if (String.IsNullOrEmpty(docId))
                return false;

            if (String.IsNullOrEmpty(user)) user = WopiOptions.Value.AnonymousUserId;
            user = user.ToUpper();
            //if (userID.ToUpper() == "UNAUTHENTICATED")
            //{
            //    userID = WopiOptions.Value.AnonymousUserId;
            //}

            if (getDocVersionString(docId, user) != "0")
            {
                if (lockedBy.ContainsKey(docId))
                {
                    Dictionary<string, string> uvr = lockedBy[docId];
                    if (uvr != null)
                    {
                        if (uvr.ContainsKey(user))
                        {
                            uvr.Remove(user);
                            if (uvr.Count() == 0)
                            {
                                lockedBy.Remove(docId);
                            }
                        }
                        return true;
                    }
                    return true;
                }
                else
                {
                    return true;
                }
                //return true;
            }
            else
            {
                //getDocVersionString(docId, user) == "0"                wwex.errMessage.code = "404";
                if (wwex == null) wwex = new WordWebException();
                wwex.errMessage.area = ErrorArea.JSON;
                wwex.errMessage.message = "Unable to locate the document version";
                wwex.errMessage.level = "warning";
                return false;
            }
        }









        public string getDocVersionString(string docId, string user)
        {
            if (String.IsNullOrEmpty(docId)) return "0";
            if (String.IsNullOrEmpty(user)) user = WopiOptions.Value.AnonymousUserId;
            user = user.ToUpper();
            //if (userID.ToUpper() == "UNAUTHENTICATED")
            //{
            //    userID = WopiOptions.Value.AnonymousUserId;
            //}
            if (lockedBy.ContainsKey(docId))
            {
                var uvr = lockedBy[docId];
                if (uvr.ContainsKey(user))
                {
                    return uvr[user];
                }
                return "0";
            }
            return "0";
        }



    }



    public class UserVersion : IEquatable<UserVersion>
    {
        public string userId { get; set; }

        public string version { get; set; }

        public UserVersion(string user, string version)
        {
            this.userId = user;
            this.version = version;

        }

        public override string ToString()
        {
            return userId + "_" + version;
        }
        public ulong getVersion()
        {
            if (ulong.TryParse(version, out ulong int64version))
                return int64version;
            else
                return (ulong)0;
        }

        public String getVersionString()
        {
            if (!String.IsNullOrEmpty(version))
                return version;
            else
                return "0";
        }

        public bool Equals(UserVersion other)
        {
            if (other == null) return false;
            if (other.userId == null) return false;
            if (other.version == null) return false;
            return (this.userId.Equals(other.userId)) && this.version.Equals(other.version);
        }

        public override bool Equals(object obj)
        {
            if (obj == null) return false;
            UserVersion objAsUserVer = obj as UserVersion;
            if (objAsUserVer == null) return false;
            else return Equals(objAsUserVer);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                int userHashCode = userId.GetHashCode();
                int verHashCode = version.GetHashCode();
                return userHashCode + verHashCode;
            }
        }

        public static bool operator ==(UserVersion obj1, UserVersion obj2)
        {
            if (ReferenceEquals(obj1, obj2))
            {
                return true;
            }
            if (ReferenceEquals(obj1, null))
            {
                return false;
            }
            if (ReferenceEquals(obj2, null))
            {
                return false;
            }

            return obj1.Equals(obj2);
        }

        public static bool operator !=(UserVersion obj1, UserVersion obj2)
        {
            return !(obj1 == obj2);
        }

    }


    public class DocSession
    {
        public DocState docState { get; set; }

        public DocState prevState { get; set; }

        [JsonPropertyName("attachVersionVrsNo")]
        public string version { get; set; }
        public DateTime creationTime { get; set; }
        public DateTime lastModifiedTime { get; set; }

        public bool cacheFlushed { get; set; }

        public DocSession()
        {
            cacheFlushed = false;
        }
    }


    public class ErrorMessage
    {
        public string code { get; set; }
        public string message { get; set; }
        public string level { get; set; }
        public ErrorArea area { get; set; }

    }
    public class APIResponse
    {
        public string status { get; set; }

        public DateTime timestamp { get; set; }

        public IList<ErrorMessage> errors { get; set; }
    }


    public class DocField
    {
        public DocField() { }

        public DocField(string id) { 
            fieldId = id;
            fieldBegin = null;
            fieldLabel = null;
            fieldSep = null;
            fieldText = null;
            fieldEnd = null;
            fieldParagraph = null;
            fieldProperties = null;
            fieldParent = null; 
            isFormField = false;
            fillerRuns0 = new List<Run>();
            fillerRuns1 = new List<Run>();
            fillerRuns2 = new List<Run>();
            fillerRuns3 = new List<Run>();
            fillerRuns4 = new List<Run>();
            fillerRuns = new ArrayList();
        }

        public DocField(string id, Run begin, Run label, Run sep, Run text, Run end, Paragraph para, RunProperties prop, Run parent = null, bool isForm = false, 
            ArrayList fr = null, List<Run> fr0 = null, List<Run> fr1 = null, List<Run> fr2 = null, List<Run> fr3 = null, List<Run> fr4 = null) { 
            fieldId = id;
            fieldBegin = begin;
            fieldLabel = label;
            fieldSep = sep;
            fieldText = text;
            fieldEnd = end;
            fieldParagraph = para;
            fieldProperties = prop;
            fieldParent = parent;
            isFormField = isForm;
            // 0 - begin - 1 - label - 2 - sep - 3 - Text - 4 - end
            if (fr0 == null) fillerRuns0 = new List<Run>(); else fillerRuns0 = fr0;
            if (fr1 == null) fillerRuns1 = new List<Run>(); else fillerRuns1 = fr1;
            if (fr2 == null) fillerRuns2 = new List<Run>(); else fillerRuns2 = fr2;
            if (fr3 == null) fillerRuns3 = new List<Run>(); else fillerRuns3 = fr3;
            if (fr4 == null) fillerRuns4 = new List<Run>(); else fillerRuns4 = fr4;
            if (fr == null) fillerRuns = new ArrayList(); else fillerRuns = fr;
        }

        public string fieldId { get; set; }

        public bool isFormField { get; set; }
        public Paragraph fieldParagraph { get; set; }

        public Run fieldParent { get; set; }

        public RunProperties fieldProperties { get; set; }

        public Run fieldBegin { get; set; }

        public Run fieldSep { get; set; }

        public Run fieldEnd { get; set; }

        public Run fieldLabel { get; set; }

        public Run fieldText { get; set; }

        public ArrayList fillerRuns { get; set; }

        public List<Run> fillerRuns0 { get; set; }
        public List<Run> fillerRuns1 { get; set; }
        public List<Run> fillerRuns2 { get; set; }
        public List<Run> fillerRuns3 { get; set; }
        public List<Run> fillerRuns4 { get; set; }

    }



    [Authorize]
    public class WordWebController : Controller
    {
        private HttpClient client = new HttpClient();
        private HttpClient readClient = new HttpClient();
        private HttpClient saveClient = new HttpClient();
        private HttpClient flushClient = new HttpClient();
        private HttpClient refreshClient = new HttpClient();
        //private UserManager<ApplicationUser> _userManager;
        private static Dictionary<string, DateTime> userSessions = new Dictionary<string, DateTime>();
        private static Dictionary<string, DocSession> docSessions = new Dictionary<string, DocSession>();
        private static Dictionary<string, string> docProtection = new Dictionary<string, string>();
        private static Dictionary<string, string> wopiFileIds = new Dictionary<string, string>();
        private static DocLock editSessions = new DocLock();
        private static Microsoft.Office.Interop.Word.Application word = null; // new Microsoft.Office.Interop.Word.Application();
        private static Microsoft.Office.Interop.Word.Application wordPointer = null;
        private static int sessions = 0;
        private readonly DateTime pointInTime = Convert.ToDateTime("1970-01-01");
        private string cmsServicePwd = null;
        private string hostUrl = null;
        private ArrayList ignoreDocumentProtection = new ArrayList();
        private static IMemoryCache _memoryCache;
        private static IDistributedCache _distributedCache;
        private const string TRUE = "true";
        private const string FALSE = "false";
        private const string ANONYMOUS = "Anonymous";




        //static ConnectionMultiplexer muxer = ConnectionMultiplexer.Connect("appd12owa:6379");
        //IDatabase _cache = muxer.GetDatabase();
        //string value = "abcde";
        //_cache.StringSet("test",value);

        //var foo = conn.StringGet("foo");
        //Console.Out.WriteLine(foo);

        /*public static Dictionary<string, Func<string, bool, object>> conversionEngine =
                new Dictionary<string, Func<string, bool, object>>
                { 
                    { "Aspose", convertDocToDocxApose},
                    { "Spire", convertDocToDocxSpire },
                    { "Interop", convertDocToDocx } 
                }; */


        protected Attachment theAttachment = new Attachment();

        private WopiUrlBuilder _urlGenerator;

        private IOptionsSnapshot<WopiOptions> WopiOptions { get; set; }

        private IOptionsSnapshot<CORSOptions> CORSOptions { get; set; }

        protected IWopiStorageProvider StorageProvider { get; set; }

        public WopiDiscoverer Discoverer => new WopiDiscoverer(new HttpDiscoveryFileProvider(WopiOptions.Value.ClientUrl));

        //TODO: remove test culture value and load it from configuration SECTION
        public WopiUrlBuilder UrlGenerator => _urlGenerator ?? (_urlGenerator = new WopiUrlBuilder(Discoverer, new WopiUrlSettings { UI_LLCC = new CultureInfo("en-US") }));

        //protected ConnectionMultiplexer Muxer { get => muxer; set => muxer = value; }

        protected bool addEditSessions(DateTime editTime, string docId, string userId = null, string versionId = null, WordWebException wwex = null)
        {
            string anonyUser = WopiOptions.Value.AnonymousUserId;

            if (String.IsNullOrEmpty(docId)) return false;
            if (!String.IsNullOrEmpty(versionId))
            {
                ulong version = (ulong)0;
                if (versionId != "0")
                {
                    if (!(ulong.TryParse(versionId, out version)))
                    {
                        if (wwex == null) wwex = new WordWebException();
                        wwex.errMessage.area = ErrorArea.JSON;
                        wwex.errMessage.message = "Unable to locate the document version";
                        wwex.errMessage.level = "warning";
                        return false;
                    }

                    if (getLoggedInUserName(userId).Equals(ANONYMOUS))
                        return editSessions.setDocVersion(docId, WopiOptions.Value.AnonymousUserId, versionId, wwex);
                    else
                        return editSessions.setDocVersion(docId, userId, versionId, wwex);

                }
                else
                { //version id == 0
                    if (wwex == null) wwex = new WordWebException();
                    wwex.errMessage.area = ErrorArea.JSON;
                    wwex.errMessage.message = "Unable to locate the document version";
                    wwex.errMessage.level = "warning";
                    return false;
                }

            }
            if (wwex == null) wwex = new WordWebException();
            wwex.errMessage.area = ErrorArea.JSON;
            wwex.errMessage.message = "Unable to locate the document version";
            wwex.errMessage.level = "error";
            return false;
        }


        protected String getLoggedInUser()
        {
            var userName = ANONYMOUS;
            WopiSecurityHandler securityHandler = new WopiSecurityHandler();
            System.Security.Claims.ClaimsPrincipal user = HttpContext.User;
            if (user.Identity.IsAuthenticated && (!String.IsNullOrEmpty(user.Identity.Name)))
            {
                userName = user.Identity.Name.Replace($"{WopiOptions.Value.WindowsDomain}\\", "");
                // check if user exists in database                
                if (!securityHandler.Exists(user.Identity.Name))
                {
                    securityHandler.AddPrincipal(user.Identity.Name);
                }

            }
            return userName;
        }

        protected String getLoggedInUserName(string userID)
        {
            var userName = ANONYMOUS;
            WopiSecurityHandler securityHandler = new WopiSecurityHandler();
            if (null != WopiOptions.Value.CMSSSOEnabled && WopiOptions.Value.CMSSSOEnabled == TRUE)
            {

                System.Security.Claims.ClaimsPrincipal user = HttpContext.User;
                if (user.Identity.IsAuthenticated && (!String.IsNullOrEmpty(user.Identity.Name)))
                {
                    // check if user exists in database                
                    if (!securityHandler.Exists(user.Identity.Name))
                    {
                        userName = user.Identity.Name.Replace($"{WopiOptions.Value.WindowsDomain}\\", "");
                        securityHandler.AddPrincipal(user.Identity.Name);
                        return userName;
                    }
                }
            }

            if (!String.IsNullOrEmpty(userID))
            {
                userID = userID.ToUpper();
                //if (userID.ToUpper() == "UNAUTHENTICATED")
                //{
                //    userID = WopiOptions.Value.AnonymousUserId;
                //}
                if (!securityHandler.Exists(userID))
                {
                    securityHandler.AddPrincipal(userID);
                }
                return userID;
            }
            return ANONYMOUS;
        }





        protected String getTempFile(string userID, string aid, string versionNo = null)
        {

            WopiSecurityHandler securityHandler = new WopiSecurityHandler();
            System.Security.Claims.ClaimsPrincipal user = HttpContext.User;
            if (user.Identity.IsAuthenticated && (!String.IsNullOrEmpty(user.Identity.Name)))
            {
                var userName = user.Identity.Name.Replace($"{WopiOptions.Value.WindowsDomain}\\", "");
                if (!securityHandler.Exists(userName))
                {
                    securityHandler.AddPrincipal(userName);
                }
                if (!String.IsNullOrEmpty(versionNo))
                {
                    return userName + "_" + aid + "_" + versionNo;
                }
                else return userName + "_" + aid;
            }
            else
            {
                var tempFileName = "";

                if (!String.IsNullOrEmpty(userID))
                {
                    var filePrefix = WopiOptions.Value.AnonymousUserId;
                    userID = userID.ToUpper();
                    //if (userID.ToUpper() == "UNAUTHENTICATED")
                    //{
                    //    userID = WopiOptions.Value.AnonymousUserId;
                    //}

                    if (!securityHandler.Exists(userID))
                    {
                        securityHandler.AddPrincipal(userID);
                    }
                    tempFileName = userID + "_" + aid;
                }
                else
                {
                    tempFileName = $"{WopiOptions.Value.AnonymousUserId}_" + aid;
                }
                if (!String.IsNullOrEmpty(versionNo))
                {
                    tempFileName = tempFileName + "_" + versionNo;
                }
                return tempFileName;
            }
        }

        protected String getTempFileWithExt(string userID, string aid, string extension, string versionNo = null)
        {
            if (String.IsNullOrEmpty(extension))
            {
                extension = WopiOptions.Value.WordExt;
            }
            WopiSecurityHandler securityHandler = new WopiSecurityHandler();
            System.Security.Claims.ClaimsPrincipal user = HttpContext.User;
            if (user.Identity.IsAuthenticated && (!String.IsNullOrEmpty(user.Identity.Name)))
            {
                var userName = user.Identity.Name.Replace($"{WopiOptions.Value.WindowsDomain}\\", "");
                if (!securityHandler.Exists(userName))
                {
                    securityHandler.AddPrincipal(userName);
                }
                if (!String.IsNullOrEmpty(versionNo))
                {
                    return userName + "_" + aid + "_" + versionNo + extension;
                }
                return userName + "_" + aid + extension;
            }
            else
            {
                var tempFileName = "";

                if (!String.IsNullOrEmpty(userID))
                {

                    var filePrefix = WopiOptions.Value.AnonymousUserId;
                    userID = userID.ToUpper();
                    //if (userID.ToUpper() == "UNAUTHENTICATED")
                    //{
                    //    userID = WopiOptions.Value.AnonymousUserId;
                    //}

                    if (!securityHandler.Exists(userID))
                    {
                        securityHandler.AddPrincipal(userID);
                    }
                    tempFileName = userID + "_" + aid;
                }
                else
                {
                    tempFileName = $"{WopiOptions.Value.AnonymousUserId}_" + aid;
                }
                if (!String.IsNullOrEmpty(versionNo))
                {
                    tempFileName = tempFileName + "_" + versionNo;
                }
                return tempFileName + extension;
            }
        }

        public DocState getDocSessionState(string key)
        {
            if (String.IsNullOrEmpty(key))
                return DocState.UNDEFINED;

            //ELSE one of the input parameters is not null
            if (docSessions.ContainsKey(key))
            {
                DocSession esess = docSessions[key];
                if (esess != null)
                {
                    return esess.docState;
                }
                else
                {
                    return DocState.UNDEFINED;
                }
            }
            else
            {
                return DocState.UNDEFINED;
            }

        }


        public DateTime getDocLastModified(string key)
        {
            if (String.IsNullOrEmpty(key))
                return pointInTime;

            //ELSE one of the input parameters is not null
            if (docSessions.ContainsKey(key))
            {
                DocSession esess = docSessions[key];
                if (esess != null)
                {
                    return esess.lastModifiedTime;
                }
                else
                {
                    return pointInTime;
                }
            }
            else
            {
                return pointInTime;
            }

        }



        public bool isCacheFlushed(string key)
        {
            if (String.IsNullOrEmpty(key))
                return false;

            //ELSE one of the input parameters is not null
            if (docSessions.ContainsKey(key))
            {
                DocSession esess = docSessions[key];
                if (esess != null)
                {
                    return esess.cacheFlushed;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }

        }


        public bool updateSessions(string key, DocState docState, DateTime mTime, string version = null, bool flush = false)
        {
            if (String.IsNullOrEmpty(key))
                return false;

            if (docState.Equals(DocState.UNDEFINED) && String.IsNullOrEmpty(version) && mTime == null)
                return false;

            //ELSE one of the input parameters is not null
            if (docSessions.ContainsKey(key))
            {
                DocSession esess = docSessions[key];

                if (esess != null)
                {
                    if (flush)
                    {
                        esess.cacheFlushed = true;
                    }



                    if (docState != null)
                    {
                        if (docState.Equals(DocState.CHANGES_DISCARD) && esess.docState.Equals(DocState.CHANGES_PENDING))
                        {
                            var attachedFile = WopiOptions.Value.RootPath + key + WopiOptions.Value.Word2010Ext;
                            var sourceFile = new FileInfo(attachedFile);
                            string newFileName = sourceFile.FullName.Replace(WopiOptions.Value.Word2010Ext, WopiOptions.Value.WordExt);
                            /* if (System.IO.File.Exists(attachedFile))
                             {
                                 System.IO.File.Delete(attachedFile);

                             }
                             if (System.IO.File.Exists(newFileName))
                             {
                                 System.IO.File.Delete(newFileName);

                             }*/

                        }


                        esess.prevState = esess.docState;
                        esess.docState = docState;


                    }
                    else
                    {
                        //docState==null, do nothing
                    }


                    if (version != null)
                        esess.version = version;

                    if (mTime != null)
                        esess.lastModifiedTime = mTime;
                }
                docSessions[key] = esess;
            }
            else
            {
                //Creating new session entry
                DocSession newSession = new DocSession();
                if (docState != null)
                {
                    newSession.prevState = DocState.UNDEFINED;
                    newSession.docState = docState;
                }
                else
                {
                    newSession.prevState = DocState.UNDEFINED;
                    newSession.docState = DocState.UNDEFINED;
                }

                if (version != null)
                    newSession.version = version;

                newSession.creationTime = DateTime.Now;

                if (mTime != null)
                    newSession.lastModifiedTime = mTime;
                else
                    newSession.lastModifiedTime = DateTime.Now;

                if (flush)
                {
                    newSession.cacheFlushed = true;
                }

                docSessions.Add(key, newSession);
            }
            return true;
        }

        protected double checkSession(string userID)
        {
            double timerInMinutes = 720;
            DateTime lastInvoked = DateTime.Now;
            TimeSpan timeout;
            var user = getLoggedInUserName(userID);
            if (!String.IsNullOrEmpty(user))
            {
                if (!(userSessions.ContainsKey(user)))
                {
                    userSessions.Add(user, DateTime.Now);
                }
                else
                {

                    if (userSessions.TryGetValue(user, out lastInvoked))
                    {
                        timeout = DateTime.Now - lastInvoked;
                        if (timeout.TotalMinutes >= 720)
                        {
                            //prompt user session expires
                            return 0;
                        }
                        else
                        {
                            //reset session
                            timerInMinutes = timeout.TotalMinutes;
                            userSessions[user] = DateTime.Now;
                            return timerInMinutes;
                        }
                    }

                }
            }
            return timerInMinutes;

        }
        public WordWebController(IOptionsSnapshot<WopiOptions> wopiOptions, IWopiStorageProvider storageProvider, IMemoryCache memoryCache, IDistributedCache distributedCache, IHostingEnvironment env)
        {
            WopiOptions = wopiOptions;
            editSessions.WopiOptions = wopiOptions;
            StorageProvider = storageProvider;
            _memoryCache = memoryCache;
            _distributedCache = distributedCache;


            Aspose.Words.License license = new Aspose.Words.License();
            // This line attempts to set a license from several locations relative to the executable and Aspose.Words.dll.
            // You can also use the additional overload to load a license from a stream, this is useful for instance when the 
            // license is stored as an embedded resource
            try
            {
                license.SetLicense($"{env.ContentRootPath}\\Startup.ini");
                Console.Out.WriteLine("License set successfully.");
            }
            catch (Exception e)
            {
                // We do not ship any license with this example, visit the Aspose site to obtain either a temporary or permanent license. 
                Console.Out.WriteLine("\nThere was an error setting the license: " + e.Message);
            }

            //CORSOptions = corsOptions;
            //readAttachment("669891871");
            //saveAttachment("669891871");
            //var theAttachment = readAttachment("669910587", theAttachment).Result;
            // theAttachment = readAttachment("669910587").Result;
            //var newVersion = saveAttachment("669910587", theAttachment.fileContent, theAttachment.versionNo).Result;

        }

        public async Task<ActionResult> Index()
        {
            try
            {
                return View(StorageProvider.GetWopiFiles(StorageProvider.RootContainerPointer.Identifier));
            }
            catch (DiscoveryException ex)
            {
                return View("Error", ex);
            }
            catch (HttpRequestException ex)
            {
                return View("Error", ex);
            }
        }

        protected bool refreshDocFromCache(string userID, string id, string ver = null, bool isSaveOperation = false, WopiLock wopiLock = null)
        {
            //IEnumerable<IWopiFile> wopiFiles = StorageProvider.GetWopiFiles(StorageProvider.RootContainerPointer.Identifier);            
            string fileLock = FALSE;
            var tempFile = getTempFile(userID, id, ver);
            IEnumerable<IWopiFile> wopiFiles = StorageProvider.GetWopiFiles(StorageProvider.RootContainerPointer.Identifier);
            string wopiFileId = new String("");
            IWopiFile file = null;

            //if (System.IO.File.GetLastAccessTime)

            wopiFileId = getWopiFileId(tempFile, out file);


            var extension = new String("");
            if (file != null)
            {

                WopiSecurityHandler securityHandler = new WopiSecurityHandler();
                //var token = securityHandler.GenerateAccessToken("CG53813", file.Identifier);
                string currentUser = null;
                if (null != WopiOptions.Value.CMSSSOEnabled && WopiOptions.Value.CMSSSOEnabled == TRUE)
                {
                    var aUser = getLoggedInUser();
                    currentUser = aUser;
                }
                else
                {
                    currentUser = getLoggedInUserName(userID);
                }
                var token = securityHandler.GenerateAccessToken(currentUser, file.Identifier);
                var accessToken = securityHandler.WriteToken(token);
                //var accessToken = token;


                extension = file.Extension.TrimStart('.');
                //var flushUrl = await UrlGenerator.GetFileUrlAsync(extension, getHostUrl()+"/wopi/files/{id}/contents");
                refreshClient.BaseAddress = new Uri(getHostUrl());
                refreshClient.DefaultRequestHeaders.Add("X-WOPI-Override", "UNLOCK");

                if (wopiLock != null && wopiLock.lockId != null && !String.IsNullOrEmpty(wopiLock.lockId))
                {
                    refreshClient.DefaultRequestHeaders.Add("X-WOPI-Lock", wopiLock.lockId);
                }

                string principal = getUserForCMSAPI(userID) ?? ANONYMOUS;
                var parameters = new Dictionary<string, string> { { "j_username", principal }, { "j_password", getCMSServicePwd() }, { "user_type", "" }, { "access_token", accessToken }, { "principal", principal } };
                //var parameters = new Dictionary<string, string> { { "j_username", getUserForCMSAPI(userID) }, { "j_password", getCMSServicePwd() }, { "user_type", "" }, { "principal", currentUser }, { "access_token", accessToken } };
                //var parameters = new Dictionary<string, string> { { "j_username", getUserForCMSAPI(userID) }, { "j_password", getCMSServicePwd() }, { "user_type", "" }, { "principal", "VW56075" }, { "access_token", accessToken } };
                var encodedContent = new FormUrlEncodedContent(parameters);

                try
                {
                    ;
                    // var resp = await UrlGenerator.GetFileUrlAsync(extension, getHostUrl()+"/wopi/files/{id}/contents", WopiActionEnum.View).ConfigureAwait(false);
                    //var resp = flushClient.PostAsync($"/wopi/files/{id}/contents", encodedContent).Result;
                    //var resp = flushClient.PostAsync($"/wopi/files/{wopiFileId}/contents", encodedContent).Result;
                    //var resp = refreshClient.PostAsync($"/wopi/files/{wopiFileId}/?access_token={accessToken}",encodedContent).Result;
                    var resp = refreshClient.PostAsync($"/wopi/files/{wopiFileId}/?access_token={accessToken}", encodedContent).Result;
                    //var ls = new FilesController().getLockStorage();
                    if (resp.IsSuccessStatusCode)
                    {
                        var filePath = WopiOptions.Value.RootPath + tempFile + WopiOptions.Value.WordExt;
                        var newFileName = filePath.Replace(WopiOptions.Value.WordExt, "_saved" + WopiOptions.Value.WordExt);
                        //var newFileName = filePath +  "_saved" + WopiOptions.Value.WordExt;

                        if (System.IO.File.Exists(newFileName))
                        {
                            System.IO.File.Delete(newFileName);
                        }

#pragma warning disable CA1307 // Specify StringComparison
                        var fileStream = System.IO.File.Create(newFileName);
#pragma warning restore CA1307 // Specify StringComparison

                        Stream apiResponse = resp.Content.ReadAsStreamAsync().Result;
                        //Stream apiResponse = resp.Content.ReadAsStreamAsync().Result;
                        apiResponse.Seek(0, SeekOrigin.Begin);
                        apiResponse.CopyTo(fileStream);
                        fileStream.Close();

                        //return fileLock; 

                        /*refreshClient.DefaultRequestHeaders.Add("X-WOPI-Override", "LOCK");
                        if (wopiLock != null && wopiLock.lockId != null && !String.IsNullOrEmpty(wopiLock.lockId))
                        {
                            refreshClient.DefaultRequestHeaders.Add("X-WOPI-Lock", wopiLock.lockId);
                        }
                        resp = refreshClient.PostAsync($"/wopi/files/{wopiFileId}/?access_token={accessToken}", encodedContent).Result;*/
                        return true;
                    }

                }
                catch (Exception ex)
                {
                    return false;
                }
                finally
                {

                }


            }
            return false;

        }

        protected string getWopiFileId(string inFile, out IWopiFile wopiFile)
        {
            string wopiFileId = null;
            string aFile = inFile?.Trim();
            if (aFile == null)
            {
                wopiFile = null;
                return null;
            }
            if (wopiFileIds.ContainsKey(aFile))
            {
                wopiFileId = wopiFileIds[aFile];
                wopiFile = StorageProvider.GetWopiFile(wopiFileId);
                return wopiFileId;
            }

            IEnumerable<IWopiFile> wopiFiles = StorageProvider.GetWopiFiles(StorageProvider.RootContainerPointer.Identifier);
            wopiFile = null;
            bool initialized = false;
            try
            {
                Dictionary<string, string> wopiDictionary = wopiFiles.ToDictionary(key => key.Name, value => value.Identifier);
                if (null != wopiDictionary && wopiDictionary.ContainsKey(aFile + WopiOptions.Value.WordExt))
                {
                    if (wopiFileIds.Count > 0) initialized = true;
                    wopiFileIds.Add(aFile, wopiDictionary[aFile + WopiOptions.Value.WordExt]);
                    wopiFileId = wopiDictionary[aFile + WopiOptions.Value.WordExt];
                }
            }
            catch (Exception exp)
            {
                //quick search not successful;
            }

            if (!initialized)
            {
                // looping through the first time
                using (var sequenceEnum = wopiFiles.GetEnumerator())
                {
                    while (sequenceEnum.MoveNext())
                    {
                        var extLength = WopiOptions.Value.WordExt.Length;
                        var wordExt = WopiOptions.Value.WordExt.Replace(".", "");

                        if (!(String.IsNullOrEmpty(sequenceEnum.Current.Name))
                            && !(String.IsNullOrEmpty(sequenceEnum.Current.Extension)))
                        {

                            var fileName = sequenceEnum.Current.Name;
                            var fileExt = sequenceEnum.Current.Extension;
                            var fileNameNoExt = sequenceEnum.Current.Name.Replace("." + fileExt, "");
                            var fileId = sequenceEnum.Current.Identifier;

                            if (wopiFileIds.ContainsKey(fileNameNoExt)) continue;

                            if (fileName.Length == aFile.Length + extLength)
                            {
                                if (fileName != null && fileExt != null && fileId != null
                                // && (fileName.Length >= aFile.Length + extLength)
                                && (fileExt == wordExt))
                                {
                                    if (fileName.Substring(0, aFile.Length) == aFile)
                                    {
                                        wopiFileId = fileId;
                                    }
                                    if (!(String.IsNullOrEmpty(fileNameNoExt) || String.IsNullOrEmpty(fileId)))
                                    {
                                        if (wopiFileIds.ContainsKey(fileNameNoExt))
                                            wopiFileIds[fileNameNoExt] = fileId;
                                        else
                                            wopiFileIds.Add(fileNameNoExt, fileId);
                                    }
                                }

                            }
                        }
                    }
                }
            }
            if (!(String.IsNullOrEmpty(wopiFileId)))
            {
                wopiFile = StorageProvider.GetWopiFile(wopiFileId);
            }
            return ((String.IsNullOrWhiteSpace(wopiFileId)) ? null : wopiFileId);
        }



        protected string flushCache(string userID, string id, string ver = null, bool isSaveOperation = false, WopiLock wopiLock = null)
        {
            string fileLock = FALSE;
            var tempFile = getTempFile(userID, id, ver);
            IEnumerable<IWopiFile> wopiFiles = StorageProvider.GetWopiFiles(StorageProvider.RootContainerPointer.Identifier);
            string wopiFileId = null;
            IWopiFile file = null;

            //if (System.IO.File.GetLastAccessTime)

            wopiFileId = getWopiFileId(tempFile, out file);
            if (String.IsNullOrWhiteSpace(wopiFileId) || file == null) throw new WordWebException("unable to find file id");

            var extension = new String("");
            if (file != null)
            {

                WopiSecurityHandler securityHandler = new WopiSecurityHandler();
                //var token = securityHandler.GenerateAccessToken("CG53813", file.Identifier);
                string currentUser = null;
                if (null != WopiOptions.Value.CMSSSOEnabled && WopiOptions.Value.CMSSSOEnabled == TRUE)
                {
                    var aUser = getLoggedInUser();
                    currentUser = aUser;
                }
                else
                {
                    currentUser = getLoggedInUserName(userID);
                }
                var token = securityHandler.GenerateAccessToken(currentUser, file.Identifier);
                var accessToken = securityHandler.WriteToken(token);


                extension = file.Extension.TrimStart('.');
                //var flushUrl = await UrlGenerator.GetFileUrlAsync(extension, getHostUrl()+"/wopi/files/{id}/contents");
                flushClient.BaseAddress = new Uri(getHostUrl());

                string principal = getUserForCMSAPI(userID) ?? ANONYMOUS;
                var parameters = new Dictionary<string, string> { { "j_username", principal }, { "j_password", getCMSServicePwd() }, { "user_type", "" }, { "access_token", accessToken }, { "principal", principal } };

                //var parameters = new Dictionary<string, string> { { "j_username", getUserForCMSAPI(userID) }, { "j_password", getCMSServicePwd() }, { "user_type", "" }, { "access_token", accessToken } };
                //var parameters = new Dictionary<string, string> { { "j_username", getUserForCMSAPI(userID) }, { "j_password", getCMSServicePwd() }, { "user_type", "" }, { "principal", currentUser }, { "access_token", accessToken } };
                //var parameters = new Dictionary<string, string> { { "j_username", getUserForCMSAPI(userID) }, { "j_password", getCMSServicePwd() }, { "user_type", "" }, { "principal", "VW56075" }, { "access_token", accessToken } };
                var encodedContent = new FormUrlEncodedContent(parameters);



                try
                {
                    ;
                    // var resp = await UrlGenerator.GetFileUrlAsync(extension, $"{WopiOptions.Value.HostUrl}/wopi/files/{id}/contents", WopiActionEnum.View).ConfigureAwait(false);
                    //var resp = flushClient.PostAsync($"/wopi/files/{id}/contents", encodedContent).Result;
                    //var resp = flushClient.PostAsync($"/wopi/files/{wopiFileId}/contents", encodedContent).Result;
                    //var resp = flushClient.PostAsync($"/wopi/files/{wopiFileId}?access_token={accessToken}", encodedContent).Result;
                    var resp = flushClient.PostAsync($"/wopi/files/{wopiFileId}?access_token={accessToken}", encodedContent).Result;
                    //var ls = new FilesController().getLockStorage();
                    if (resp.IsSuccessStatusCode)
                    {


                        string apiResponse = resp.Content.ReadAsStringAsync().Result;
                        JObject checkFileInfoResponse = JObject.Parse(apiResponse);
                        JToken lockInfo;
                        if (checkFileInfoResponse.TryGetValue("LockInfo", out lockInfo))
                        {
                            if (lockInfo.HasValues)
                            {
                                fileLock = TRUE;
                                if (wopiLock != null)
                                {
                                    wopiLock.lockId = lockInfo.First.ToString();
                                }
                            }
                        }

                        /*try
                        {
                            if (isSaveOperation)
                            {
                                if (!(refreshDocFromCache(userID, tempFile, wopiFileId, accessToken)))
                                {
                                    throw new WordWebException("An error occurred while the file is being saved");
                                }
                            }
                        }
                        catch(Exception x)
                        {
                            throw new WordWebException("An error occurred while the file is being saved");
                        }*/
                    }
                    return fileLock;
                }
                catch (Exception ex)
                {
                    return fileLock;
                }
                finally
                {
                    //return fileLock;
                }


            }
            return fileLock;

        }


        [EnableCors("AllowAll")]
        [Microsoft.AspNetCore.Mvc.Route("WordWeb/Save")]
        [AllowAnonymous]
        public ActionResult Save(string id, [FromQuery] string attachVersionID, [FromQuery] string userID)
        {


            try
            {
                WopiSecurityHandler securityHandler = new WopiSecurityHandler();
                if (!String.IsNullOrEmpty(attachVersionID))
                {
                    id = attachVersionID;
                }
                if (!String.IsNullOrEmpty(userID))
                {
                    securityHandler = new WopiSecurityHandler();
                    if (!securityHandler.Exists(userID))
                    {
                        securityHandler.AddPrincipal(userID);
                    }
                }
                else
                {
                    var anonymous = ANONYMOUS;
                    if (null != WopiOptions.Value.CMSSSOEnabled && WopiOptions.Value.CMSSSOEnabled == TRUE)
                    {
                        var aUser = getLoggedInUser();
                        if (!String.IsNullOrEmpty(aUser) && aUser != anonymous) userID = aUser;

                    }
                    else
                    {
                        userID = getLoggedInUserName(anonymous);
                    }
                }


                /*Task<string> obTask = Task.Run(() => (
                        var watcher = new FileWatcher(Path.GetFullPath(watchFile), this);
                        
                ));*/

                HttpContext.Response.Headers["content-type"] = "application/json";
                var verEdited = editSessions.getDocVersionString(id, userID);

                if (verEdited == "0")
                {
                    WordWebException ww = new WordWebException();
                    var tempFile = getTempFile(userID, id, "1");
                    if (!addEditSessions(DateTime.Now, id, userID, "1", ww))
                    {
                        return ErrorView(ww, WordWebAction.SAVE).Result;
                    }
                }


                if (!updateSessions(getTempFile(userID, id, verEdited), DocState.SAVE_PENDING, DateTime.Now))
                {
                    return handleErrorMessage("500", "Error in retrieving the document edit session", "error").Result;
                }
                else
                {
                    //var verEdited = editSessions.getDocVersionString(id, userID);
                    if (Int32.TryParse(verEdited, out int numValue))
                    {
                        numValue += 1;
                        verEdited = Convert.ToString(numValue);
                    }


                    return Ok(new
                    {
                        status = "0",
                        timpstamp = String.Format("{0:G}", DateTime.Now),
                        attachVersionVrsNo = verEdited
                        //attachVersionVrsNo = "0"
                    });

                }
                /*

                 var source = new String("");
                 var target = new String("");


                 if (!String.IsNullOrEmpty(attachVersionID))
                 {
                     id = attachVersionID;
                 }

                 //# POST / wopi / files / (file_id) / contents


                 var beforeFlush = DateTime.Now;
                 var afterFlush = beforeFlush.AddMinutes((double)1.0);
                 WopiLock wopiLock = new WopiLock();


                 if (verEdited != "0")
                 {

                     flushCache(userID, id, verEdited, true, wopiLock);
                     //refreshDocFromCache(userID, id, verEdited, true, wopiLock);

                     //refreshDocFromCache(userID, id, verEdited, true);
                     //flushCache(userID, id, verEdited);
                     //while (flushCache(userID, id, verEdited) == TRUE || DateTime.Now <= afterFlush)
                     //{
                     //    ;//Thread.Sleep(100);
                     //}
                 }*/
                /*
                else
                {
                    flushCache(userID, id, null, true, wopiLock);
                    //refreshDocFromCache(userID, id, null, true, wopiLock);



                    //while (flushCache(userID, id) == TRUE || DateTime.Now <= afterFlush)
                    //{
                    //    ;// Thread.Sleep(100);
                    //}
                }




        */
                /*
                        source = WopiOptions.Value.RootPath + getTempFile(userID, id, verEdited) + WopiOptions.Value.WordExt;
                        target = WopiOptions.Value.RootPath + getTempFile(userID, id, verEdited) + WopiOptions.Value.Word2010Ext;

                        if (System.IO.File.Exists(source))
                        {
                            if (System.IO.File.Exists(target))
                            {
                                System.IO.File.Delete(target);
                            }
                            convertDocxToDocAspose(source);
                        }

                        WordWebException ww = new WordWebException();
                        //Attachment theAttachment = getAttachment(id, ww).Result;
                        //Attachment attachedDoc = getAttachment(id, ww).Result;
                        Attachment attachedDoc = new Attachment
                        {
                            mimeType = "application/msword",
                            fileContent = readAndEncodeFile(target),
                            versionNo = verEdited
                        };

                        if (attachedDoc == null && ww.errMessage != null)
                            return (ActionResult)handleError(ww.errMessage);
                        //return await ErrorView(ww, WordWebAction.SAVE);

                        attachedDoc.versionNo = saveAttachment(id, attachedDoc.fileContent, verEdited, ww).Result;

                        if (!String.IsNullOrWhiteSpace(attachedDoc.versionNo))
                        {
                            //ViewData["message"] = "Error in calling CMS save API for attachementID" + id;
                            if (attachedDoc.versionNo == "0")
                            {
                                return (ActionResult)handleError(ww.errMessage);
                                //return await ErrorView(ww, WordWebAction.SAVE);
                                //return await handleErrorMessage(Response.StatusCode.ToString(), "Error in calling CMS save API for attachementID" + id, "error");
                            }
                            //if (editSessions.rmDocVersion(id, userID, attachedDoc.versionNo, ww))
                            //{


                            bool verUpdated = editSessions.setDocVersion(id, userID, attachedDoc.versionNo, ww);

                            if (!verUpdated)
                            {
                                return (ActionResult)handleError(ww.errMessage);
                            }

                            HttpContext.Response.Headers["content-type"] = "application/json";
                            return Ok(new
                            {
                                status = "0",
                                timpstamp = String.Format("{0:G}", DateTime.Now),
                                attachVersionVrsNo = attachedDoc.versionNo
                                //attachVersionVrsNo = "0"
                            });
                            //}
                            //else
                            //return (ActionResult)handleError(ww.errMessage);

                        }
                        //return await handleErrorMessage(Response.StatusCode.ToString(), "Error in calling CMS save API for attachementID" + id, "error");
                        //return await ErrorView(ww, WordWebAction.SAVE);
                        return (ActionResult)handleError(ww.errMessage);
                     */
            }
            catch (DiscoveryException ex)
            {
                //return View("Error in calling CMS save API for attachementID" + id, ex);
                return handleErrorMessage("500", ex.Message, "error").Result;
            }
            catch (WebException we)
            {
                //return View("Error in calling CMS save API for attachementID" + id, ex);                
                return handleErrorMessage(((HttpWebResponse)we.Response).StatusCode.ToString(), we.Message, "error").Result;
            }
            catch (Exception e)
            {
                //return View("Error in calling CMS save API for attachementID" + id, ex);                
                return handleErrorMessage("500", e.Message, "error").Result;
            }
            finally
            {
                //if (System.IO.File.Exists(target))
                //{
                //    System.IO.File.Delete(target);
                //}
            }


        }


        public ActionResult SaveOriginal(string id, [FromQuery] string attachVersionID, [FromQuery] string userID)
        {


            try
            {
                // updateSessions(getTempFile(userID, id, version), DocState.SAVE_PENDING, DateTime.Now);


                /*HttpContext.Response.Headers["content-type"] = "application/json";
                return Ok(new
                {
                    status = "0",
                    timpstamp = String.Format("{0:G}", DateTime.Now),
                    attachVersionVrsNo = verEdited
                    //attachVersionVrsNo = "0"
                });*/


                var source = new String("");
                var target = new String("");



                if (!String.IsNullOrEmpty(attachVersionID))
                {
                    id = attachVersionID;
                }

                //# POST / wopi / files / (file_id) / contents

                var verEdited = editSessions.getDocVersionString(id, userID);
                var beforeFlush = DateTime.Now;
                var afterFlush = beforeFlush.AddMinutes((double)1.0);
                WopiLock wopiLock = new WopiLock();


                if (verEdited != "0")
                {

                    flushCache(userID, id, verEdited, true, wopiLock);
                    //refreshDocFromCache(userID, id, verEdited, true, wopiLock);

                    //refreshDocFromCache(userID, id, verEdited, true);
                    //flushCache(userID, id, verEdited);
                    /*while (flushCache(userID, id, verEdited) == TRUE || DateTime.Now <= afterFlush)
                    {
                        ;//Thread.Sleep(100);
                    }*/
                }
                else
                {
                    flushCache(userID, id, null, true, wopiLock);
                    //refreshDocFromCache(userID, id, null, true, wopiLock);



                    /*while (flushCache(userID, id) == TRUE || DateTime.Now <= afterFlush)
                    {
                        ;// Thread.Sleep(100);
                    }*/
                }





                source = WopiOptions.Value.RootPath + getTempFile(userID, id, verEdited) + WopiOptions.Value.WordExt;
                target = WopiOptions.Value.RootPath + getTempFile(userID, id, verEdited) + WopiOptions.Value.Word2010Ext;

                if (System.IO.File.Exists(source))
                {
                    if (System.IO.File.Exists(target))
                    {
                        System.IO.File.Delete(target);
                    }
                    //convertDocxToDocAspose(source);
                    //convertDocxToDocSpire(source);
                    var downConversionEngine = "convertDocxToDoc";
                    var downConversionEngineFallBack = "convertDocxToDoc";

                    if (String.IsNullOrEmpty($"WopiOptions.Value.ConversionEngine"))
                    {
                        downConversionEngineFallBack += WopiOptions.Value.ConversionEngineFallBack; //default
                        downConversionEngine = downConversionEngineFallBack;
                    }
                    else
                    {
                        downConversionEngine += WopiOptions.Value.ConversionEngine;
                    }
                    MethodInfo conversionMethod = this.GetType().GetMethod(downConversionEngine);
                    MethodInfo conversionMethodFallBack = this.GetType().GetMethod(downConversionEngineFallBack);
                    object result = null;
                    try
                    {
                        result = conversionMethod.Invoke(this, new object[] { source });
                    }
                    catch (Exception exc)
                    {
                        result = conversionMethodFallBack.Invoke(this, new object[] { source });
                    }

                }

                WordWebException ww = new WordWebException();
                //Attachment theAttachment = getAttachment(id, ww).Result;
                //Attachment attachedDoc = getAttachment(id, ww).Result;
                Attachment attachedDoc = new Attachment
                {
                    mimeType = "application/msword",
                    fileContent = readAndEncodeFile(target),
                    versionNo = verEdited
                };

                if (attachedDoc == null && ww.errMessage != null)
                    return (ActionResult)handleError(ww.errMessage);
                //return await ErrorView(ww, WordWebAction.SAVE);

                attachedDoc.versionNo = saveAttachment(id, attachedDoc.fileContent, verEdited, ww, userID).Result;

                if (!String.IsNullOrWhiteSpace(attachedDoc.versionNo))
                {
                    //ViewData["message"] = "Error in calling CMS save API for attachementID" + id;
                    if (attachedDoc.versionNo == "0")
                    {
                        return (ActionResult)handleError(ww.errMessage);
                        //return await ErrorView(ww, WordWebAction.SAVE);
                        //return await handleErrorMessage(Response.StatusCode.ToString(), "Error in calling CMS save API for attachementID" + id, "error");
                    }
                    //if (editSessions.rmDocVersion(id, userID, attachedDoc.versionNo, ww))
                    //{


                    bool verUpdated = editSessions.setDocVersion(id, userID, attachedDoc.versionNo, ww);

                    if (!verUpdated)
                    {
                        return (ActionResult)handleError(ww.errMessage);
                    }

                    HttpContext.Response.Headers["content-type"] = "application/json";
                    return Ok(new
                    {
                        status = "0",
                        timpstamp = String.Format("{0:G}", DateTime.Now),
                        attachVersionVrsNo = attachedDoc.versionNo
                        //attachVersionVrsNo = "0"
                    });
                    //}
                    //else
                    //return (ActionResult)handleError(ww.errMessage);

                }
                //return await handleErrorMessage(Response.StatusCode.ToString(), "Error in calling CMS save API for attachementID" + id, "error");
                //return await ErrorView(ww, WordWebAction.SAVE);
                return (ActionResult)handleError(ww.errMessage);

            }
            catch (DiscoveryException ex)
            {
                //return View("Error in calling CMS save API for attachementID" + id, ex);
                return handleErrorMessage("500", ex.Message, "error").Result;
            }
            catch (WebException we)
            {
                //return View("Error in calling CMS save API for attachementID" + id, ex);                
                return handleErrorMessage(((HttpWebResponse)we.Response).StatusCode.ToString(), we.Message, "error").Result;
            }
            catch (Exception e)
            {
                //return View("Error in calling CMS save API for attachementID" + id, ex);                
                return handleErrorMessage("500", e.Message, "error").Result;
            }
            finally
            {
                /*if (System.IO.File.Exists(target))
                {
                    System.IO.File.Delete(target);
                }*/
            }
        }



        public string checkPendingSave(string id, string attachVersionID, string userID, WordWebException ww)
        {


            try
            {

                var source = new String("");
                var target = new String("");



                if (!String.IsNullOrEmpty(attachVersionID))
                {
                    id = attachVersionID;
                }

                //# POST / wopi / files / (file_id) / contents

                var verEdited = editSessions.getDocVersionString(id, userID);
                var beforeFlush = DateTime.Now;
                var afterFlush = beforeFlush.AddMinutes((double)1.0);
                WopiLock wopiLock = new WopiLock();


                if (verEdited != "0")
                {

                    flushCache(userID, id, verEdited, true, wopiLock);
                    //refreshDocFromCache(userID, id, verEdited, true, wopiLock);

                    //refreshDocFromCache(userID, id, verEdited, true);
                    //flushCache(userID, id, verEdited);
                    /*while (flushCache(userID, id, verEdited) == TRUE || DateTime.Now <= afterFlush)
                    {
                        ;//Thread.Sleep(100);
                    }*/
                }
                else
                {
                    flushCache(userID, id, null, true, wopiLock);
                    //refreshDocFromCache(userID, id, null, true, wopiLock);



                    /*while (flushCache(userID, id) == TRUE || DateTime.Now <= afterFlush)
                    {
                        ;// Thread.Sleep(100);
                    }*/
                }





                source = WopiOptions.Value.RootPath + getTempFile(userID, id, verEdited) + WopiOptions.Value.WordExt;
                target = WopiOptions.Value.RootPath + getTempFile(userID, id, verEdited) + WopiOptions.Value.Word2010Ext;

                if (System.IO.File.Exists(source))
                {
                    if (System.IO.File.Exists(target))
                    {
                        System.IO.File.Delete(target);
                    }
                    //convertDocxToDocAspose(source);
                    //convertDocxToDocSpire(source);
                    var downConversionEngine = "convertDocxToDoc";
                    var downConversionEngineFallback = "convertDocxToDoc";

                    if (String.IsNullOrEmpty(WopiOptions.Value.ConversionEngine))
                    {
                        if (String.IsNullOrEmpty(WopiOptions.Value.ConversionEngineFallBack))
                        {
                            downConversionEngineFallback += WopiOptions.Value.ConversionEngineFallBack; //default
                            downConversionEngine = downConversionEngineFallback;
                        }
                    }
                    else
                    {
                        downConversionEngine += WopiOptions.Value.ConversionEngine;
                    }
                    MethodInfo conversionMethod = this.GetType().GetMethod(downConversionEngine);
                    MethodInfo conversionMethodFallback = this.GetType().GetMethod(downConversionEngineFallback);
                    object result = null;
                    try
                    {
                        result = conversionMethod.Invoke(this, new object[] { source });
                    }
                    catch (Exception exc)
                    {
                        result = conversionMethodFallback.Invoke(this, new object[] { source });
                    }
                }

                //WordWebException ww = new WordWebException();
                //Attachment theAttachment = getAttachment(id, ww).Result;
                //Attachment attachedDoc = getAttachment(id, ww).Result;
                Attachment attachedDoc = new Attachment
                {
                    mimeType = "application/msword",
                    fileContent = readAndEncodeFile(target),
                    versionNo = verEdited
                };

                if (attachedDoc == null && ww.errMessage != null)
                    return "-1";
                //return await ErrorView(ww, WordWebAction.SAVE);

                attachedDoc.versionNo = saveAttachment(id, attachedDoc.fileContent, verEdited, ww, userID).Result;

                if (!String.IsNullOrWhiteSpace(attachedDoc.versionNo))
                {
                    //ViewData["message"] = "Error in calling CMS save API for attachementID" + id;
                    if (attachedDoc.versionNo == "0")
                    {
                        return "-1";
                        //return await ErrorView(ww, WordWebAction.SAVE);
                        //return await handleErrorMessage(Response.StatusCode.ToString(), "Error in calling CMS save API for attachementID" + id, "error");
                    }
                    //if (editSessions.rmDocVersion(id, userID, attachedDoc.versionNo, ww))
                    //{


                    bool verUpdated = editSessions.setDocVersion(id, userID, attachedDoc.versionNo, ww);

                    if (!verUpdated)
                    {
                        return "-1";
                    }

                    return attachedDoc.versionNo;
                    //HttpContext.Response.Headers["content-type"] = "application/json";
                    /*return Ok(new
                    {
                        status = "0",
                        timpstamp = String.Format("{0:G}", DateTime.Now),
                        attachVersionVrsNo = attachedDoc.versionNo
                        //attachVersionVrsNo = "0"
                    });*/
                    //}
                    //else
                    //return (ActionResult)handleError(ww.errMessage);

                }
                //return await handleErrorMessage(Response.StatusCode.ToString(), "Error in calling CMS save API for attachementID" + id, "error");
                //return await ErrorView(ww, WordWebAction.SAVE);
                return "0";

            }
            catch (DiscoveryException ex)
            {
                //return View("Error in calling CMS save API for attachementID" + id, ex);
                return "-1";
            }
            catch (WebException we)
            {
                //return View("Error in calling CMS save API for attachementID" + id, ex);                
                return "-1";
            }
            catch (Exception e)
            {
                //return View("Error in calling CMS save API for attachementID" + id, ex);                
                return "-1";
            }
            finally
            {
                /*if (System.IO.File.Exists(target))
                {
                    System.IO.File.Delete(target);
                }*/
            }
        }




        public static HashSet<T> ToHashSet<T>(IEnumerable<T> items)
        {
            return new HashSet<T>(items);
        }

        public static string toBase64(string s)
        {

            byte[] buffer = System.Text.Encoding.ASCII.GetBytes(s);
            return System.Convert.ToBase64String(buffer);
        }

        //[Microsoft.AspNetCore.HttpGet]
        //[Microsoft.AspNetCore.Mvc.Route("WordWeb/unSavedChangeExists")]
        //public async Task<HttpResponseMessage> unSavedChangeExists(string id, [FromQuery]string attachVersionID)
        [EnableCors("AllowAll")]
        [Microsoft.AspNetCore.Mvc.Route("WordWeb/unSavedChangeExists")]
        [AllowAnonymous]
        public async Task<ActionResult> unSavedChangeExists(string id, [FromQuery] string attachVersionID, [FromQuery] string userID)
        {
            WopiSecurityHandler securityHandler = new WopiSecurityHandler();
            if (!String.IsNullOrEmpty(attachVersionID))
            {
                id = attachVersionID;
            }
            if (!String.IsNullOrEmpty(userID))
            {
                securityHandler = new WopiSecurityHandler();
                if (!securityHandler.Exists(userID))
                {
                    securityHandler.AddPrincipal(userID);
                }
            }
            else
            {
                var anonymous = ANONYMOUS;
                if (null != WopiOptions.Value.CMSSSOEnabled && WopiOptions.Value.CMSSSOEnabled == TRUE)
                {
                    var aUser = getLoggedInUser();
                    if (!String.IsNullOrEmpty(aUser) && aUser != anonymous) userID = aUser;

                }
                else
                {
                    userID = getLoggedInUserName(anonymous);
                }
            }
            var version = editSessions.getDocVersionString(id, userID);
            string existingLock = new string(FALSE);
            if (version != "0")
                existingLock = flushCache(userID, id, version);

            //unsavedFileExist = true
            //Newtonsoft.Json.JsonSerializer serializer = new Newtonsoft.Json.JsonSerializer();
            //serializer.
            //ViewData["unsavedFileExist"]=TRUE;

            //var response = Microsoft.AspNetCore.Http.HttpRequest.
            //var response =  request.CreateResponse(HttpStatusCode.OK, "unsavedFileExist = true");

            // Set headers for paging
            //response.Headers.Add("content-type","application/json");

            var unsavedFileExist = FALSE;
            if (id == null) return await handleErrorMessage("400", "missing required parameter attachVersionID", "error");

            if (version == "0")
            {
                return Ok(new
                {
                    //unsavedFileExist = $"{unsavedFileExist} {HttpContext.User}"
                    //unsavedFileExist = $"{HttpContext.User.ToString()}"
                    status = "0",
                    unsavedFileExist = FALSE,
                    fileLock = FALSE
                });
            }

            var fileName = getTempFile(userID, id, version) + WopiOptions.Value.WordExt;
            if (System.IO.File.Exists(WopiOptions.Value.RootPath + fileName))
            {
                var createTime = System.IO.File.GetCreationTime(WopiOptions.Value.RootPath + fileName);
                var lastWriteTime = System.IO.File.GetLastWriteTime(WopiOptions.Value.RootPath + fileName);
                var tolerance = createTime.AddSeconds(1);



                if (lastWriteTime.CompareTo(tolerance) <= 0)
                {
                    unsavedFileExist = FALSE;
                }
                else
                {
                    unsavedFileExist = TRUE;
                    //updateSessions(getTempFile(userID, id, version), DocState.CHANGES_PENDING, DateTime.Now);
                }


                //if (existingLock == TRUE)
                //{
                //    unsavedFileExist = FALSE;
                //}

            }
            //System.Security.Claims.ClaimsPrincipal user = HttpContext.User;

            HttpContext.Response.Headers["content-type"] = "application/json";
            return Ok(new
            {
                //unsavedFileExist = $"{unsavedFileExist} {HttpContext.User}"
                //unsavedFileExist = $"{HttpContext.User.ToString()}"
                status = "0",
                unsavedFileExist = $"{unsavedFileExist}",
                fileLock = $"{existingLock}"
            });
            //return await handleErrorMessage("400","A 400 error","error");
            //return await handleErrorMessage("404", "A 404 error", "error");
            //return await handleErrorMessage("405", "A 405 error", "error");
            //return await handleErrorMessage("505", "A 505 error", "error");
        }

        [EnableCors("AllowAll")]
        [Microsoft.AspNetCore.Mvc.Route("WordWeb/Discard")]
        [AllowAnonymous]
        public async Task<ActionResult> Discard(string id, [FromQuery] string attachVersionID, [FromQuery] string userID)
        {
            if (!String.IsNullOrEmpty(attachVersionID))
            {
                id = attachVersionID;
            }
            if (id == null) return await handleErrorMessage("400", "missing required parameter attachVersionID", "error");
            WopiSecurityHandler securityHandler = new WopiSecurityHandler();
            if (!String.IsNullOrEmpty(userID))
            {
                securityHandler = new WopiSecurityHandler();
                if (!securityHandler.Exists(userID))
                {
                    securityHandler.AddPrincipal(userID);
                }
            }
            else
            {
                var anonymous = ANONYMOUS;
                if (null != WopiOptions.Value.CMSSSOEnabled && WopiOptions.Value.CMSSSOEnabled == TRUE)
                {
                    var aUser = getLoggedInUser();
                    if (!String.IsNullOrEmpty(aUser) && aUser != anonymous) userID = aUser;

                }
                else
                {
                    userID = getLoggedInUserName(anonymous);
                }
            }
            //unsavedFileExist = true
            //Newtonsoft.Json.JsonSerializer serializer = new Newtonsoft.Json.JsonSerializer();
            //serializer.
            //ViewData["unsavedFileExist"]=TRUE;

            //var response = Microsoft.AspNetCore.Http.HttpRequest.
            //var response =  request.CreateResponse(HttpStatusCode.OK, "unsavedFileExist = true");

            // Set headers for paging
            //response.Headers.Add("content-type","application/json");
            var versionId = editSessions.getDocVersionString(id, userID);
            WordWebException wwex = new WordWebException();

            string existingLock = flushCache(userID, id, versionId);
            bool deleted = false;
            if (existingLock == TRUE)
            {
                return await handleErrorMessage("400", "The file is currently locked by a user for editing.", "error");
            }

            var fileName = getTempFile(userID, id, versionId) + WopiOptions.Value.WordExt;
            if (System.IO.File.Exists(WopiOptions.Value.RootPath + fileName))
            {
                System.IO.File.Delete(WopiOptions.Value.RootPath + fileName);
                deleted = true;
                editSessions.rmDocVersion(id, userID, versionId, wwex);
            }
            //System.Security.Claims.ClaimsPrincipal user = HttpContext.User;

            HttpContext.Response.Headers["content-type"] = "application/json";
            if (deleted)
            {
                return Ok(new
                {
                    status = "0",
                    timpstamp = String.Format("{0:G}", DateTime.Now),
                    deletedFile = fileName
                });
            }
            else
            {
                return Ok(new
                {
                    status = "1",
                    timpstamp = String.Format("{0:G}", DateTime.Now),
                    message = "Temp file " + fileName + " does not exist. No deletion occurred."
                });
            }
            //return await handleErrorMessage("400","A 400 error","error");
            //return await handleErrorMessage("404", "A 404 error", "error");
            //return await handleErrorMessage("405", "A 405 error", "error");
            //return await handleErrorMessage("505", "A 505 error", "error");
        }


        public async Task<ActionResult> View(string id, string fileName)
        {
            WopiSecurityHandler securityHandler = new WopiSecurityHandler();
            IWopiFile file = StorageProvider.GetWopiFile(id);
            var token = securityHandler.GenerateAccessToken(ANONYMOUS, file.Identifier);
            //var token = securityHandler.GenerateAccessToken(getLoggedInUserName(userID), file.Identifier);


            ViewData["access_token"] = securityHandler.WriteToken(token);
            //TODO: fix
            //ViewData["access_token_ttl"] = WopiOptions.Value.SessionTimeout; //token.ValidTo;
            ViewData["access_token_ttl"] = getTTL(); //token.ValidTo;
            //ViewData[WopiOptions.Value.DocumentProtectionKey] = TRUE;
            /*var docProtKey = fileName;
            if (null != docProtKey && docProtection.ContainsKey(docProtKey))
            {
                string docProt = null;
                if (docProtection.TryGetValue(docProtKey, out docProt))
                    ViewData[WopiOptions.Value.DocumentProtectionKey] = TRUE;
                else
                    ViewData[WopiOptions.Value.DocumentProtectionKey] = FALSE;
            }
            else
            {
                ViewData[WopiOptions.Value.DocumentProtectionKey] = FALSE;
            }*/

            var extension = file.Extension.TrimStart('.');
            string urlsrc = await UrlGenerator.GetFileUrlAsync(extension, $"{WopiOptions.Value.HostUrl}/wopi/files/{id}", WopiActionEnum.View);
            //ViewData["urlsrc"] = WopiUrlBuilder.RepalceDomainNameInURL(urlsrc, "http://ld449820:9000/owa");
            ViewData["urlsrc"] = WopiUrlBuilder.RepalceDomainNameInURL(urlsrc, WopiOptions.Value.ClientUrl);
            //ViewData["urlsrc"] = WopiUrlBuilder.RepalceDomainNameInURL(urlsrc, WopiOptions.Value.ClientUrl);
            ViewData["favicon"] = await Discoverer.GetApplicationFavIconAsync(extension);
            return View();
        }


        [EnableCors("AllowAll")]
        [Microsoft.AspNetCore.Mvc.Route("WordWeb/Detail")]
        [AllowAnonymous]
        public async Task<ActionResult> Detail(string id, string attachVersionID, string mode, string docSrc, string userID, string fileName = null, string errCode = null, string errMsg = null, string errStack = null)
        {
            WopiSecurityHandler securityHandler = new WopiSecurityHandler();

            if (errCode != null)
            {
                ViewData["error_code"] = "400";
                ViewData["error_message"] = "The document id is invalid and cannot be retrieved from CMS.";
                ViewData["error_stack"] = "";
                return View();
            }


            IWopiFile file = StorageProvider.GetWopiFile(id);
            //var token = securityHandler.GenerateAccessToken(ANONYMOUS, file.Identifier);
            var token = securityHandler.GenerateAccessToken(getLoggedInUserName(userID), file.Identifier);


            ViewData["access_token"] = securityHandler.WriteToken(token);
            //TODO: fix
            //ViewData["access_token_ttl"] = WopiOptions.Value.SessionTimeout; //token.ValidTo;
            ViewData["access_token_ttl"] = getTTL(); //token.ValidTo;
            if (null != fileName) ViewData["doc_session_id"] = fileName; //token.ValidTo;

            var docProtKey = fileName;
            if (null != docProtKey && docProtection.ContainsKey(docProtKey))
            {
                string docProt = null;
                if (docProtection.TryGetValue(docProtKey, out docProt))
                {
                    if (WopiOptions?.Value?.DocumentProtectionNotification == TRUE)
                        ViewData[WopiOptions.Value.DocumentProtectionKey] = TRUE;
                    else
                        ViewData[WopiOptions.Value.DocumentProtectionKey] = FALSE;
                }
                else
                    ViewData[WopiOptions.Value.DocumentProtectionKey] = FALSE;
            }
            else
            {
                ViewData[WopiOptions.Value.DocumentProtectionKey] = FALSE;
            }

            var versionId = editSessions.getDocVersionString(attachVersionID, userID);
            var extension = file.Extension.TrimStart('.');

            string urlsrc;
            if (!(String.IsNullOrWhiteSpace(mode)) && mode.ToLower() == "read_only")
            {
                urlsrc = await UrlGenerator.GetFileUrlAsync(extension, $"{WopiOptions.Value.HostUrl}/wopi/files/{id}", WopiActionEnum.View);
            }
            else
            {
                urlsrc = await UrlGenerator.GetFileUrlAsync(extension, $"{WopiOptions.Value.HostUrl}/wopi/files/{id}", WopiActionEnum.Edit);
                updateSessions(getTempFile(userID, id, versionId), DocState.EDIT_STARTED, DateTime.Now);
            }
            //ViewData["urlsrc"] = WopiUrlBuilder.RepalceDomainNameInURL(urlsrc, "http://ld449820:9000/owa");
            ViewData["urlsrc"] = WopiUrlBuilder.RepalceDomainNameInURL(urlsrc, WopiOptions.Value.ClientUrl);
            ViewData["favicon"] = await Discoverer.GetApplicationFavIconAsync(extension);

            return View();
        }

        [EnableCors("AllowAll")]
        [Microsoft.AspNetCore.Mvc.Route("WordWeb/ErrorView")]
        [AllowAnonymous]
        public async Task<ActionResult> ErrorView(WordWebException wwex = null, WordWebAction wwa = 0)
        {

            if (wwex != null)
            {
                ViewData["error_code"] = wwex.errMessage.code;
                ViewData["error_message"] = wwex.errMessage.message;
                ViewData["error_level"] = wwex.errMessage.level;
                switch (wwex.errMessage.area)
                {
                    case ErrorArea.VIEW:
                        ViewData["error_area"] = "VIEW";
                        break;
                    case ErrorArea.JSON:
                    default:
                        ViewData["error_area"] = "JSON";
                        break;
                }
                ViewData["error_timestamp"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                /*DEFAULT = 0,
                RUNTIME = 1,
                VIEW = 2,
                EDIT = 3,
                SAVE = 4,
                UNSAVED = 5,
                DISCARD = 6
                */

                if (wwa != null)
                {
                    switch (wwa)
                    {
                        case WordWebAction.EDIT:
                            ViewData["web_action"] = "EDIT_DOCUMENT";
                            break;
                        case WordWebAction.VIEW:
                            ViewData["web_action"] = "VIEW_DOCUMENT";
                            break;
                        case WordWebAction.SAVE:
                            ViewData["web_action"] = "SAVE_DOCUMENT";
                            break;
                        case WordWebAction.UNSAVED:
                            ViewData["web_action"] = "CHECK_FOR_UNSAVED_CHANGES";
                            break;
                        case WordWebAction.DISCARD:
                            ViewData["web_action"] = "DISCARD_DOCUMENT";
                            break;
                        case WordWebAction.RUNTIME:
                        default:
                            ViewData["web_action"] = "UNABLE_TO_RETRIEVE_REQUESTED_INFO_FROM_CMS";
                            break;
                    }
                }
            }
            else
            {
                ViewData["error_code"] = "500";
                ViewData["error_message"] = "Error in retrieving exception messages from CMS";
                ViewData["error_level"] = "error";
                ViewData["error_area"] = "VIEW";
                ViewData["error_timestamp"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                ViewData["web_action"] = "RUNTIME";
            }
            return View();
        }


        [AllowAnonymous]
        public ActionResult Edit(string id, [FromQuery] string attachVersionID, [FromQuery] string mode, [FromQuery] string docSource, [FromQuery] string userID, [FromQuery] string userRole)
        {
            WopiSecurityHandler securityHandler = new WopiSecurityHandler();
            ///WordWeb/Edit/Llw2Njk5MTA1ODcuZG9jeA%3D%3D
            //"Llw2Njk5MTA1ODcuZG9jeA=="
            var docMode = new String("");
            var docSrc = new String("");



            try
            {
                ViewData.Remove(WopiOptions.Value.DocumentProtectionKey);

                bool readResult = false;
                WordWebException ww = new WordWebException();
                Attachment attachedDoc = new Attachment();
                //DateTime currentIime = DateTime.Now;
                //DateTime oneMinute = currentIime.AddMinutes(1);

                if (!String.IsNullOrEmpty(userRole))
                {
                    userRole = userRole.ToLower();
                }
                if (!String.IsNullOrEmpty(mode))
                {
                    docMode = mode.ToLower();
                }
                if (!String.IsNullOrEmpty(docSource))
                {
                    docSrc = docSource.ToLower();
                }
                if (!String.IsNullOrEmpty(attachVersionID))
                {
                    id = attachVersionID;
                }
                if (!String.IsNullOrEmpty(userID))
                {
                    securityHandler = new WopiSecurityHandler();
                    if (!securityHandler.Exists(userID))
                    {
                        securityHandler.AddPrincipal(userID);
                    }
                }
                else
                {
                    var anonymous = ANONYMOUS;
                    if (null != WopiOptions.Value.CMSSSOEnabled && WopiOptions.Value.CMSSSOEnabled == TRUE)
                    {
                        var aUser = getLoggedInUser();
                        if (!String.IsNullOrEmpty(aUser) && aUser != anonymous) userID = aUser;

                    }
                    else
                    {
                        userID = getLoggedInUserName(anonymous);
                    }
                }

                var lastVersionId = editSessions.getDocVersionString(id, userID);
                if (lastVersionId != "0" && !String.IsNullOrEmpty(lastVersionId) && docSrc != "temp")
                {
                    var lastDoc = getTempFile(userID, id, lastVersionId);
                    if (getDocSessionState(lastDoc) == DocState.SAVE_PENDING)
                    {
                        //make sure the contents is flushed
                        //var attachedFile = @"..\\WopiHost\\wwwroot\\wopi-docs\\"+attachmentID + extension;

                        DateTime startTime = getDocLastModified(lastDoc);
                        DateTime endTime = System.IO.File.GetLastWriteTime(WopiOptions.Value.RootPath + lastDoc + WopiOptions.Value.WordExt);

                        if (endTime == null) endTime = pointInTime;

                        TimeSpan timeSpan = endTime.Subtract(startTime);

                        if (timeSpan.TotalMilliseconds <= 0)
                        {


                            System.Threading.Tasks.Task cacheFlusher = null;
                            try
                            {
                                cacheFlusher = System.Threading.Tasks.Task.Factory.StartNew(() =>
                                {
                                    //cacheFlusher = System.Threading.Tasks.Task.Factory.StartNew(() => {                                    
                                    var watchFile = WopiOptions.Value.RootPath + lastDoc + WopiOptions.Value.WordExt;
                                    var convertedFile = WopiOptions.Value.RootPath + lastDoc + WopiOptions.Value.Word2010Ext;
                                    // Console.Out.WriteLine(System.IO.File.GetLastWriteTime(watchFile));

                                    DateTime lastModified = System.IO.File.GetLastWriteTime(watchFile);
                                    DateTime lastConverted;
                                    if (System.IO.File.Exists(convertedFile)) lastConverted = System.IO.File.GetLastWriteTime(convertedFile);

                                    var watcher = new FileWatcher(Path.GetFullPath(watchFile), getDocLastModified(lastDoc), this);
                                    int eventAllowance = 2;

                                    if (!String.IsNullOrEmpty(WopiOptions.Value.FileSystemEventTimeout))
                                        eventAllowance = Int32.Parse(WopiOptions.Value.FileSystemEventTimeout);

                                    DateTime allowedTime = DateTime.Now.AddSeconds(eventAllowance);
                                    //Console.Out.WriteLine("Max_Cutoff: " + allowedTime.ToString());
                                    int i = 0;
                                    do
                                    {
                                        if (DateTime.Now.CompareTo(allowedTime) > 0) break;
                                        i += 1;
                                        //Console.Out.WriteLine(i.ToString());
                                        //Console.Out.WriteLine(DateTime.Now.ToString());
                                        //Console.Out.WriteLine("File timestamp:" + System.IO.File.GetLastWriteTime(watchFile));
                                        //Console.Out.WriteLine("isCacheFlushed: false");
                                    } while (!isCacheFlushed(lastDoc));
                                    Console.Out.WriteLine("isCacheFlushed:" + isCacheFlushed(lastDoc).ToString());
                                    watcher.Dispose();
                                }, TaskCreationOptions.LongRunning | TaskCreationOptions.PreferFairness);

                                System.Threading.Tasks.Task[] tasks = new System.Threading.Tasks.Task[] { cacheFlusher };
                                System.Threading.Tasks.Task.WaitAll(tasks);

                                //System.Threading.Tasks.Task actualCacheFlusher  =cacheFlusher.Unwrap();
                                //System.Threading.Tasks.Task.WaitAll(cacheFlusher);
                                //cacheFlusher.Wait();

                            }
                            catch (Exception ex)
                            {
                                return ErrorView(new WordWebException("Event error. Office is overloaded. Please try again later."), WordWebAction.SAVE).Result;
                            }
                            finally
                            {
                                if (cacheFlusher != null)
                                    System.Threading.Tasks.Task.WaitAll(cacheFlusher);

                            }
                        }
                        Thread.Sleep(1000);
                        WordWebException wwe = new WordWebException();
                        string updatedVersion = checkPendingSave(id, attachVersionID, userID, wwe);
                        //updateSessions(getTempFile(userID, id, lastVersionId), DocState.EDIT_COMPLETED, DateTime.Now);

                        if (updatedVersion == "-1" && null != wwe && null != wwe.errMessage)
                        {
                            updateSessions(getTempFile(userID, id, lastVersionId), DocState.EDIT_COMPLETED, DateTime.Now);
                            return ErrorView(wwe, WordWebAction.SAVE).Result;
                        }


                        if (updatedVersion != "-1" && !String.IsNullOrEmpty(updatedVersion))
                        {
                            updateSessions(getTempFile(userID, id, lastVersionId), DocState.EDIT_COMPLETED, DateTime.Now);
                        }


                        if (updatedVersion != "0" && !String.IsNullOrEmpty(updatedVersion))
                        {
                            updateSessions(getTempFile(userID, id, lastVersionId), DocState.EDIT_COMPLETED, DateTime.Now);
                            updateSessions(getTempFile(userID, id, updatedVersion), DocState.BEFORE_EDIT, DateTime.Now);
                        }
                    }
                }

                if (docSrc == "temp")
                {
                    //readResult = readAttachment(id, userID, true, WordWebAction.EDIT, attachedDoc, ww).Result;
                    readResult = true;
                    var versionId = editSessions.getDocVersionString(id, userID);
                    if (versionId != "0")
                    {
                        attachedDoc.versionNo = versionId;
                    }
                    else
                    {
                        readResult = readAttachment(id, userID, true, WordWebAction.EDIT, attachedDoc, ww, userRole).Result;
                    }
                }
                else
                {
                    readResult = readAttachment(id, userID, false, WordWebAction.EDIT, attachedDoc, ww, userRole).Result;
                }

                if (!readResult)
                {
                    if (docMode == "read_only")
                        return ErrorView(ww, WordWebAction.VIEW).Result;
                    else
                        return ErrorView(ww, WordWebAction.EDIT).Result;
                }


                var tempFile = getTempFile(userID, id);
                if (readResult)
                {
                    tempFile = getTempFile(userID, id, attachedDoc.versionNo);
                    if (!addEditSessions(DateTime.Now, id, userID, attachedDoc.versionNo, ww))
                    {
                        if (docMode == "read_only")
                            return ErrorView(ww, WordWebAction.VIEW).Result;
                        else
                            return ErrorView(ww, WordWebAction.EDIT).Result;
                    }
                }



                if (docSessions.ContainsKey(tempFile))
                {
                    DocSession currentSession = docSessions[tempFile];
                    if (currentSession.docState.Equals(DocState.CHANGES_PENDING))
                    {
                        if (docSrc != "temp")
                        //Remove the temp file
                        {
                            updateSessions(tempFile, DocState.CHANGES_DISCARD, DateTime.Now);
                        }
                        else
                        {
                            updateSessions(tempFile, DocState.CHANGES_RESUMED, DateTime.Now);
                        }
                    }
                }
                else
                {
                    updateSessions(tempFile, DocState.BEFORE_EDIT, DateTime.Now);
                }


                //var tempFile = getTempFile(id);
                IEnumerable<IWopiFile> wopiFiles = StorageProvider.GetWopiFiles(StorageProvider.RootContainerPointer.Identifier);
                string wopiFileId = null;
                IWopiFile wopiFile = null;
                wopiFileId = getWopiFileId(tempFile, out wopiFile);

                //wopiFile = wopiFile.Replace("=", "%3D");
                //flushCache(userID, id);
                return Detail(wopiFileId, attachVersionID, docMode, docSrc, userID, tempFile).Result;

            } /*catch(WordWebException wwex)
            {
                //return await Detail(null, null, null, userID, "400", "The document id is invalid and cannot be retrieved from CMS.");  
                return await ErrorView(wwex);
            }
            catch (Exception ex)
            {
                //return await Detail(null, null, null, userID, "400", "The document id is invalid and cannot be retrieved from CMS.");  
                return await ErrorView();
            }*/
            catch (Exception e)
            {
                throw e;
            }
        }

        public async Task<bool> readAttachment(string attachmentID, string userID, Boolean useTemp, WordWebAction action, Attachment attachment = null, WordWebException ww = null, string userRole = null)
        {
            if (action == null) action = WordWebAction.EDIT;
            readClient.BaseAddress = new Uri($"{ WopiOptions.Value.CMSUrl }");
            readClient.DefaultRequestHeaders.Accept.Clear();
            //readClient.DefaultRequestHeaders.Accept.Add(
            //    new MediaTypeWithQualityHeaderValue("application/msword"));
            readClient.DefaultRequestHeaders.Add("User-Agent", "CMS.Word.Web.Application.Client");
            readClient.DefaultRequestHeaders.Add("Referer", $"{ WopiOptions.Value.CMSReferer }");

            var defaultExtension = WopiOptions.Value.Word2010Ext;
            var multiContent = new MultipartFormDataContent();

            //multiContent.Add(fileStreamContent, "fileToUpload", file.FileName);
            //multiContent.Add(new StringContent(formData.id.ToString()), "id");

            multiContent.Add(new StringContent(getUserForCMSAPI(userID)), "j_username");
            multiContent.Add(new StringContent(getCMSServicePwd()), "j_password");
            multiContent.Add(new StringContent("EXTERNAL"), "user_type");

            string mimeType = "";
            string fileContent = "";
            string versionNo = "";
            bool secEnabled = false;
            if ($"{ WopiOptions.Value.CMSRestSecurityEnabled}" == TRUE)
            {
                secEnabled = true;
            }

            //var parameters = new Dictionary<string, string> { { "j_username", getUserForCMSAPI(userID) }, { "j_password", getCMSServicePwd() }, { "user_type", "EXTERNAL" } };
            var parameters = new Dictionary<string, string> { { "j_username", getUserForCMSAPI(userID) }, { "j_password", getCMSServicePwd() }, { "user_type", "" } };
            var encodedContent = new FormUrlEncodedContent(parameters);

            //var response = await readClient.PostAsync("/Rest/j_security_check?j_username=RS17806&j_password=*******", encodedContent);
            //var response = await readClient.GetAsync("/Rest/j_security_check?j_username=RS17806&j_password=*******&user_type=EXTERNAL");

            HttpResponseMessage response = null;

            if (secEnabled)
            {
                response = await readClient.PostAsync("/Rest/j_security_check", encodedContent).ConfigureAwait(false);
            }
            if (!secEnabled || (null != response && response.IsSuccessStatusCode))
            {
                //readClient.DefaultRequestHeaders.Accept.Clear();
                //readClient.DefaultRequestHeaders.Add("User-Agent", "CMS Word Web Application readClient");
                //readClient.DefaultRequestHeaders.Add("Referer", "http://appd13was");
                //var streamTask = readClient.GetStreamAsync("http://appd13was:9044/Rest/v1/docmgmt/readAttachmentVersionContent?attachVersionID=669891871");
                var streamResult = await readClient.GetAsync($"{ WopiOptions.Value.CMSUrl }" + $"/Rest/v1/docmgmt/readAttachmentVersionContent?{WopiOptions.Value.CMSAttachVersionId}=" + attachmentID);
                //using (var streamResult = await readClient.GetAsync("http://appd13was:9044/Rest/v1/docmgmt/readAttachmentVersionContent?attachVersionID=669891871"))
                //{
                //Attachment attachment = await streamTask;
                //var streamResult = await streamTask;
                string apiResponse = await streamResult.Content.ReadAsStringAsync();
                //byte[] apiResponse = await streamResult.Content.ReadAsByteArrayAsync();

                //MemoryStream ms = new MemoryStream(apiResponse);
                //using (BsonDataReader reader = new BsonDataReader(ms))
                //{
                //   Newtonsoft.Json.JsonSerializer serializer = new Newtonsoft.Json.JsonSerializer();
                //   attachment = serializer.Deserialize<Attachment>(reader);
                //}
                //Attachment myattachment = JsonConverter.DeserializeObject<Attachment>(apiResponse);
                //var bytesAsString = Encoding.UTF8.GetString(apiResponse);
                // attachment = JsonConvert.DeserializeObject<Attachment>(apiResponse);
                //Attachment myattachment = JsonSerializer.Deserialize<Attachment>(apiResponse);

                JObject _attachment = JObject.Parse(apiResponse);
                JToken errorsJson = null;

                if (_attachment.TryGetValue("errors", out errorsJson))
                {
                    JArray errorsArray = JArray.Parse(errorsJson.ToString());

                    IList<ErrorMessage> error = errorsArray.Select(err => new ErrorMessage
                    {
                        code = (string)err["code"],
                        message = (string)err["message"],
                        level = (string)err["level"],
                        area = ErrorArea.VIEW
                    }).ToList();

                    if (error.Count >= 1)
                    {
                        //throw new CMSException(error[0]);
                        /*throw new WordWebException
                        {
                            errMessage = error[0]                            
                        };*/
                        // null;

                        /*ww = new WordWebException
                        {
                            errMessage = error[0]
                        };*/
                        ww.errMessage = error[0];
                        attachment = null;
                        return false;

                    }
                    else
                    {
                        ErrorMessage err = new ErrorMessage
                        {
                            code = "400",
                            message = "Unable to retrieve the attachment " + attachmentID + " from CMS",
                            level = "error",
                            area = ErrorArea.VIEW
                        };
                        ww.errMessage = err;
                        attachment = null;
                        return false;
                    }

                }
                else
                {


                    mimeType = _attachment.GetValue("mimeType").ToString();
                    fileContent = _attachment.GetValue("contentsBlob").ToString();
                    versionNo = _attachment.GetValue("attachVersionVrsNo").ToString();
                    //string fileContent = attachment.GetValue("fileContent").ToString();
                    //Console.WriteLine(attachment.GetValue("mimeType").ToString());
                    //var data1 = Encoding.UTF8.GetString(attachment.GetValue("fileContent").ToString());
                    //Console.WriteLine(attachment.GetValue("mimeType").ToString());
                    //Console.WriteLine(attachment.GetValue("fileContent").ToString());
                    //Console.WriteLine(attachment.GetValue("versionNo").ToString());

                    //attachment = JsonConvert.DeserializeObject<Attachment>(apiResponse);
                    //var utf8 = Encoding.UTF8;
                    //byte[] utfBytes = utf8.GetBytes(fileContent);
                    //var myString = utf8.GetString(utfBytes, 0, utfBytes.Length);
                    //byte[] bytes = Encoding.Default.GetBytes(fileContent);
                    //myString = Encoding.b.GetString(bytes);
                    var fileLength = fileContent.Length;
                    //string newfileContent = "encoded_file_content";
                    //byte[] bytes = Encoding.ASCII.GetBytes(fileContent);
                    //byte[] bytes = Encoding.ASCII.GetBytes(newfileContent);
                    //char[] encodedFile = Encoding.ASCII.GetChars(bytes);
                    //byte[] data = Convert.FromBase64CharArray((char[])encodedFile, 0, encodedFile.Length);
                    //byte[] data = Encoding.ASCII.GetBytes(fileContent);

                    byte[] data = Convert.FromBase64String(fileContent);

                    var extension = getMimeExtension(mimeType);
                    extension = (extension == null) ? defaultExtension : extension;

                    //var attachedFile = @"..\\WopiHost\\wwwroot\\wopi-docs\\"+attachmentID + extension;
                    var attachedFile = WopiOptions.Value.RootPath + getTempFile(userID, attachmentID, versionNo) + extension;

                    if (System.IO.File.Exists(attachedFile))
                    {
                        System.IO.File.Delete(attachedFile);
                    }

                    System.IO.File.WriteAllBytes(attachedFile, data);
                    if (extension == WopiOptions.Value.Word2010Ext)
                    {
                        //convertDocToDocxOpenXML(attachedFile, useTemp);
                        //convertDocToDocx(attachedFile, useTemp);
                        //convertDocToDocxAspose(attachedFile, useTemp);
                        //convertDocToDocxGlue(attachedFile, useTemp);
                        //convertDocToDocxSpire(attachedFile, useTemp);
                        //convertDOCMtoDOCX(attachedFile, useTemp);
                        var upConversionEngine = "convertDocToDocx";
                        var upConversionEngineFallback = "convertDocToDocx";

                        if (String.IsNullOrEmpty(WopiOptions.Value.ConversionEngine))
                        {
                            if (String.IsNullOrEmpty(WopiOptions.Value.ConversionEngineFallBack))
                            {
                                upConversionEngineFallback += WopiOptions.Value.ConversionEngineFallBack; //default
                                upConversionEngine = upConversionEngineFallback; //default
                            }
                        }
                        else
                        {
                            upConversionEngine += WopiOptions.Value.ConversionEngine;
                        }
                        MethodInfo conversionMethod = this.GetType().GetMethod(upConversionEngine);
                        MethodInfo conversionMethodFallback = this.GetType().GetMethod(upConversionEngineFallback);
                        object result = null;

                        try
                        {
                            conversionMethod.Invoke(this, new object[] { attachedFile, useTemp, userRole });
                        }
                        catch (Exception exc)
                        {
                            conversionMethodFallback.Invoke(this, new object[] { attachedFile, useTemp, userRole });
                        }

                    }
                    else
                    {
                        if (extension != WopiOptions.Value.WordExt)
                        {
                            convertType(attachedFile, WopiOptions.Value.WordExt);
                        }
                    }




                    // var mime = getMimeExtension(mimeType);

                    //string mime = getMimeFromFile(@"c:\\wopi\\wopi_test_base64.doc");
                    //Console.WriteLine(mime);

                    //byte[] data = Convert.FromBase64String(fileContent);
                    //byte[] data = Encoding.UTF8.GetBytes(fileContent);
                    //System.IO.File.WriteAllBytes("c:\\wopi\\wopi_test_base64.docx", data);

                    //byte[] data = Convert.FromBase64String(fileContent);
                    //byte[] data = Encoding.UTF8.GetBytes(fileContent);
                    //System.IO.File.WriteAllBytes("c:\\wopi\\wopi_test_base64.docx", data);

                    //var streamTask = readClient.GetStreamAsync("http://appd13was:9044/Rest/v1/docmgmt/readAttachmentVersionContent?attachVersionID=669891871");
                    //Attachment attachment = await streamTask;
                    // Attachment attachment = await JsonSerializer.DeserializeAsync<Attachment>(await streamTask);
                    //Console.WriteLine(attachment.mimeType);
                }
            }
            else
            {
                throw new WebException("unauthorized access to CMS");
            }

            theAttachment.fileContent = fileContent;
            theAttachment.mimeType = mimeType;
            theAttachment.versionNo = versionNo;
            attachment.fileContent = fileContent;
            attachment.mimeType = mimeType;
            attachment.versionNo = versionNo;

            return true;
        }


        public async Task<Attachment> getAttachment(string attachmentID, WordWebException ww, string userID = "RESTAPI")
        {
            readClient.BaseAddress = new Uri($"{ WopiOptions.Value.CMSUrl }");
            readClient.DefaultRequestHeaders.Accept.Clear();
            //readClient.DefaultRequestHeaders.Accept.Add(
            //    new MediaTypeWithQualityHeaderValue("application/msword"));
            readClient.DefaultRequestHeaders.Add("User-Agent", "CMS.Word.Web.Application.Client");
            readClient.DefaultRequestHeaders.Add("Referer", $"{ WopiOptions.Value.CMSReferer }");

            var defaultExtension = WopiOptions.Value.Word2010Ext;
            var multiContent = new MultipartFormDataContent();

            //multiContent.Add(fileStreamContent, "fileToUpload", file.FileName);
            //multiContent.Add(new StringContent(formData.id.ToString()), "id");

            multiContent.Add(new StringContent(getUserForCMSAPI(userID)), "j_username");
            multiContent.Add(new StringContent(getCMSServicePwd()), "j_password");
            multiContent.Add(new StringContent("EXTERNAL"), "user_type");

            string mimeType = "";
            string fileContent = "";
            string versionNo = "";
            bool secEnabled = false;
            if ($"{ WopiOptions.Value.CMSRestSecurityEnabled}" == TRUE)
            {
                secEnabled = true;
            }


            //var parameters = new Dictionary<string, string> { { "j_username", getUserForCMSAPI(userID) }, { "j_password", getCMSServicePwd() }, { "user_type", "EXTERNAL" } };
            var parameters = new Dictionary<string, string> { { "j_username", getUserForCMSAPI(userID) }, { "j_password", getCMSServicePwd() }, { "user_type", "" } };
            var encodedContent = new FormUrlEncodedContent(parameters);

            //var response = await readClient.PostAsync("/Rest/j_security_check?j_username=RS17806&j_password=*******", encodedContent);
            //var response = await readClient.GetAsync("/Rest/j_security_check?j_username=RS17806&j_password=********&user_type=EXTERNAL");


            HttpResponseMessage response = null;

            if (secEnabled)
            {
                response = await readClient.PostAsync("/Rest/j_security_check", encodedContent);
            }
            if (!secEnabled || (null != response && response.IsSuccessStatusCode))
            {



                //readClient.DefaultRequestHeaders.Accept.Clear();
                //readClient.DefaultRequestHeaders.Add("User-Agent", "CMS Word Web Application readClient");
                //readClient.DefaultRequestHeaders.Add("Referer", "http://appd13was");
                //var streamTask = readClient.GetStreamAsync("http://appd13was:9044/Rest/v1/docmgmt/readAttachmentVersionContent?attachVersionID=669891871");
                var streamResult = await readClient.GetAsync($"{ WopiOptions.Value.CMSUrl }" + $"/Rest/v1/docmgmt/readAttachmentVersionContent?{WopiOptions.Value.CMSAttachVersionId}=" + attachmentID);
                //using (var streamResult = await readClient.GetAsync("http://appd13was:9044/Rest/v1/docmgmt/readAttachmentVersionContent?attachVersionID=669891871"))
                //{
                //Attachment attachment = await streamTask;
                //var streamResult = await streamTask;
                string apiResponse = await streamResult.Content.ReadAsStringAsync();
                //byte[] apiResponse = await streamResult.Content.ReadAsByteArrayAsync();

                //MemoryStream ms = new MemoryStream(apiResponse);
                //using (BsonDataReader reader = new BsonDataReader(ms))
                //{
                //   Newtonsoft.Json.JsonSerializer serializer = new Newtonsoft.Json.JsonSerializer();
                //   attachment = serializer.Deserialize<Attachment>(reader);
                //}
                //Attachment myattachment = JsonConverter.DeserializeObject<Attachment>(apiResponse);
                //var bytesAsString = Encoding.UTF8.GetString(apiResponse);
                // attachment = JsonConvert.DeserializeObject<Attachment>(apiResponse);
                //Attachment myattachment = JsonSerializer.Deserialize<Attachment>(apiResponse);



                //JObject attachment = JObject.Parse(apiResponse);
                JObject _attachment = JObject.Parse(apiResponse);
                JToken errorsJson = null;

                if (_attachment.TryGetValue("errors", out errorsJson))
                {
                    JArray errorsArray = JArray.Parse(errorsJson.ToString());

                    IList<ErrorMessage> error = errorsArray.Select(err => new ErrorMessage
                    {
                        code = (string)err["code"],
                        message = (string)err["message"],
                        level = (string)err["level"],
                        area = ErrorArea.VIEW
                    }).ToList();

                    if (error.Count >= 1)
                    {
                        //throw new CMSException(error[0]);
                        /*throw new WordWebException
                        {
                            errMessage = error[0]                            
                        };*/
                        // null;

                        /*ww = new WordWebException
                        {
                            errMessage = error[0]
                        };*/
                        ww.errMessage = error[0];
                        return null;

                    }
                    else
                    {
                        ErrorMessage err = new ErrorMessage
                        {
                            code = "400",
                            message = "Unable to retrieve the attachment " + attachmentID + " from CMS",
                            level = "error",
                            area = ErrorArea.VIEW
                        };
                        ww.errMessage = err;
                        return null;
                    }
                }

                mimeType = _attachment.GetValue("mimeType").ToString();
                fileContent = _attachment.GetValue("contentsBlob").ToString();
                versionNo = _attachment.GetValue("attachVersionVrsNo").ToString();


                //string fileContent = attachment.GetValue("fileContent").ToString();
                //Console.WriteLine(attachment.GetValue("mimeType").ToString());
                //var data1 = Encoding.UTF8.GetString(attachment.GetValue("fileContent").ToString());
                //Console.WriteLine(attachment.GetValue("mimeType").ToString());
                //Console.WriteLine(attachment.GetValue("fileContent").ToString());
                //Console.WriteLine(attachment.GetValue("versionNo").ToString());

                //attachment = JsonConvert.DeserializeObject<Attachment>(apiResponse);
                //var utf8 = Encoding.UTF8;
                //byte[] utfBytes = utf8.GetBytes(fileContent);
                //var myString = utf8.GetString(utfBytes, 0, utfBytes.Length);
                //byte[] bytes = Encoding.Default.GetBytes(fileContent);
                //myString = Encoding.b.GetString(bytes);


                // var mime = getMimeExtension(mimeType);

                //string mime = getMimeFromFile(@"c:\\wopi\\wopi_test_base64.doc");
                //Console.WriteLine(mime);

                //byte[] data = Convert.FromBase64String(fileContent);
                //byte[] data = Encoding.UTF8.GetBytes(fileContent);
                //System.IO.File.WriteAllBytes("c:\\wopi\\wopi_test_base64.docx", data);

                //byte[] data = Convert.FromBase64String(fileContent);
                //byte[] data = Encoding.UTF8.GetBytes(fileContent);
                //System.IO.File.WriteAllBytes("c:\\wopi\\wopi_test_base64.docx", data);

                //var streamTask = readClient.GetStreamAsync("http://appd13was:9044/Rest/v1/docmgmt/readAttachmentVersionContent?attachVersionID=669891871");
                //Attachment attachment = await streamTask;
                // Attachment attachment = await JsonSerializer.DeserializeAsync<Attachment>(await streamTask);
                //Console.WriteLine(attachment.mimeType);

            }
            else
            {
                ErrorMessage err = new ErrorMessage
                {
                    code = "401",
                    message = "Unable to connect to CMS endpoint at this moment.",
                    level = "error",
                    area = ErrorArea.VIEW
                };
                ww.errMessage = err;
                return null;

            }

            theAttachment.fileContent = fileContent;
            theAttachment.mimeType = mimeType;
            theAttachment.versionNo = versionNo;

            return theAttachment;
        }


        public String readAndEncodeFile(String filename)
        {
            byte[] data = System.IO.File.ReadAllBytes(filename);
            string fileContent = Convert.ToBase64String(data);
            return fileContent;
        }

        public async Task<string> saveAttachment(string attachmentID, string fileContent, string versionNo, WordWebException ww = null, string userID = "RESTAPI")
        {
            var updatedVersion = "0";
            try
            {

                saveClient.BaseAddress = new Uri($"{ WopiOptions.Value.CMSUrl }");
                saveClient.DefaultRequestHeaders.Accept.Clear();
                saveClient.DefaultRequestHeaders.Add("User-Agent", "CMS.Word.Web.Application.Client");

                //saveClient.DefaultRequestHeaders.Add("Referer", "http://appd13was");
                //saveClient.BaseAddress.
                //saveClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/msword"));
                //saveClient.DefaultRequestHeaders.Add("User-Agent", "CMS.Word.Web.Application.saveClient");
                saveClient.DefaultRequestHeaders.Add("Referer", $"{WopiOptions.Value.CMSReferer}");

                client = saveClient;

                var defaultExtension = WopiOptions.Value.Word2010Ext;
                var multiContent = new MultipartFormDataContent();

                //multiContent.Add(fileStreamContent, "fileToUpload", file.FileName);
                //multiContent.Add(new StringContent(formData.id.ToString()), "id");

                multiContent.Add(new StringContent(getUserForCMSAPI(userID)), "j_username");
                multiContent.Add(new StringContent(getCMSServicePwd()), "j_password");
                multiContent.Add(new StringContent("EXTERNAL"), "user_type");

                var parameters = new Dictionary<string, string> { { "j_username", getUserForCMSAPI(userID) }, { "j_password", getCMSServicePwd() }, { "user_type", "" } };
                var encodedContent = new FormUrlEncodedContent(parameters);

                string apiResponse = "";

                bool secEnabled = false;
                if ($"{ WopiOptions.Value.CMSRestSecurityEnabled}" == TRUE)
                {
                    secEnabled = true;
                }

                HttpResponseMessage response = null;

                if (secEnabled)
                {
                    //var response = await saveClient.PostAsync("/Rest/j_security_check?j_username=RS17806&j_password=********", encodedContent);
                    response = await saveClient.PostAsync("/Rest/j_security_check", encodedContent);
                }
                if (!secEnabled || (null != response && response.IsSuccessStatusCode))
                {
                    //saveClient.DefaultRequestHeaders.Accept.Clear();
                    //saveClient.DefaultRequestHeaders.Add("User-Agent", "CMS Word Web Application saveClient");
                    //saveClient.DefaultRequestHeaders.Add("Referer", "http://appd13was");
                    //var streamTask = saveClient.GetStreamAsync("http://appd13was:9044/Rest/v1/docmgmt/readAttachmentVersionContent?attachVersionID=669891871");

                    saveClient.DefaultRequestHeaders.Accept.Add(
                    new MediaTypeWithQualityHeaderValue("application/json"));
                    multiContent = new MultipartFormDataContent();
                    multiContent.Add(new StringContent(attachmentID), $"{WopiOptions.Value.CMSAttachVersionId}");
                    multiContent.Add(new StringContent(versionNo), $"{WopiOptions.Value.CMSAttachVersionNo}");
                    multiContent.Add(new StringContent(fileContent), $"{WopiOptions.Value.CMSContentsBlob}");

                    var options = new
                    {
                        attachVersionID = attachmentID,
                        //apikey = ConfigurationManager.AppSettings["word:key"],
                        attachVersionVrsNo = versionNo,
                        //attachVersionVrsNo = "0",
                        contentsBlob = fileContent
                    };

                    // Serialize our concrete class into a JSON String
                    var stringPayload = JsonConvert.SerializeObject(options);
                    var content = new StringContent(stringPayload, Encoding.UTF8, "application/json");

                    //var response = await saveClient.PostAsync("", content)

                    //var parameters = new Dictionary<string, string> { { "attachVersionID", attachmentID }, { "attachVersionVrsNo", versionNo } , { "contentsBlob", fileContent } };
                    //var encodedContent = new FormUrlEncodedContent(parameters);

                    //var streamResult = await saveClient.PutAsync("http://appd13was:9044/Rest/v1/docmgmt/modifyAttachmentVersionContent?attachVersionID=" + attachmentID + "&attachVersionVrsNo=" + versionNo +"&contentsBlob="+fileContent, encodedContent);
                    var streamResult = saveClient.PutAsync($"{ WopiOptions.Value.CMSUrl }" + "/Rest/v1/docmgmt/modifyAttachmentVersionContent", content).Result;


                    //using (var streamResult = await saveClient.GetAsync("http://appd13was:9044/Rest/v1/docmgmt/readAttachmentVersionContent?attachVersionID=669891871"))
                    //{
                    //Attachment attachment = await streamTask;
                    //var streamResult = await streamTask;

                    if (streamResult.IsSuccessStatusCode)
                    {
                        apiResponse = await streamResult.Content.ReadAsStringAsync();
                    }
                    else
                    {

                        /*throw new WebException(streamResult.StatusCode+ "\n"+ streamResult.ToString() + "\n" +streamResult.Content + "\n" + streamResult.RequestMessage + "\n" + streamResult.Headers + "\n" + streamResult.ReasonPhrase);
                         */
                        apiResponse = await streamResult.Content.ReadAsStringAsync();
                        JObject _attachment = JObject.Parse(apiResponse);
                        JToken errorsJson = null;

                        if (_attachment.TryGetValue("errors", out errorsJson))
                        {
                            JArray errorsArray = JArray.Parse(errorsJson.ToString());

                            IList<ErrorMessage> error = errorsArray.Select(err => new ErrorMessage
                            {
                                code = (string)err["code"],
                                message = (string)err["message"],
                                level = (string)err["level"],
                                area = ErrorArea.VIEW
                            }).ToList();

                            if (error.Count >= 1)
                            {
                                //throw new CMSException(error[0]);
                                /*throw new WordWebException
                                {
                                    errMessage = error[0]                            
                                };*/
                                // null;

                                /*ww = new WordWebException
                                {
                                    errMessage = error[0]
                                };*/
                                ww.errMessage = error[0];
                                return "0";

                            }
                            else
                            {
                                ErrorMessage err = new ErrorMessage
                                {
                                    code = "400",
                                    message = "Unable to save the attachment " + attachmentID + " to CMS",
                                    level = "error",
                                    area = ErrorArea.VIEW
                                };
                                ww.errMessage = err;
                                return "0";
                            }

                        }

                        //byte[] apiResponse = await streamResult.Content.ReadAsByteArrayAsync();

                        //MemoryStream ms = new MemoryStream(apiResponse);
                        //using (BsonDataReader reader = new BsonDataReader(ms))
                        //{
                        //   Newtonsoft.Json.JsonSerializer serializer = new Newtonsoft.Json.JsonSerializer();
                        //   attachment = serializer.Deserialize<Attachment>(reader);
                        //}
                        //Attachment myattachment = JsonConverter.DeserializeObject<Attachment>(apiResponse);
                        //var bytesAsString = Encoding.UTF8.GetString(apiResponse);
                        // attachment = JsonConvert.DeserializeObject<Attachment>(apiResponse);
                        //Attachment myattachment = JsonSerializer.Deserialize<Attachment>(apiResponse);

                        //JObject attachment = JObject.Parse(apiResponse);
                        //mimeType = attachment.GetValue("mimeType").ToString();
                        //fileContent = attachment.GetValue("contentsBlob").ToString();
                        //versionNo = attachment.GetValue("attachVersionVrsNo").ToString();
                        //string fileContent = attachment.GetValue("fileContent").ToString();
                        //Console.WriteLine(attachment.GetValue("mimeType").ToString());
                        //var data1 = Encoding.UTF8.GetString(attachment.GetValue("fileContent").ToString());
                        //Console.WriteLine(attachment.GetValue("mimeType").ToString());
                        //Console.WriteLine(attachment.GetValue("fileContent").ToString());
                        //Console.WriteLine(attachment.GetValue("versionNo").ToString());

                        //attachment = JsonConvert.DeserializeObject<Attachment>(apiResponse);
                        //var utf8 = Encoding.UTF8;
                        //byte[] utfBytes = utf8.GetBytes(fileContent);
                        //var myString = utf8.GetString(utfBytes, 0, utfBytes.Length);
                        //byte[] bytes = Encoding.Default.GetBytes(fileContent);
                        //myString = Encoding.b.GetString(bytes);
                        //var fileLength = fileContent.Length;
                        //string newfileContent = "encoded_file_content";
                        //byte[] bytes = Encoding.ASCII.GetBytes(fileContent);
                        //byte[] bytes = Encoding.ASCII.GetBytes(newfileContent);
                        //char[] encodedFile = Encoding.ASCII.GetChars(bytes);
                        //byte[] data = Convert.FromBase64CharArray((char[])encodedFile, 0, encodedFile.Length);
                        //byte[] data = Encoding.ASCII.GetBytes(fileContent);

                        //byte[] data = Convert.FromBase64String(fileContent);

                        //var extension = getMimeExtension(mimeType);
                        //extension = (extension == null) ? defaultExtension : extension;

                        //var attachedFile = @"c:\\wopi\\wopi_test_base64" + extension;

                        //System.IO.File.WriteAllBytes(attachedFile, data);



                        // var mime = getMimeExtension(mimeType);

                        //string mime = getMimeFromFile(@"c:\\wopi\\wopi_test_base64.doc");
                        //Console.WriteLine(mime);

                        //byte[] data = Convert.FromBase64String(fileContent);
                        //byte[] data = Encoding.UTF8.GetBytes(fileContent);
                        //System.IO.File.WriteAllBytes("c:\\wopi\\wopi_test_base64.docx", data);

                        //byte[] data = Convert.FromBase64String(fileContent);
                        //byte[] data = Encoding.UTF8.GetBytes(fileContent);
                        //System.IO.File.WriteAllBytes("c:\\wopi\\wopi_test_base64.docx", data);

                        //var streamTask = saveClient.GetStreamAsync("http://appd13was:9044/Rest/v1/docmgmt/readAttachmentVersionContent?attachVersionID=669891871");
                        //Attachment attachment = await streamTask;
                        // Attachment attachment = await JsonSerializer.DeserializeAsync<Attachment>(await streamTask);
                        //Console.WriteLine(attachment.mimeType);
                    }
                }
                //theAttachment.fileContent = fileContent;
                //theAttachment.mimeType = mimeType;
                //theAttachment.versionNo = versionNo;


                if (apiResponse != null && apiResponse.Length > 0)
                {
                    JObject responseJSON = JObject.Parse(apiResponse);
                    updatedVersion = responseJSON.GetValue(WopiOptions.Value.CMSAttachVersionNo).ToString();
                }
                return updatedVersion;
            }
            catch (WebException ex)
            {
                updatedVersion = "0";
                //throw ex;
                //return updatedVersion;
            }
            catch (Exception e)
            {
                updatedVersion = "0";
                //throw e;
                //return updatedVersion;
            }
            return updatedVersion;
        }




        public static int MimeSampleSize = 256;

        public static string DefaultMimeType = "application/octet-stream";

        [DllImport(@"urlmon.dll", CharSet = CharSet.Auto)]
        private extern static uint FindMimeFromData(
            uint pBC,
            [MarshalAs(UnmanagedType.LPStr)] string pwzUrl,
            [MarshalAs(UnmanagedType.LPArray)] byte[] pBuffer,
            uint cbSize,
            [MarshalAs(UnmanagedType.LPStr)] string pwzMimeProposed,
            uint dwMimeFlags,
            out uint ppwzMimeOut,
            uint dwReserverd
        );

        public static string getMimeFromFile(byte[] bytes)
        {
            //if (!System.IO.File.Exists(filename))
            //    throw new FileNotFoundException(filename + " not found");

            byte[] buffer = new byte[256];
            //using (FileStream fs = new FileStream(filename, FileMode.Open))
            //{
            if (bytes.Length >= 256)
                Buffer.BlockCopy(buffer, 0, bytes, 0, 256);
            else
                Buffer.BlockCopy(buffer, 0, bytes, 0, (int)bytes.Length);
            //}
            try
            {
                uint mimeType;
                FindMimeFromData(0, null, buffer, (uint)MimeSampleSize, null, 0, out mimeType, 0);

                var mimePointer = new IntPtr(mimeType);
                var mime = Marshal.PtrToStringUni(mimePointer);
                Marshal.FreeCoTaskMem(mimePointer);

                return mime ?? DefaultMimeType;
            }
            catch
            {
                return DefaultMimeType;
            }
        }

        //}




        //GET api/download/12345abc
        //[HttpGet("{id}")]
        //public async Task<IActionResult> retrieveAttachmment(string id)
        //{

        //  String readApi = "http://appd13was:9044/Rest/v1/docmgmt/readAttachmentVersionContent?attachVersionID=669891871";


        //if (stream == null)
        //    return NotFound(); // returns a NotFoundResult with Status404NotFound response.

        //return File(stream, "application/octet-stream"); // returns a FileStreamResult
        //}

        public string getMimeExtension(string mimeType)
        {
            string result;
            RegistryKey key;
            object value;

            key = Registry.ClassesRoot.OpenSubKey(@"MIME\Database\Content Type\" + mimeType, false);
            value = key != null ? key.GetValue("Extension", null) : null;
            result = value != null ? value.ToString() : new string(WopiOptions.Value.Word2010Ext);

            return result;
        }

        /*public void convertDocToDocxGlue(string path, Boolean useTemp = false)
        {


            if (path.ToLower().EndsWith(WopiOptions.Value.Word2010Ext))
            {

                var sourceFile = new FileInfo(path);
                string newFileName = sourceFile.FullName.Replace(WopiOptions.Value.Word2010Ext, WopiOptions.Value.WordExt);

                if (!System.IO.File.Exists(newFileName))
                {
                    System.IO.File.Delete(newFileName);

                    //ComponentInfo.SetLicense("FREE-LIMITED-KEY");

                    // In order to convert Word to PDF, we just need to:
                    // 1. Load DOC or DOCX file into DocumentModel object.
                    // 2. Save DocumentModel object to PDF file.
                    //DocumentModel document = DocumentModel.Load(sourceFile.FullName);
                    //document.Save(newFileName);
                    using (Doc doc = new Doc(sourceFile.FullName))
                    doc.SaveAs(newFileName);


                    try
                    {
                                var current = DateTime.Now;
                                System.IO.File.SetCreationTime(newFileName, current);
                                System.IO.File.SetLastWriteTime(newFileName, current);
                                System.IO.File.SetLastAccessTime(newFileName, current);
                                
                        Console.WriteLine(System.IO.File.GetCreationTime(newFileName));
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }

                }
                else
                {
                    //the file exists
                    if (!useTemp)
                    {
                        System.IO.File.Delete(newFileName);
                        //var document = word.Documents.Open(sourceFile.FullName);

                        //var project = document.VBProject;
                        //var module = project.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
                        // module.CodeModule.AddFromString("CMSUpdateFields");
                        //word.Run("CMSUpdateFields");
                        //word.ActiveDocument.ActiveWindow.View.ShowFieldCodes = false;
                        //word.ActiveDocument.ActiveWindow.View.ShowFieldCodes = false;
                        //word.ActiveDocument.StoryRanges.Fields.Update();
                        //word.ActiveDocument.StoryRanges.Fields.Update();


                        //document.SaveAs2(newFileName, WdSaveFormat.wdFormatXMLDocument,
                        //             CompatibilityMode: WdCompatibilityMode.wdWord2010);

                        //word.ActiveDocument.Close();
                        //word.Quit();

                        //ComponentInfo.SetLicense("FREE-LIMITED-KEY");

                        // In order to convert Word to PDF, we just need to:
                        // 1. Load DOC or DOCX file into DocumentModel object.
                        // 2. Save DocumentModel object to PDF file.
                        //DocumentModel document = DocumentModel.Load(sourceFile.FullName);
                        //document.Save(newFileName);
                        using (Doc doc = new Doc(sourceFile.FullName))
                            doc.SaveAs(newFileName);

                        try
                        {
                                var current = DateTime.Now;
                                System.IO.File.SetCreationTime(newFileName, current);
                                System.IO.File.SetLastWriteTime(newFileName, current);
                                System.IO.File.SetLastAccessTime(newFileName, current);
                                
                            Console.WriteLine(System.IO.File.GetCreationTime(newFileName));
                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }

                    }
                    else
                    {
                        //leave it alone
                    }

                }

                var tempfile = path.Replace(WopiOptions.Value.Word2010Ext, WopiOptions.Value.RetreivedSuffix) + WopiOptions.Value.Word2010Ext;
                if (System.IO.File.Exists(tempfile)) System.IO.File.Delete(tempfile);
                System.IO.File.Move(path, tempfile);

            }
            else
            {
                //TODO
                ;
            }
        }*/


        public void convertDocToDocxOpenXML(string fileName, Boolean useTemp = false, string extension = ".docx")
        {
            bool fileChanged = false;

            using (WordprocessingDocument document =
                WordprocessingDocument.Open(fileName, true))
            {
                // Access the main document part.
                var docPart = document.MainDocumentPart;

                // Look for the vbaProject part. If it is there, delete it.
                var vbaPart = docPart.VbaProjectPart;
                if (vbaPart != null)
                {
                    // Delete the vbaProject part and then save the document.
                    docPart.DeletePart(vbaPart);
                    docPart.Document.Save();

                    // Change the document type to
                    // not macro-enabled.
                    document.ChangeDocumentType(
                        WordprocessingDocumentType.Document);

                    // Track that the document has been changed.
                    fileChanged = true;
                }
            }

            // If anything goes wrong in this file handling,
            // the code will raise an exception back to the caller.
            if (fileChanged)
            {
                // Create the new .docx filename.
                var newFileName = Path.ChangeExtension(fileName, extension);

                // If it already exists, it will be deleted!
                if (System.IO.File.Exists(newFileName))
                {
                    System.IO.File.Delete(newFileName);
                }

                // Rename the file.
                System.IO.File.Move(fileName, newFileName);
            }
        }


        public void convertDocToDocxAspose(string path, Boolean useTemp = false, string userRole = null)
        {


            if (path.ToLower().EndsWith(WopiOptions.Value.Word2010Ext))
            {

                var sourceFile = new FileInfo(path);
                bool processedByWWA = false;
                string newFileName = sourceFile.FullName.Replace(WopiOptions.Value.Word2010Ext, WopiOptions.Value.WordExt);
                Dictionary<string, string> fieldMap = new Dictionary<string, string>();

                Aspose.Words.Loading.LoadOptions loadOptions = new Aspose.Words.Loading.LoadOptions
                {
                    LoadFormat = LoadFormat.WordML
                };

                OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
                {
                    //Compliance = OoxmlCompliance.Iso29500_2008_Strict,
                    SaveFormat = SaveFormat.Docx
                };

                if (!System.IO.File.Exists(newFileName))
                {
                    System.IO.File.Delete(newFileName);

                    //ComponentInfo.SetLicense("FREE-LIMITED-KEY");

                    // In order to convert Word to PDF, we just need to:
                    // 1. Load DOC or DOCX file into DocumentModel object.
                    // 2. Save DocumentModel object to PDF file.
                    //DocumentModel document = DocumentModel.Load(sourceFile.FullName);
                    //document.Save(newFileName);

                    var disableRemoteResourcesOptions = new Aspose.Words.Loading.LoadOptions
                    {
                        ResourceLoadingCallback = new DisableRemoteResourcesHandler(),
                        LoadFormat = LoadFormat.WordML
                    };

                    Aspose.Words.Document document = null;
                    if (!(null == WopiOptions.Value.ConversionEngineDisableExternalResources || WopiOptions.Value.ConversionEngineDisableExternalResources.Contains(FALSE)))
                        document = new Aspose.Words.Document(sourceFile.FullName, disableRemoteResourcesOptions);
                    else
                        document = new Aspose.Words.Document(sourceFile.FullName, loadOptions);
                    //document = new Aspose.Words.Document(sourceFile.FullName, loadOptions);

                    //Aspose.Words.Document document = new Aspose.Words.Document(sourceFile.FullName, disableRemoteResourcesOptions);
                    //document.LoadFromFile(sourceFile.FullName, FileFormat.WordXml);
                    //                    document.LoadFromFile(sourceFile.FullName);
                    document.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2010);

                    foreach (Aspose.Words.Properties.DocumentProperty docProperty in document.CustomDocumentProperties)
                    {
                        Console.Out.WriteLine(docProperty.Name);
                        Console.Out.WriteLine($"\tType:\t{docProperty.Type}");

                        // Some properties may store multiple values.
                        if (docProperty.Value is Array)
                        {
                            string stringArray = "[";
                            int i = 0;
                            foreach (object value in docProperty.Value as Array)
                            {

                                if (null != value)
                                {
                                    Console.Out.WriteLine($"\tValue:\t\"" + value.ToString() + "\"");
                                    if (i == 0)
                                    {
                                        stringArray += value.ToString();
                                    }
                                    else
                                    {
                                        stringArray += ", ";
                                        stringArray += value.ToString();
                                    }
                                }
                                i = i + 1;
                            }
                            stringArray += "]";
                            fieldMap.Add(docProperty.Name.ToUpper(), stringArray);
                        }
                        else
                        {
                            Console.Out.WriteLine($"\tValue:\t\"{docProperty.Value}\"");
                            fieldMap.Add(docProperty.Name.ToString().ToUpper(), docProperty.Value.ToString());
                        }

                        if (null != docProperty.Name && docProperty.Name.ToString().ToUpper() == WopiOptions.Value.ProcessedFlag)
                        {
                            if (null != docProperty.Value && docProperty.Value.ToString() == WopiOptions.Value.ApplicationName)
                            {
                                processedByWWA = true;
                            }
                        }
                    }

                    //var withMacroFile = newFileName.Replace(WopiOptions.Value.WordExt, WopiOptions.Value.WithMacroSuffix) + WopiOptions.Value.WordExt;
                    //document.Save(withMacroFile, saveOptions);
                    var macroFile = newFileName.Replace(WopiOptions.Value.WordExt, WopiOptions.Value.WithMacroSuffix) + WopiOptions.Value.WordMacroExt;
                    document.Save(macroFile, SaveFormat.Docm);

                    document.RemoveMacros();
                    //document.Save(newFileName, SaveFormat.Doc);
                    var preMacroFile = newFileName.Replace(WopiOptions.Value.WordExt, WopiOptions.Value.PreMacroSuffix) + WopiOptions.Value.WordExt;
                    document.Save(preMacroFile, saveOptions);
                    document.Save(newFileName, saveOptions);

                    if (WopiOptions.Value.RunMacro.ToLower() == TRUE)
                    {
                        var callRunMacro = "runMacro";

                        if (!String.IsNullOrEmpty(WopiOptions.Value.RunMacroVersion))
                        {
                            callRunMacro += WopiOptions.Value.RunMacroVersion;
                        }

                        MethodInfo runMacroMethod = this.GetType().GetMethod(callRunMacro);
                        object result = null;

                        try
                        {
                            runMacroMethod.Invoke(this, new object[] { newFileName, fieldMap, processedByWWA, true, userRole });
                        }
                        catch (Exception exc)
                        {
                            Console.Out.WriteLine("******** Error running macro method " + callRunMacro);
                        }

                    }

                    var preEditFile = newFileName.Replace(WopiOptions.Value.WordExt, WopiOptions.Value.PostMacroSuffix) + WopiOptions.Value.WordExt;
                    if (System.IO.File.Exists(preEditFile)) System.IO.File.Delete(preEditFile);
                    System.IO.File.Copy(newFileName, preEditFile);


                    try
                    {
                        var current = DateTime.Now;
                        System.IO.File.SetCreationTime(newFileName, current);
                        System.IO.File.SetLastWriteTime(newFileName, current);
                        System.IO.File.SetLastAccessTime(newFileName, current);

                        Console.WriteLine(System.IO.File.GetCreationTime(newFileName));
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }

                }
                else
                {
                    //the file exists
                    if (!useTemp)
                    {
                        System.IO.File.Delete(newFileName);
                        //var document = word.Documents.Open(sourceFile.FullName);

                        //var project = document.VBProject;
                        //var module = project.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
                        // module.CodeModule.AddFromString("CMSUpdateFields");
                        //word.Run("CMSUpdateFields");
                        //word.ActiveDocument.ActiveWindow.View.ShowFieldCodes = false;
                        //word.ActiveDocument.ActiveWindow.View.ShowFieldCodes = false;
                        //word.ActiveDocument.StoryRanges.Fields.Update();
                        //word.ActiveDocument.StoryRanges.Fields.Update();


                        //document.SaveAs2(newFileName, WdSaveFormat.wdFormatXMLDocument,
                        //             CompatibilityMode: WdCompatibilityMode.wdWord2010);

                        //word.ActiveDocument.Close();
                        //word.Quit();

                        //ComponentInfo.SetLicense("FREE-LIMITED-KEY");

                        // In order to convert Word to PDF, we just need to:
                        // 1. Load DOC or DOCX file into DocumentModel object.
                        // 2. Save DocumentModel object to PDF file.
                        //DocumentModel document = DocumentModel.Load(sourceFile.FullName);
                        //document.Save(newFileName);
                        //Document document = new Document();
                        //document.LoadFromFile(sourceFile.FullName, FileFormat.WordXml);
                        //document.SaveToFile(newFileName, FileFormat.Docx2013);
                        //Aspose.Words.Document document = new Aspose.Words.Document(sourceFile.FullName);
                        //document.LoadFromFile(sourceFile.FullName, FileFormat.WordXml);
                        //                    document.LoadFromFile(sourceFile.FullName);
                        //document.Save(newFileName, SaveFormat.Doc);
                        //document.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2010);

                        var disableRemoteResourcesOptions = new Aspose.Words.Loading.LoadOptions
                        {
                            ResourceLoadingCallback = new DisableRemoteResourcesHandler(),
                            LoadFormat = LoadFormat.WordML
                        };

                        Aspose.Words.Document document = null;
                        if (!(null == WopiOptions.Value.ConversionEngineDisableExternalResources || WopiOptions.Value.ConversionEngineDisableExternalResources.Contains(FALSE)))
                            document = new Aspose.Words.Document(sourceFile.FullName, disableRemoteResourcesOptions);
                        else
                            document = new Aspose.Words.Document(sourceFile.FullName, loadOptions);
                        //document = new Aspose.Words.Document(sourceFile.FullName, loadOptions);

                        //Aspose.Words.Document document = new Aspose.Words.Document(sourceFile.FullName, disableRemoteResourcesOptions);
                        //document.LoadFromFile(sourceFile.FullName, FileFormat.WordXml);
                        //                    document.LoadFromFile(sourceFile.FullName);
                        document.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2010);

                        foreach (Aspose.Words.Properties.DocumentProperty docProperty in document.CustomDocumentProperties)
                        {
                            Console.Out.WriteLine(docProperty.Name);
                            Console.Out.WriteLine($"\tType:\t{docProperty.Type}");

                            // Some properties may store multiple values.
                            if (docProperty.Value is Array)
                            {
                                string stringArray = "[";
                                int i = 0;
                                foreach (object value in docProperty.Value as Array)
                                {

                                    if (null != value)
                                    {
                                        Console.Out.WriteLine($"\tValue:\t\"" + value.ToString() + "\"");
                                        if (i == 0)
                                        {
                                            stringArray += value.ToString();
                                        }
                                        else
                                        {
                                            stringArray += ", ";
                                            stringArray += value.ToString();
                                        }
                                    }
                                    i = i + 1;
                                }
                                stringArray += "]";
                                fieldMap.Add(docProperty.Name.ToUpper(), stringArray);
                            }
                            else
                            {
                                Console.Out.WriteLine($"\tValue:\t\"{docProperty.Value}\"");
                                fieldMap.Add(docProperty.Name.ToUpper(), docProperty.Value.ToString());
                            }

                            if (null != docProperty.Name && docProperty.Name.ToUpper() == WopiOptions.Value.ProcessedFlag)
                            {
                                if (null != docProperty.Value && docProperty.Value.ToString() == WopiOptions.Value.ApplicationName)
                                {
                                    processedByWWA = true;
                                }
                            }
                        }


                        var macroFile = newFileName.Replace(WopiOptions.Value.WordExt, WopiOptions.Value.WithMacroSuffix) + WopiOptions.Value.WordMacroExt;
                        document.Save(macroFile, SaveFormat.Docm);

                        document.RemoveMacros();

                        var preMacroFile = newFileName.Replace(WopiOptions.Value.WordExt, WopiOptions.Value.PreMacroSuffix) + WopiOptions.Value.WordExt;
                        document.Save(preMacroFile, saveOptions);
                        document.Save(newFileName, saveOptions);

                        /*if (WopiOptions.Value.RunMacro.ToLower() == TRUE)
                        {
                            runMacroX(newFileName, fieldMap, processedByWWA, true, userRole);
                        }*/

                        if (WopiOptions.Value.RunMacro.ToLower() == TRUE)
                        {
                            var callRunMacro = "runMacro";

                            if (!String.IsNullOrEmpty(WopiOptions.Value.RunMacroVersion))
                            {
                                callRunMacro += WopiOptions.Value.RunMacroVersion;
                            }

                            MethodInfo runMacroMethod = this.GetType().GetMethod(callRunMacro);
                            object result = null;

                            try
                            {
                                runMacroMethod.Invoke(this, new object[] { newFileName, fieldMap, processedByWWA, true, userRole });
                            }
                            catch (Exception exc)
                            {
                                Console.Out.WriteLine("******** Error running macro method " + callRunMacro);
                            }

                        }

                        var preEditFile = newFileName.Replace(WopiOptions.Value.WordExt, WopiOptions.Value.PostMacroSuffix) + WopiOptions.Value.WordExt;
                        if (System.IO.File.Exists(preEditFile)) System.IO.File.Delete(preEditFile);
                        System.IO.File.Copy(newFileName, preEditFile);

                        try
                        {
                            var current = DateTime.Now;
                            System.IO.File.SetCreationTime(newFileName, current);
                            System.IO.File.SetLastWriteTime(newFileName, current);
                            System.IO.File.SetLastAccessTime(newFileName, current);

                            Console.WriteLine(System.IO.File.GetCreationTime(newFileName));
                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }

                    }
                    else
                    {
                        //leave it alone
                    }

                }

                var tempfile = path.Replace(WopiOptions.Value.Word2010Ext, WopiOptions.Value.RetreivedSuffix) + WopiOptions.Value.Word2010Ext;
                if (System.IO.File.Exists(tempfile)) System.IO.File.Delete(tempfile);
                System.IO.File.Move(path, tempfile);

            }
            else
            {
                //TODO
                ;
            }



        }


        public void runMacroX3(string newFileName, Dictionary<string, string> fieldMap, bool processedByWWA = false, bool normalize = false, string userRole = null)
        {
            Regex fileVer = new Regex(@"\s*(?<userID>)_(?<attachmentID>)_(?<versionNo>)\..*");
            var result = fileVer.Matches(newFileName);
            string userID = null;
            string attachmentID = null;
            string versionNo = null;

            /*if (Int32.TryParse(versionNo, out int v))
            {
                if (v > 1) return;
            }*/
            //if (processedByWWA) return;

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(newFileName, true))
            {

                if (null != wordDocument)
                {

                    try
                    {
                        /*for (var field in wordDocument.MainDocumentPart.RootElement.Descendants<FieldCode>())
                        {

                        }*/

                        /*List<Picture> pictures = new List<Picture>(wordDocument.MainDocumentPart.RootElement.Descendants<Picture>());

                        foreach (Picture p in pictures)
                        {
                            p.Remove();
                        }*/

                        //headerPart.DeleteParts(imagePartList);

                        // COMMENT OUT AFTER ACQUIRING ASPOSE LICENSE
                        /*
                        foreach (Paragraph p in wordDocument.MainDocumentPart.Document.Body.Descendants<Paragraph>().Where<Paragraph>(p => p.InnerText.Contains(WopiOptions.Value.ConversionEngine)))
                        {
                            // Do something with the Paragraphs.
                            p.Remove();
                        }

                        if (wordDocument.MainDocumentPart.HeaderParts.Count() > 0)
                        {
                            foreach (HeaderPart headerPart in wordDocument.MainDocumentPart.HeaderParts)
                            {
                                foreach (Paragraph p in headerPart.Header.Descendants<Paragraph>().Where<Paragraph>(p => p.InnerText.Contains(WopiOptions.Value.ConversionEngine)))
                                {
                                    p.Remove();
                                }
                            }
                        }

                        if (wordDocument.MainDocumentPart.FooterParts.Count() > 0)
                        {
                            foreach (FooterPart footerPart in wordDocument.MainDocumentPart.FooterParts)
                            {
                                foreach (Paragraph p in footerPart.Footer.Descendants<Paragraph>().Where<Paragraph>(p => p.InnerText.Contains(WopiOptions.Value.ConversionEngine)))
                                {
                                    p.Remove();
                                }
                            }
                        }
                        */
                        // COMMENT OUT AFTER ACQUIRING ASPOSE LICENSE

                        if (WopiOptions?.Value?.DocumentProtection == TRUE)
                        {
                            if (getIDPList().Contains("all") || getIDPList().Contains(userRole) || userRole == null)
                            {
                                ArrayList docProtClassList = new ArrayList();
                                ArrayList docProtTypeList = new ArrayList();

                                if (null != WopiOptions.Value.DocumentProtectionClass && WopiOptions.Value.DocumentProtectionClass.Length > 0)
                                    docProtClassList.AddRange(WopiOptions.Value.DocumentProtectionClass);

                                if (null != WopiOptions.Value.DocumentProtectionType && WopiOptions.Value.DocumentProtectionType.Length > 0)
                                    docProtTypeList.AddRange(WopiOptions.Value.DocumentProtectionType);


                                var dsp = wordDocument.MainDocumentPart.DocumentSettingsPart;
                                foreach (DocumentProtection dp in wordDocument.MainDocumentPart.DocumentSettingsPart.Settings.ChildElements.OfType<DocumentProtection>())
                                {
                                    if (!(String.IsNullOrEmpty(WopiOptions.Value.DocumentProtectionFlag)))
                                    {
                                        if (null != dp.Enforcement)
                                        {

                                            if (dp.Enforcement == new OnOffValue(true))
                                            {
                                                //setCustomProperty(wordDocument, WopiOptions.Value.DocumentProtectionFlag, WopiOptions.Value.DocumentProtectionClass + "," + WopiOptions.Value.DocumentProtectionType, CustomPropertyTypes.Text);
                                                string docProtString = null;
                                                string docProtClass = null;
                                                string docProtType = null;
                                                if (docProtClassList.Contains("edit") && null != dp.Edit)
                                                {
                                                    docProtClass = "edit";

                                                    /*//
                                                    // Summary:
                                                    //     No Editing Restrictions.
                                                    //     When the item is serialized out as xml, its value is "none".
                                                    None = 0,
                                                    //
                                                    // Summary:
                                                    //     Allow No Editing.
                                                    //     When the item is serialized out as xml, its value is "readOnly".
                                                    ReadOnly = 1,
                                                    //
                                                    // Summary:
                                                    //     Allow Editing of Comments.
                                                    //     When the item is serialized out as xml, its value is "comments".
                                                    Comments = 2,
                                                    //
                                                    // Summary:
                                                    //     Allow Editing With Revision Tracking.
                                                    //     When the item is serialized out as xml, its value is "trackedChanges".
                                                    TrackedChanges = 3,
                                                    //
                                                    // Summary:
                                                    //     Allow Editing of Form Fields.
                                                    //     When the item is serialized out as xml, its value is "forms".
                                                    Forms = 4 */
                                                    if (dp.Edit == DocumentProtectionValues.ReadOnly) docProtType = "1";
                                                    if (dp.Edit == DocumentProtectionValues.Comments) docProtType = "2";
                                                    if (dp.Edit == DocumentProtectionValues.TrackedChanges) docProtType = "3";
                                                    if (dp.Edit == DocumentProtectionValues.Forms) docProtType = "4";
                                                    if (dp.Edit == DocumentProtectionValues.None) docProtType = "0";

                                                    var dpnew = new DocumentProtection()
                                                    {
                                                        Edit = dp.Edit,
                                                        Enforcement = new OnOffValue(false),
                                                        Formatting = dp.Formatting
                                                        //CryptographicProviderType = CryptProviderValues.RsaFull,
                                                        //CryptographicAlgorithmClass = CryptAlgorithmClassValues.Hash,
                                                        //CryptographicAlgorithmType = CryptAlgorithmValues.TypeAny,
                                                        //CryptographicAlgorithmSid = 4,
                                                        //CryptographicSpinCount = 100000U,
                                                        //Hash = "2krUoz1qWd0WBeXqVrOq81l8xpk=",
                                                        //Salt = "9kIgmDDYtt2r5U2idCOwMA=="
                                                    };


                                                    if (!(null == dsp || null == dsp.Settings))
                                                    {
                                                        dsp.Settings.ReplaceChild(dpnew, dp);
                                                    }
                                                }

                                                // handles other types of protection here


                                                setCustomProperty(wordDocument, WopiOptions.Value.DocumentProtectionFlag, docProtClass + "," + docProtType, CustomPropertyTypes.Text);

                                                var docProtKey = Path.GetFileNameWithoutExtension(newFileName);
                                                if (null != docProtKey && docProtection.ContainsKey(docProtKey))
                                                {
                                                    docProtection[docProtKey] = TRUE;
                                                }
                                                else
                                                {
                                                    docProtection.Add(docProtKey, TRUE);
                                                }


                                            }
                                        }
                                    }
                                    //dp.Remove();




                                    /*
                                    < w:documentProtection w:edit = "forms"
                                      w: formatting = "1"
                                      w: enforcement = "1" />
                                    */

                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new WordWebException("This document contains improper document protection", ex);
                    }


                    try
                    {
                        //const string FieldDelimeter = @" MERGEFIELD ";
                        string FieldDelimeter = @" DOCPROPERTY ";
                        List<string> listeChamps = new List<string>();


                        if (normalize)
                        {
                            normalizeMarkup(wordDocument);
                            normalizeFieldCodesRuns(wordDocument);
                        }

                    }
                    catch (Exception ex)
                    {
                        Console.Out.WriteLine("Unable to normalize the document markup...");
                    }

                    //ANCHOR
                    //try
                    //{
                    Run prevRun = null;
                    Run prevBegin = null;
                    Run prevEnd = null;

                    string[] nonEditable = WopiOptions.Value.BookmarksNonEditable;

                    /* new string[] {
                      "LetterDate",
                     "NameAndAddressLn1",
                     "NameAndAddressLn2",
                     "NameAndAddressLn3",
                     "NameAndAddressLn4",
                     "NameAndAddressLn5",
                     "NameAndAddressLn6",
                     "NameAndAddressLn7",
                     "(P)",
                     "\"CC\"",
                     " CC ",
                     "COPYIND",
                     "CCTOKENLIST",
                     "IPTOKENLIST",
                     "PRIMARYRECIPCONTACTNAME" }; */



                    ArrayList nonEditableArray = new ArrayList();
                    for (var k = 0; k < nonEditable.Length; k++)
                    {
                        nonEditableArray.Add(nonEditable[k].ToUpper());
                    }

                    //nonEditableArray.AddRange(nonEditable);



                    int j = 0;
                    var fieldCodes = wordDocument.MainDocumentPart.RootElement.Descendants<FieldCode>();
                    //FieldCode[] fieldCodesArray = fieldCodes.ToArray();
                    //Dictionary<string, DocField> docFields = new Dictionary<string, DocField>();
                    //List<(string, DocField)> docFields = new List<(string, DocField)>();
                    //List<(string, DocField)> malFormedFields = new List<(string, DocField)>();



                    ///////////////////////////////////////////////////////////////////////////////////////////////////////////
                    /// REPAIR PASS
                    ///////////////////////////////////////////////////////////////////////////////////////////////////////////
                    foreach (var field in fieldCodes)
                    {
                        j++;
                        Console.Out.WriteLine("***Field " + j.ToString() + ">>>" + "Field Type:" + field.ToString() + "<<<Field Text:>>>" + field.Text.ToString() + "<<<");
                        if (null != field.InnerXml) Console.Out.WriteLine("***** internal XML:" + field.InnerXml + " *****");
                        bool isFormField = false;

                        var fieldId = field?.Text?.ToString().Trim() ?? "";
                        if (null != fieldId)
                        {
                            if (fieldId.Contains("FORMTEXT"))
                            {
                                isFormField = true;
                            }
                            Console.Out.WriteLine("fieldId:" + fieldId.ToString());
                        }

                        Run xxxfield = null;
                        Run rBegin = null;
                        Run rSep = null;
                        Run rText = null;
                        Run rEnd = null;
                        Text t = null;
                        RunProperties rProp = null;
                        Paragraph rPara = null;
                        Run rParent = null;
                        Run pivotRun = null;
                        string aFieldName = null;
                        var fillerRuns1 = new List<Run>();
                        var fillerRuns2 = new List<Run>();
                        var fillerRuns3 = new List<Run>();
                        var fillerRuns4 = new List<Run>();
                        var fillerRuns = new ArrayList();
                        var k = 0;
                        bool rBeginFound = false;
                        bool rFieldFound = false;
                        bool rSepFound = false;
                        bool rTextFound = false;
                        bool rEndFound = false;
                        var malFormattedFields = new List<Run>();

                            xxxfield = (Run)field?.Parent;
                            if (null != xxxfield)
                            {
                                rFieldFound = true;
                                rPara = (Paragraph)xxxfield.Parent;
                            }

                            k = 0;
                            if (null != xxxfield) pivotRun = xxxfield;
                            while (null != pivotRun && (rBegin == null && k < 10))
                            {
                                var aBegin = pivotRun.PreviousSibling<Run>();
                                if (null != aBegin)
                                {

                                    if (aBegin.GetType() == typeof(Run) &&
                                        aBegin.Elements<FieldChar>()?.FirstOrDefault(fc =>
                                            fc.FieldCharType == FieldCharValues.Begin) != null)
                                    {
                                        rBegin = aBegin;
                                        rBeginFound = true;
                                    }
                                    else
                                    {
                                        if ((aBegin.GetType() == typeof(Run)) &&
                                            (aBegin.Elements<FieldChar>()?.FirstOrDefault().FieldCharType == FieldCharValues.End
                                            || aBegin.Elements<FieldChar>()?.FirstOrDefault().FieldCharType == FieldCharValues.Separate)
                                        )
                                        {
                                            break;
                                        }
                                        else
                                        {


                                            //fillerRuns1[k] = new Run();
                                            fillerRuns1.Add(aBegin);
                                        }
                                    }                                    
                                }
                                pivotRun = aBegin;
                                k++;
                            }

                            //if (null != xxxfield) rSep = xxxfield.NextSibling<Run>();
                            k = 0;
                            if (null != xxxfield) pivotRun = xxxfield;
                            while (null != pivotRun && (rSep == null && k < 10))
                            {
                                var aSep = pivotRun.NextSibling<Run>();
                                if (null != aSep)
                                {

                                    if (aSep.GetType() == typeof(Run) &&
                                        aSep.Elements<FieldChar>().FirstOrDefault(fc =>
                                            fc.FieldCharType == FieldCharValues.Separate) != null)
                                    {
                                        rSep = aSep;
                                        rSepFound = true;
                                    }
                                    else
                                    {
                                        if ((aSep.GetType() == typeof(Run)) &&
                                        (aSep.Elements<FieldChar>()?.FirstOrDefault().FieldCharType == FieldCharValues.End
                                        || aSep.Elements<FieldChar>()?.FirstOrDefault().FieldCharType == FieldCharValues.Begin)
                                        )
                                        {
                                            break;
                                        }
                                        else
                                        { //fillerRuns2[k] = new Run();
                                            fillerRuns2.Add(aSep);
                                        }
                                    }
                                }
                                pivotRun = aSep;
                                k++;
                            }

                            //if (null != rSep) rText = rSep.NextSibling<Run>();
                            k = 0;
                            if (null != rSep) pivotRun = rSep;
                            while (null != pivotRun && (rText == null && k < 10))
                            {
                                var aText = pivotRun.NextSibling<Run>();
                                if (null != aText)
                                {

                                    if (aText.GetType() == typeof(Run) &&
                                        aText.GetFirstChild<Text>() != null)
                                    {
                                        rText = aText;
                                        rTextFound = true;
                                    }
                                    else
                                    {
                                        if ((aText.GetType() == typeof(Run)) &&
                                           (aText.Elements<FieldChar>()?.FirstOrDefault().FieldCharType == FieldCharValues.End
                                           || aText.Elements<FieldChar>()?.FirstOrDefault().FieldCharType == FieldCharValues.Begin
                                           || aText.Elements<FieldChar>()?.FirstOrDefault().FieldCharType == FieldCharValues.Separate)
                                           )
                                        {
                                            break;
                                        }
                                        else
                                        {    //fillerRuns3[k] = new Run();
                                            fillerRuns3.Add(aText);
                                        }
                                    }
                                }
                                pivotRun = aText;
                                k++;
                            }


                            //if (null != rText) rEnd = rText.NextSibling<Run>();
                            k = 0;
                            if (null != rText) pivotRun = rText;
                            while (null != pivotRun && (rEnd == null && k < 10))
                            {
                                var aEnd = pivotRun.NextSibling<Run>();
                                if (null != aEnd)
                                {

                                    if (aEnd.GetType() == typeof(Run) &&
                                        aEnd.Elements<FieldChar>().FirstOrDefault(fc =>
                                            fc.FieldCharType == FieldCharValues.End) != null)
                                    {
                                        rEnd = aEnd;
                                        rEndFound = true;
                                    }
                                    else
                                    {
                                        if ((aEnd.GetType() == typeof(Run)) &&
                                           (aEnd.Elements<FieldChar>()?.FirstOrDefault().FieldCharType == FieldCharValues.Begin
                                           || aEnd.Elements<FieldChar>()?.FirstOrDefault().FieldCharType == FieldCharValues.Separate)
                                           )
                                        {
                                            break;
                                        }
                                        else
                                        {//fillerRuns4[k] = new Run();
                                            fillerRuns4.Add(aEnd);
                                        }
                                    }
                                }
                                pivotRun = aEnd;
                                k++;
                            }

                            k = 0;
                            Run innerRun = null;
                            int l = 0;
                            if (!rBeginFound && null == rBegin)
                            {
                                if (fillerRuns1.Count > 0) { fillerRuns1.Clear(); }
                                if (null != rPara)
                                {
                                    Run first = (Run)rPara.FirstChild;
                                    if (null != first)
                                    {
                                        l = 0;
                                        innerRun = first;
                                        while (null != innerRun && (rBegin == null && l < 10))
                                        {
                                            if (innerRun == xxxfield)
                                            {
                                                Run rField = new Run();
                                                rField = (Run)xxxfield.Clone();
                                                innerRun.Append(rField);
                                                innerRun = new Run(new FieldChar() { FieldCharType = FieldCharValues.Begin });
                                                rBegin = innerRun;
                                                rBeginFound = true;
                                            }
                                            else
                                            {
                                                var aField = innerRun.NextSibling<Run>();
                                                if (null != aField)
                                                {

                                                    if (aField.GetType() == typeof(Run) &&
                                                        aField == xxxfield)
                                                    {
                                                        rBegin = new Run(new FieldChar() { FieldCharType = FieldCharValues.Begin });
                                                        innerRun.Append(rBegin);
                                                        rBeginFound = true;
                                                    }
                                                }
                                                innerRun = aField;
                                            }
                                            l++;
                                        }
                                    }
                                }
                            }

                            k = 0;
                            innerRun = null;
                            if (!rSepFound && null == rSep)
                            {
                                if (null != xxxfield) innerRun = xxxfield;
                                if (fillerRuns2.Count > 0) { fillerRuns2.Clear(); }
                                if (null != innerRun)
                                {
                                    rSep = new Run(new FieldChar() { FieldCharType = FieldCharValues.Separate });
                                    innerRun.Append(rSep);
                                    rSepFound = true;
                                }

                            }

                            k = 0;
                            innerRun = null;
                            if (!rTextFound && null == rText)
                            {
                                if (null != rEnd) innerRun = rEnd;
                                if (fillerRuns3.Count > 0) { fillerRuns3.Clear(); }
                                if (null != rEnd)
                                {
                                    rText = new Run();
                                    rText.AppendChild<Text>(new Text(" "));
                                    innerRun.InsertBeforeSelf<Run>(rText);
                                    rTextFound = true;
                                }
                            }

                            k = 0;
                            innerRun = null;
                            if (!rEndFound && null == rEnd)
                            {
                                if (null != rText) innerRun = rText;
                                if (fillerRuns4.Count > 0) { fillerRuns4.Clear(); }
                                if (null != innerRun)
                                {
                                    rEnd = new Run(new FieldChar() { FieldCharType = FieldCharValues.End });
                                    innerRun.Append(rEnd);
                                    rEndFound = true;
                                }
                            }

                            if (null != rText) t = rText.GetFirstChild<Text>();
                            if (null != rText) rProp = (RunProperties)rText.RunProperties;
                            if (null == t)
                            {
                                rText.AppendChild(new Text(" "));
                            }

                            l = 0;
                            Run rMalFormatted = rText.GetFirstChild<Run>();
                            Run nextError = rMalFormatted;
                            while (null != nextError && l < 10)
                            {
                                nextError = rMalFormatted.NextSibling<Run>();
                                rMalFormatted.Remove();
                                l++;
                            }
                        


                    }

                    ///////////////////////////////////////////////////////////////////////////////////////////////////////////
                    j = 0;
                    
                    fieldCodes = wordDocument.MainDocumentPart.RootElement.Descendants<FieldCode>();
                    FieldCode[] fieldCodesArray = fieldCodes.ToArray();
                    //Dictionary<string, DocField> docFields = new Dictionary<string, DocField>();
                    List<(string, DocField)> docFields = new List<(string, DocField)>();
                    List<(string, DocField)> malFormedFields = new List<(string, DocField)>();
                    
                    ///////////////////////////////////////////////////////////////////////////////////////////////////////////
                    foreach (var field in fieldCodes)
                    {
                        j++;
                        Console.Out.WriteLine("***Field " + j.ToString() + ">>>" + "Field Type:" + field.ToString() + "<<<Field Text:>>>" + field.Text.ToString() + "<<<");
                        if (null != field.InnerXml) Console.Out.WriteLine("***** internal XML:" + field.InnerXml + " *****");
                        bool isFormField = false;

                        var fieldId = field?.Text?.ToString().Trim() ?? "";
                        if (null != fieldId)
                        {
                            if (fieldId.Contains("FORMTEXT"))
                            {
                                isFormField = true;
                            }
                            Console.Out.WriteLine("fieldId:" + fieldId.ToString());
                        }

                        Run xxxfield = null;
                        Run rBegin = null;
                        Run rSep = null;
                        Run rText = null;
                        Run rEnd = null;
                        Text t = null;
                        RunProperties rProp = null;
                        Paragraph rPara = null;
                        Run rParent = null;
                        Run pivotRun = null;
                        string aFieldName = null;
                        var fillerRuns1 = new List<Run>();
                        var fillerRuns2 = new List<Run>();
                        var fillerRuns3 = new List<Run>();
                        var fillerRuns4 = new List<Run>();
                        var fillerRuns = new ArrayList();
                        var k = 0;
                        bool rBeginFound = false;
                        bool rFieldFound = false;
                        bool rSepFound   = false;
                        bool rTextFound  = false;
                        bool rEndFound   = false;
                        var malFormattedFields = new List<Run>();

                        if (isFormField)
                        {
                            xxxfield = (Run)field?.Parent;                            
                            if (null != xxxfield) {
                                rFieldFound = true;
                                rPara = (Paragraph)xxxfield.Parent;                                
                            }

                            k = 0;
                            pivotRun = xxxfield;
                            while (null != pivotRun && (rBegin == null && k < 10))
                            {
                                var aBegin = pivotRun.PreviousSibling<Run>();
                                if (null != aBegin)
                                {

                                    if (aBegin.GetType() == typeof(Run) &&
                                        aBegin.Elements<FieldChar>().FirstOrDefault(fc =>
                                            fc.FieldCharType == FieldCharValues.Begin) != null)
                                    {
                                        rBegin = aBegin;
                                        rBeginFound = true;
                                    }
                                    else
                                    {
                                        if ((aBegin.GetType() == typeof(Run)) &&
                                           (aBegin.Elements<FieldChar>()?.FirstOrDefault().FieldCharType == FieldCharValues.End
                                           || aBegin.Elements<FieldChar>()?.FirstOrDefault().FieldCharType == FieldCharValues.Separate)
                                           )
                                        {
                                            break;
                                        }
                                        else
                                        {//fillerRuns1[k] = new Run();
                                            fillerRuns1.Add(aBegin);
                                        }
                                    }
                                }
                                pivotRun = aBegin;
                                k++;
                            }

                            //if (null != xxxfield) rSep = xxxfield.NextSibling<Run>();
                            k = 0;
                            pivotRun = xxxfield;
                            while (null != pivotRun && (rSep == null && k < 10))
                            {
                                var aSep = pivotRun.NextSibling<Run>();
                                if (null != aSep)
                                {

                                    if (aSep.GetType() == typeof(Run) &&
                                        aSep.Elements<FieldChar>().FirstOrDefault(fc =>
                                            fc.FieldCharType == FieldCharValues.Separate) != null)
                                    {
                                        rSep = aSep;
                                        rSepFound = true;
                                    }
                                    else
                                    {
                                        if ((aSep.GetType() == typeof(Run)) &&
                                           (aSep.Elements<FieldChar>()?.FirstOrDefault().FieldCharType == FieldCharValues.Begin
                                           || aSep.Elements<FieldChar>()?.FirstOrDefault().FieldCharType == FieldCharValues.End)
                                           )
                                        {
                                            break;
                                        }
                                        else
                                        {//fillerRuns2[k] = new Run();
                                            fillerRuns2.Add(aSep);
                                        }
                                    }
                                }
                                pivotRun = aSep;
                                k++;
                            }

                            //if (null != rSep) rText = rSep.NextSibling<Run>();
                            k = 0;
                            pivotRun = rSep;
                            while (null != pivotRun && (rText == null && k < 10))
                            {
                                var aText = pivotRun.NextSibling<Run>();
                                if (null != aText)
                                {

                                    if (aText.GetType() == typeof(Run) &&
                                        aText.GetFirstChild<Text>() != null)
                                    {
                                        rText = aText;
                                        rTextFound = true;
                                    }
                                    else
                                    {
                                        //fillerRuns3[k] = new Run();
                                        if ((aText.GetType() == typeof(Run)) &&
                                           (aText.Elements<FieldChar>()?.FirstOrDefault().FieldCharType == FieldCharValues.Begin
                                           || aText.Elements<FieldChar>()?.FirstOrDefault().FieldCharType == FieldCharValues.Separate
                                           || aText.Elements<FieldChar>()?.FirstOrDefault().FieldCharType == FieldCharValues.End)
                                           )
                                        {
                                            break;
                                        }
                                        else
                                        {
                                            fillerRuns3.Add(aText);
                                        }
                                    }
                                }
                                pivotRun = aText;
                                k++;
                            }


                            //if (null != rText) rEnd = rText.NextSibling<Run>();
                            k = 0;
                            pivotRun = rText;
                            while (null != pivotRun && (rEnd == null && k < 10))
                            {
                                var aEnd = pivotRun.NextSibling<Run>();
                                if (null != aEnd)
                                {

                                    if (aEnd.GetType() == typeof(Run) &&
                                        aEnd.Elements<FieldChar>().FirstOrDefault(fc =>
                                            fc.FieldCharType == FieldCharValues.End) != null)
                                    {
                                        rEnd = aEnd;
                                        rEndFound = true;
                                    }
                                    else
                                    {
                                        //fillerRuns4[k] = new Run();
                                        if ((aEnd.GetType() == typeof(Run)) &&
                                           (aEnd.Elements<FieldChar>()?.FirstOrDefault().FieldCharType == FieldCharValues.Begin
                                           || aEnd.Elements<FieldChar>()?.FirstOrDefault().FieldCharType == FieldCharValues.Separate)
                                           )
                                        {
                                            break;
                                        }
                                        else
                                        {
                                            fillerRuns4.Add(aEnd);
                                        }
                                    }
                                }
                                pivotRun = aEnd;
                                k++;
                            }

                            k = 0;

                            /*
                            Run innerRun = null;
                            if (!rBeginFound && null == rBegin)
                            {
                                if (fillerRuns1.Count > 0) { fillerRuns1.Clear(); }
                                if (null != rPara)
                                {
                                    Run first = (Run)rPara.FirstChild;
                                    if (null != first)
                                    {
                                        int l = 0;
                                        innerRun = first;
                                        while (null != innerRun && (rBegin == null && l < 10))
                                        {
                                            if (innerRun == xxxfield)
                                            {
                                                Run rField = new Run();
                                                rField = (Run) xxxfield.Clone();
                                                innerRun.Append(rField);
                                                innerRun = new Run(new FieldChar() { FieldCharType = FieldCharValues.Begin });
                                                rBegin = innerRun;
                                                rBeginFound = true;
                                            }
                                            else
                                            {
                                                var aField = innerRun.NextSibling<Run>();
                                                if (null != aField)
                                                {

                                                    if (aField.GetType() == typeof(Run) &&
                                                        aField == xxxfield)
                                                    {
                                                        rBegin = new Run(new FieldChar() { FieldCharType = FieldCharValues.Begin });
                                                        innerRun.Append(rBegin);                                                        
                                                        rBeginFound = true;
                                                    }
                                                }
                                                innerRun = aField;
                                            }
                                            l++;
                                        }
                                    }
                                }
                            }

                            k = 0;
                            innerRun = null;
                            if (!rSepFound && null == rSep)
                            {
                                if (null != xxxfield) innerRun = xxxfield;
                                if (fillerRuns2.Count > 0) { fillerRuns2.Clear(); }
                                if (null != innerRun)
                                {
                                    rSep = new Run(new FieldChar() { FieldCharType = FieldCharValues.Separate });
                                    innerRun.Append(rSep);
                                    rSepFound = true;
                                }
                                
                            }

                            k = 0;
                            innerRun = null;
                            if (!rTextFound && null == rText)
                            {
                                if (null != rSep) innerRun = rSep;
                                if (fillerRuns3.Count > 0) { fillerRuns3.Clear(); }
                                if (null != innerRun)
                                {
                                    rText = new Run(new Text(" "));
                                    innerRun.Append(rText);
                                    rTextFound = true;
                                }
                            }

                            k = 0;
                            innerRun = null;
                            if (!rEndFound && null == rEnd)
                            {
                                if (null != rText) innerRun = rText;
                                if (fillerRuns4.Count > 0) { fillerRuns4.Clear(); }
                                if (null != innerRun)
                                {
                                    rEnd = new Run(new FieldChar() { FieldCharType = FieldCharValues.End });
                                    innerRun.Append(rEnd);
                                    rEndFound = true;
                                }
                            }
                            */

                            //|| !rFieldFound || !rSepFound || !rEndFound || !rTextFound


                            if (fillerRuns1.Count > 0) { fillerRuns.AddRange(fillerRuns1); }
                            if (fillerRuns2.Count > 0) { fillerRuns.AddRange(fillerRuns2); }
                            if (fillerRuns3.Count > 0) { fillerRuns.AddRange(fillerRuns3); }
                            if (fillerRuns4.Count > 0) { fillerRuns.AddRange(fillerRuns4); }


                            if (null != rText) t = rText.GetFirstChild<Text>();
                            if (null != rText) rProp = (RunProperties)rText.RunProperties;
                            if (null != t)
                            {
                                if (null != t.Text)
                                {
                                    //aFieldName=formatText(t.Text);
                                    aFieldName = j.ToString().Trim();
                                    if (String.IsNullOrEmpty(aFieldName) || String.IsNullOrWhiteSpace(t.Text)) aFieldName = j.ToString().Trim();
                                }
                            }
                        }
                        else
                        {

                            Regex expr = new Regex(@"\s*(?<docProperty>\S+)\s+(?<aFieldName>\S+)\s+(?<dummy>\S+)\s*(?<formatType>\S*)\s*");

                            var results = expr.Matches(field.Text.ToString().Trim());
                            string docProperty = null;
                            aFieldName = null;
                            string formatType = null;
                            string dummy = null;
                            bool isCorrectFormat = true;

                            foreach (Match match in results)
                            {
                                docProperty = match.Groups["docProperty"].Value;
                                aFieldName = match.Groups["aFieldName"].Value;
                                formatType = match.Groups["formatType"].Value;
                                dummy = match.Groups["dummy"].Value;
                            }

                            if (null == results || results?.Count < 1 || null == aFieldName ||
                                aFieldName.Contains("\\*") || aFieldName.Contains("DOCPROPERTY") || aFieldName.Contains("CHARFORMAT")
                            //|| (formatType?.Contains("CHARFORMAT") ?? false)
                            )
                            {
                                isCorrectFormat = false;
                                expr = new Regex(@"\s*(?<docProperty>\S+)\s+(?<dummy>\S+)\s*(?<formatType>\S*)\s+(?<aFieldName>\S+)\s*");
                                results = expr.Matches(field.Text.ToString().Trim());

                                foreach (Match match in results)
                                {
                                    docProperty = match.Groups["docProperty"].Value;
                                    aFieldName = match.Groups["aFieldName"].Value;
                                    formatType = match.Groups["formatType"].Value;
                                    dummy = match.Groups["dummy"].Value;
                                }

                                if (null == results || results?.Count < 1 || null == aFieldName ||
                                aFieldName.Contains("\\*") || aFieldName.Contains("DOCPROPERTY") || aFieldName.Contains("CHARFORMAT"))
                                {
                                    isCorrectFormat = false;
                                }
                                else
                                {
                                    isCorrectFormat = true;
                                }
                            }

                            if (!isCorrectFormat) continue;




                            /*xxxfield = (Run)field.Parent;
                            if (null != xxxfield) rPara = (Paragraph)xxxfield.Parent;

                            while (null != xxxfield && rBegin == null)
                            {
                                var aBegin = xxxfield.PreviousSibling<Run>();
                                //if (null != aBegin) aBegin.
                                if (null != aBegin) rBegin = aBegin;
                            }
                            if (null != xxxfield) rSep = xxxfield.NextSibling<Run>();
                            if (null != rSep) rText = rSep.NextSibling<Run>();
                            if (null != rText) rEnd = rText.NextSibling<Run>();
                            if (null != rText) t = rText.GetFirstChild<Text>();
                            if (null != rText) rProp = (RunProperties)rText.RunProperties;*/
                            xxxfield = (Run)field?.Parent;
                            if (null != xxxfield)
                            {
                                rFieldFound = true;
                                rPara = (Paragraph)xxxfield.Parent;
                            }

                            k = 0;
                            pivotRun = xxxfield;
                            while (null != pivotRun && (rBegin == null && k < 10))
                            {
                                var aBegin = pivotRun.PreviousSibling<Run>();
                                if (null != aBegin)
                                {

                                    if (aBegin.GetType() == typeof(Run) &&
                                        aBegin.Elements<FieldChar>().FirstOrDefault(fc =>
                                            fc.FieldCharType == FieldCharValues.Begin) != null)
                                    {
                                        rBegin = aBegin;
                                        rBeginFound = true;
                                    }
                                    else
                                    {
                                        if ((aBegin.GetType() == typeof(Run)) &&
                                           (aBegin.Elements<FieldChar>()?.FirstOrDefault().FieldCharType == FieldCharValues.End
                                           || aBegin.Elements<FieldChar>()?.FirstOrDefault().FieldCharType == FieldCharValues.Separate)
                                           )
                                        {
                                            break;
                                        }
                                        else
                                        {//fillerRuns1[k] = new Run();
                                            fillerRuns1.Add(aBegin);
                                        }
                                    }
                                }
                                pivotRun = aBegin;
                                k++;
                            }

                            //if (null != xxxfield) rSep = xxxfield.NextSibling<Run>();
                            k = 0;
                            pivotRun = xxxfield;
                            while (null != pivotRun && (rSep == null && k < 10))
                            {
                                var aSep = pivotRun.NextSibling<Run>();
                                if (null != aSep)
                                {

                                    if (aSep.GetType() == typeof(Run) &&
                                        aSep.Elements<FieldChar>().FirstOrDefault(fc =>
                                            fc.FieldCharType == FieldCharValues.Separate) != null)
                                    {
                                        rSep = aSep;
                                        rSepFound = true;
                                    }
                                    else
                                    {
                                        if ((aSep.GetType() == typeof(Run)) &&
                                           (aSep.Elements<FieldChar>()?.FirstOrDefault().FieldCharType == FieldCharValues.Begin
                                           || aSep.Elements<FieldChar>()?.FirstOrDefault().FieldCharType == FieldCharValues.End)
                                           )
                                        {
                                            break;
                                        }
                                        else
                                        {//fillerRuns2[k] = new Run();
                                            fillerRuns2.Add(aSep);
                                        }
                                    }
                                }
                                pivotRun = aSep;
                                k++;
                            }

                            //if (null != rSep) rText = rSep.NextSibling<Run>();
                            k = 0;
                            pivotRun = rSep;
                            while (null != pivotRun && (rText == null && k < 10))
                            {
                                var aText = pivotRun.NextSibling<Run>();
                                if (null != aText)
                                {

                                    if (aText.GetType() == typeof(Run) &&
                                        aText.GetFirstChild<Text>() != null)
                                    {
                                        rText = aText;
                                        rTextFound = true;
                                    }
                                    else
                                    {
                                        //fillerRuns3[k] = new Run();
                                        if ((aText.GetType() == typeof(Run)) &&
                                           (aText.Elements<FieldChar>()?.FirstOrDefault().FieldCharType == FieldCharValues.Begin
                                           || aText.Elements<FieldChar>()?.FirstOrDefault().FieldCharType == FieldCharValues.Separate
                                           || aText.Elements<FieldChar>()?.FirstOrDefault().FieldCharType == FieldCharValues.End)
                                           )
                                        {
                                            break;
                                        }
                                        else
                                        {
                                            fillerRuns3.Add(aText);
                                        }
                                    }
                                }
                                pivotRun = aText;
                                k++;
                            }


                            //if (null != rText) rEnd = rText.NextSibling<Run>();
                            k = 0;
                            pivotRun = rText;
                            while (null != pivotRun && (rEnd == null && k < 10))
                            {
                                var aEnd = pivotRun.NextSibling<Run>();
                                if (null != aEnd)
                                {

                                    if (aEnd.GetType() == typeof(Run) &&
                                        aEnd.Elements<FieldChar>().FirstOrDefault(fc =>
                                            fc.FieldCharType == FieldCharValues.End) != null)
                                    {
                                        rEnd = aEnd;
                                        rEndFound = true;
                                    }
                                    else
                                    {
                                        //fillerRuns4[k] = new Run();
                                        if ((aEnd.GetType() == typeof(Run)) &&
                                           (aEnd.Elements<FieldChar>()?.FirstOrDefault().FieldCharType == FieldCharValues.Begin
                                           || aEnd.Elements<FieldChar>()?.FirstOrDefault().FieldCharType == FieldCharValues.Separate)
                                           )
                                        {
                                            break;
                                        }
                                        else
                                        {
                                            fillerRuns4.Add(aEnd);
                                        }
                                    }
                                }
                                pivotRun = aEnd;
                                k++;
                            }

                            //if (null != rText) rEnd = rText.NextSibling<Run>();
                            k = 0;
                            /*Run innerRun = null;
                            if (!rBeginFound && null == rBegin)
                            {
                                if (fillerRuns1.Count > 0) { fillerRuns1.Clear(); }
                                if (null != rPara)
                                {
                                    Run first = (Run)rPara.FirstChild;
                                    if (null != first)
                                    {
                                        int l = 0;
                                        innerRun = first;
                                        while (null != innerRun && (rBegin == null && l < 10))
                                        {
                                            if (innerRun == xxxfield)
                                            {
                                                Run rField = new Run();
                                                rField = (Run)xxxfield.Clone();
                                                innerRun.Append(rField);
                                                innerRun = new Run(new FieldChar() { FieldCharType = FieldCharValues.Begin });
                                                rBegin = innerRun;
                                                rBeginFound = true;
                                            }
                                            else
                                            {
                                                var aField = innerRun.NextSibling<Run>();
                                                if (null != aField)
                                                {

                                                    if (aField.GetType() == typeof(Run) &&
                                                        aField == xxxfield)
                                                    {
                                                        rBegin = new Run(new FieldChar() { FieldCharType = FieldCharValues.Begin });
                                                        innerRun.Append(rBegin);
                                                        rBeginFound = true;
                                                    }
                                                }
                                                innerRun = aField;
                                            }
                                            l++;
                                        }
                                    }
                                }
                            }

                            k = 0;
                            innerRun = null;
                            if (!rSepFound && null == rSep)
                            {
                                if (null != xxxfield) innerRun = xxxfield;
                                if (fillerRuns2.Count > 0) { fillerRuns2.Clear(); }
                                if (null != innerRun)
                                {
                                    rSep = new Run(new FieldChar() { FieldCharType = FieldCharValues.Separate });
                                    innerRun.Append(rSep);
                                    rSepFound = true;
                                }

                            }

                            k = 0;
                            innerRun = null;
                            if (!rTextFound && null == rText)
                            {
                                if (null != rSep) innerRun = rSep;
                                if (fillerRuns3.Count > 0) { fillerRuns3.Clear(); }
                                if (null != innerRun)
                                {
                                    rText = new Run(new Text(" "));
                                    innerRun.Append(rText);
                                    rTextFound = true;
                                }
                            }

                            k = 0;
                            innerRun = null;
                            if (!rEndFound && null == rEnd)
                            {
                                if (null != rText) innerRun = rText;
                                if (fillerRuns4.Count > 0) { fillerRuns4.Clear(); }
                                if (null != innerRun)
                                {
                                    rEnd = new Run(new FieldChar() { FieldCharType = FieldCharValues.End });
                                    innerRun.Append(rEnd);
                                    rEndFound = true;
                                }
                            }
                            */


                            if (fillerRuns1.Count > 0) { fillerRuns.AddRange(fillerRuns1); }
                            if (fillerRuns2.Count > 0) { fillerRuns.AddRange(fillerRuns2); }
                            if (fillerRuns3.Count > 0) { fillerRuns.AddRange(fillerRuns3); }
                            if (fillerRuns4.Count > 0) { fillerRuns.AddRange(fillerRuns4); }

                        }

                        if (!String.IsNullOrWhiteSpace(aFieldName))
                            aFieldName = aFieldName.ToUpper().Trim();

                        DocField aDocField = new DocField(aFieldName, rBegin, xxxfield, rSep, rText, rEnd, rPara, rProp, rParent, isFormField,
                            fillerRuns, null, fillerRuns1, fillerRuns2, fillerRuns3, fillerRuns4);

                        //DocField(string id, Run begin, Run label, Run sep, Run text, Run end, Paragraph para, RunProperties prop, Run parent = null)
                        /*if (docFields.ContainsKey(aFieldName)) docFields[aFieldName] = aDocField;
                        else docFields.Add(aFieldName, aDocField);*/
                        docFields.Add((aFieldName, aDocField));

                    }

                    /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                    // First Pass finished
                    /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                    j = 0;
                    Console.Out.WriteLine("##### Total num of fields " + docFields.Count + " #####");
                    Console.Out.WriteLine("##### DocFields " + docFields.ToArray().ToString() + " #####");

                    /*Console.WriteLine("Title = " + wordDocument.ExtendedFilePropertiesPart.Properties.TitlesOfParts.InnerText);
                    Console.WriteLine("Subject = " + wordDocument.PackageProperties.Subject);
                    {
                        foreach (var control in wordDocument.ContentControls())
                        {
                            Console.WriteLine(control.Title + "==>" + control.Value);
                        }
                    }*/

                    foreach (var docField in docFields)
                    {
                        var field = docField.Item1;
                        j++;
                        Console.Out.WriteLine("*** Field " + j.ToString() + " starts *********************************");
                        Console.Out.WriteLine("Field Type:" + field.ToString());
                        //Console.Out.WriteLine("Field Text:>>>" + field.Text.ToString() + "<<<");

                        //var fieldId = field.Text.Trim().ToString();
                        var fieldId = field.ToString();
                        ; // breakpoint anchor
                        /*if (null != fieldId)
                        {
                            if (fieldId.Contains("FORMTEXT"))
                            {
                                isFormField = true;
                            }
                            Console.Out.WriteLine("fieldId:" + fieldId.ToString());
                        }*/




                        /*if (field.Text.ToString() == " DOCPROPERTY  WorkerName  \\* CHARFORMAT ")
                         {
                             field.Text = new string("REPLACED WORKER NAME");


                             Console.Out.WriteLine("Replaceing Field Text:>>>" + field.Text.ToString() + "<<<");
                             Console.Out.WriteLine("with:>>>REPLACED_WORKER_NAME<<<");
                         }*/

                        DocField aDocField = docField.Item2;

                        bool isFormField = aDocField.isFormField;

                        Run xxxfield = null;
                        Run rBegin = null;
                        Run rSep = null;
                        Run rText = null;
                        Run rEnd = null;
                        Text t = null;

                        Run nFormat = null;
                        Run nBegin = null;
                        Run nTag = null;
                        Run nSep = null;
                        Run nText = null;
                        Run nEnd = null;
                        Text nt = null;

                        Run nnFormat = null;
                        Run nnBegin = null;
                        Run nnTag = null;
                        Run nnSep = null;
                        Run nnText = null;
                        Run nnEnd = null;
                        Text nnt = null;


                        Run pivotRun = null;
                        Run pivotEnd = null;
                        Run pivotBegin = null;
                        Run rAdd = null;

                        xxxfield = aDocField.fieldLabel;
                        rBegin = aDocField.fieldBegin;
                        rSep = aDocField.fieldSep;
                        rEnd = aDocField.fieldEnd;
                        rText = aDocField.fieldText;
                        if (null != rText) t = rText.GetFirstChild<Text>();
                        string aFieldName = field;

                        ArrayList fillerRuns = aDocField.fillerRuns ?? null;



                        /*xxxfield = (Run)field.Parent;
                        while (null != xxxfield && rBegin == null)
                        {
                            var aBegin = xxxfield.PreviousSibling<Run>();
                            //if (null != aBegin) aBegin.
                            if (null != aBegin) rBegin = aBegin;
                        }
                        if (null != xxxfield) rSep = xxxfield.NextSibling<Run>();
                        if (null != rSep) rText = rSep.NextSibling<Run>();
                        if (null != rText) rEnd = rText.NextSibling<Run>();
                        if (null != rText) t = rText.GetFirstChild<Text>();*/


                        /*DocumentFormat.OpenXml.Wordprocessing.Run run2 = new DocumentFormat.OpenXml.Wordprocessing.Run(new Text() { Text = "Table ", Space = SpaceProcessingModeValues.Preserve });
                        SimpleField simpleField2 = new SimpleField(new DocumentFormat.OpenXml.Wordprocessing.Run(new RunProperties(new NoProof()), new Text() { Text = " ", Space = SpaceProcessingModeValues.Preserve }));
                        simpleField2.Instruction = @"SEQ " + "Table";

                        DocumentFormat.OpenXml.Wordprocessing.Paragraph refP = new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                          new DocumentFormat.OpenXml.Wordprocessing.Run(new FieldChar() { FieldCharType = FieldCharValues.Begin }),
                          //todo instrTxt bookmark ref
                          new DocumentFormat.OpenXml.Wordprocessing.Run(new FieldChar() { FieldCharType = FieldCharValues.Separate }),
                          run2,
                          simpleField2,
                          new DocumentFormat.OpenXml.Wordprocessing.Run(new FieldChar() { FieldCharType = FieldCharValues.End })

                          );

                        */













                        pivotRun = xxxfield;
                        pivotBegin = rBegin;
                        pivotEnd = rEnd;
                        bool found = false;

                        Console.Out.WriteLine("@@@ Checking field Id...");
                        if (null != xxxfield)
                        {
                            if (null != xxxfield.InnerText) Console.Out.WriteLine("xxxfield:" + xxxfield.ToString() + ">>>" + xxxfield.InnerText);
                        }
                        if (null != rBegin)
                        {
                            if (null != rBegin.InnerText) Console.Out.WriteLine("rBegin:" + rBegin.ToString() + ">>>" + rBegin.InnerText);
                        }
                        if (null != rText)
                        {
                            if (null != rText.InnerText) Console.Out.WriteLine("rText:" + rText.ToString() + ">>>" + rText.InnerText);
                        }
                        if (null != rEnd)
                        {
                            if (null != rEnd.InnerText) Console.Out.WriteLine("rEnd:" + rEnd.ToString() + ">>>" + rEnd.InnerText);
                        }
                        if (null != t)
                        {
                            if (null != t.InnerText) Console.Out.WriteLine("t:" + t.ToString() + ">>>" + t.InnerText);
                        }


                        if (isFormField)
                        {
                            if (null != t)
                            {
                                if (null != t.Text)
                                {
                                    var replaceFormValue = "";
                                    Console.Out.WriteLine("****Substitute form field " + t.Text);
                                    if (String.IsNullOrWhiteSpace(t.Text.Trim()))
                                        replaceFormValue = "[     ]";
                                    else
                                        replaceFormValue = "[" + t.Text.Trim() + "]";
                                    t.Text = formatText(replaceFormValue);
                                    if (null != rEnd) rEnd.Remove();
                                    if (null != rSep) rSep.Remove();
                                    if (null != xxxfield) xxxfield.Remove();
                                    if (null != rBegin) rBegin.Remove();
                                    //SetFormFieldValue(t.Text, formatText(replaceFormValue));

                                    if (null != rText)
                                    {
                                        var rp = rText?.RunProperties;
                                        Highlight highlight = new Highlight() { Val = HighlightColorValues.Yellow };
                                        if (null != rp)
                                        {
                                            rp.Append(highlight);
                                        }
                                        else
                                        {
                                            rText.RunProperties = new RunProperties();
                                            if (null != rText.RunProperties) rText.RunProperties.Append(highlight);
                                        }
                                    }


                                }
                            }
                        }
                        else
                        {

                            /*Regex expr = new Regex(@"\s*(?<docProperty>\S+)\s+(?<aFieldName>\S+)\.*\s+(?<formatType>\S+)\s*");
                            var results = expr.Matches(fieldId);
                            string docProperty = null;
                            string aFieldName = null;
                            string formatType = null;

                            foreach (Match match in results)
                            {
                                docProperty = match.Groups["docProperty"].Value;
                                aFieldName = match.Groups["aFieldName"].Value;
                                formatType = match.Groups["formatType"].Value;
                            }

                            if (null != docProperty) Console.Out.WriteLine("docProperty=" + docProperty);
                            if (null != aFieldName) Console.Out.WriteLine("aFieldName=" + aFieldName);
                            if (null != formatType) Console.Out.WriteLine("formatType=" + formatType);*/

                            if (null != aFieldName)
                            {
                                var aFieldKey = aFieldName.Trim().ToUpper();
                                string aFieldValue = null;
                                if (fieldMap.ContainsKey(aFieldKey))
                                {
                                    if (null != fieldMap[aFieldKey])
                                    {
                                        aFieldValue = fieldMap[aFieldKey].Trim();
                                    }
                                }
                                if (String.IsNullOrEmpty(aFieldValue)) aFieldValue = new string("");

                                if (t != null)
                                {
                                    if (t.Text != null && aFieldKey != null && fieldMap.ContainsKey(aFieldKey))
                                    {
                                        if (!fieldMap.ContainsKey(aFieldKey) || (String.IsNullOrWhiteSpace(aFieldValue)
                                        || aFieldValue.Contains("BOOKMARK_UNDEFINED")))
                                        {
                                            Run rBegin2 = null;
                                            Run rBegin1 = null;
                                            Run rBegin0 = null;
                                            Paragraph rParent = null;
                                            Run rParentFirst = null;
                                            Run rParentLeft = null;
                                            Run rParentLeftFirst = null;
                                            if (null != rBegin) rBegin2 = rBegin.PreviousSibling<Run>();
                                            if (null != rBegin2) rBegin1 = rBegin2.PreviousSibling<Run>();
                                            if (null != rBegin1) rBegin0 = rBegin1.PreviousSibling<Run>();
                                            if (null != rText) rParent = (Paragraph)rText.Parent;
                                            /*if (null != rParent) rParentFirst = (Run)rParent.GetFirstChild<Run>();
                                            if (null != rParent) rParentLeft = rParent.PreviousSibling<Run>();
                                            if (null != rParentLeft) rParentLeftFirst = rParentLeft.GetFirstChild<Run>();*/
                                            //t.Text = " ";
                                            //if (null != rEnd) rAdd = rEnd.NextSibling<Run>();
                                            //if (null != rAdd)
                                            //    rAdd.AppendChild(new Text(" "));
                                            //rAdd.AppendChild(new Text(fieldMap[aFieldKey]));*/
                                            //if (null != rParent) rParent.InnerXml.Replace(fieldId, "");
                                            //if (null != rText) rText.RemoveAllChildren();
                                            //if (null != rText) rText.Remove();
                                            if (null != t) t.Text = formatText(aFieldValue);
                                            if (null != rEnd) rEnd.Remove();
                                            if (null != rSep) rSep.Remove();
                                            if (null != rBegin) rBegin.Remove();
                                            if (null != xxxfield) xxxfield.Remove();
                                            //if (null != rParent) rParent.RemoveAllChildren();
                                            //if (null != rParent) rParent.AppendChild<Run>(new Run(new Text("")));                                        
                                            /*if (null != rBegin2) rBegin2.RemoveAllChildren();
                                            if (null != rBegin2) rBegin2.Remove();
                                            if (null != rBegin1) rBegin1.RemoveAllChildren();
                                            if (null != rBegin1) rBegin1.Remove();
                                            if (null != rBegin0) rBegin0.RemoveAllChildren();
                                            if (null != rBegin0) rBegin0.Remove();*/
                                            //if (null != t.Text) Console.Out.WriteLine("****Substitute value " + t.Text + "with " + fieldMap[aFieldKey]);
                                            //rText.Remove();
                                        }
                                        else
                                        {
                                            Console.Out.WriteLine("****Substitute value " + t.Text + " with " + aFieldValue);
                                            if (fieldMap.ContainsKey(aFieldKey) && !(String.IsNullOrEmpty(aFieldValue)) && !(String.IsNullOrWhiteSpace(aFieldValue)))
                                            {



                                                if (nonEditableArray.Contains(aFieldKey))
                                                {
                                                    t.Text = formatText(aFieldValue);
                                                }
                                                else
                                                {


                                                    Run rBegin2 = null;
                                                    Run rBegin1 = null;
                                                    Run rBegin0 = null;
                                                    Paragraph rParent = null;
                                                    Run rParentFirst = null;
                                                    Run rParentLeft = null;
                                                    Run rParentLeftFirst = null;
                                                    if (null != rBegin) rBegin2 = rBegin.PreviousSibling<Run>();
                                                    if (null != rBegin2) rBegin1 = rBegin2.PreviousSibling<Run>();
                                                    if (null != rBegin1) rBegin0 = rBegin1.PreviousSibling<Run>();
                                                    if (null != rText) rParent = (Paragraph)rText.Parent;
                                                    /*if (null != rParent) rParentFirst = (Run)rParent.GetFirstChild<Run>();
                                                    if (null != rParent) rParentLeft = rParent.PreviousSibling<Run>();
                                                    if (null != rParentLeft) rParentLeftFirst = rParentLeft.GetFirstChild<Run>();*/
                                                    //t.Text = fieldMap[aFieldKey];
                                                    //if (null != rEnd) rAdd = rEnd.NextSibling<Run>();
                                                    //if (null != rAdd)
                                                    //    rAdd.AppendChild(new Text(fieldMap[aFieldKey]));
                                                    //rAdd = rText.NextSibling<Run>();
                                                    /*if (null != rAdd)
                                                        rAdd.AppendChild(new Text(fieldMap[aFieldKey]));*/
                                                    //if (null != rParent) rParent.InnerXml.Replace(fieldId, "");
                                                    //if (null != rText) rText.RemoveAllChildren();
                                                    //if (null != rText) rText.Remove();

                                                    string aaFieldKey = null;
                                                    string aaFieldValue = null;
                                                    string aaaFieldKey = null;
                                                    string aaaFieldValue = null;

                                                    if (null != rEnd) nFormat = rEnd.NextSibling<Run>();
                                                    if (null != nFormat) nBegin = nFormat.NextSibling<Run>();
                                                    if (null != nBegin) nTag = nBegin.NextSibling<Run>();
                                                    if (null != nTag) nSep = nTag.NextSibling<Run>();
                                                    if (null != nSep) nText = nSep.NextSibling<Run>();
                                                    if (null != nText) nEnd = nText.NextSibling<Run>();
                                                    if (null != nText) nt = nText.GetFirstChild<Text>();
                                                    if (null != nEnd) nnFormat = nEnd.NextSibling<Run>();
                                                    if (null != nnFormat) nnBegin = nnFormat.NextSibling<Run>();
                                                    if (null != nnBegin) nnTag = nnBegin.NextSibling<Run>();
                                                    if (null != nnTag) nnSep = nnTag.NextSibling<Run>();
                                                    if (null != nnSep) nnText = nnSep.NextSibling<Run>();
                                                    if (null != nnText) nnEnd = nnText.NextSibling<Run>();
                                                    if (null != nnText) nnt = nnText.GetFirstChild<Text>();


                                                    /*if (null != nnt)
                                                    {
                                                        if (null != nnt.Text && !String.IsNullOrWhiteSpace(nnt.Text))
                                                        {
                                                            aaaFieldKey = nnt.Text.Trim().ToUpper();
                                                            if (null != aaaFieldKey && fieldMap.ContainsKey(aaaFieldKey)) aaaFieldValue = fieldMap[aaaFieldKey];
                                                            if (null != aaaFieldValue && !String.IsNullOrWhiteSpace(aaaFieldValue) && aaaFieldValue.ToUpper() != "BOOKMARK_UNDEFINED")
                                                            {
                                                                nnt.Text = formatText(aaaFieldValue);
                                                                if (null != nnEnd) nnEnd.Remove();
                                                                if (null != nnSep) nnSep.Remove();
                                                                if (null != nnTag) nnTag.Remove();
                                                                if (null != nnBegin) nnBegin.Remove();
                                                                //if (null != rText) rText = new Run(new Text(aFieldValue));

                                                            }
                                                        }
                                                    }

                                                    if (null != nt)
                                                    {
                                                        if (null != nt.Text && !String.IsNullOrWhiteSpace(nt.Text))
                                                        {
                                                            aaFieldKey = nt.Text.Trim().ToUpper();
                                                            if (null != aaFieldKey && fieldMap.ContainsKey(aaFieldKey)) aaFieldValue = fieldMap[aaFieldKey];
                                                            if (null != aaFieldValue && !String.IsNullOrWhiteSpace(aaFieldValue) && aaFieldValue.ToUpper() != "BOOKMARK_UNDEFINED")
                                                            {
                                                                nt.Text = formatText(aaFieldValue);
                                                                if (null != nEnd) nEnd.Remove();
                                                                if (null != nSep) nSep.Remove();
                                                                if (null != nTag) nTag.Remove();
                                                                if (null != nBegin) nBegin.Remove();
                                                            }
                                                        }
                                                    }*/






                                                    if (null != t && null != t.Text) t.Text = formatText(aFieldValue);
                                                    if (null != rEnd) rEnd.Remove();
                                                    if (null != rSep) rSep.Remove();
                                                    if (null != rBegin) rBegin.Remove();
                                                    if (null != xxxfield) xxxfield.Remove();
                                                    //if (null != rText) rText = new Run(new Text(aFieldValue));
                                                    //if (null != rParent) rParent.RemoveAllChildren();
                                                    //if (null != rParent) rParent.AppendChild<Run>(new Run(new Text(aFieldValue)));
                                                    /*if (null != rBegin2) rBegin2.RemoveAllChildren();
                                                    if (null != rBegin2) rBegin2.Remove();
                                                    if (null != rBegin1) rBegin1.RemoveAllChildren();
                                                    if (null != rBegin1) rBegin1.Remove();
                                                    if (null != rBegin0) rBegin0.RemoveAllChildren();
                                                    if (null != rBegin0) rBegin0.Remove();*/


                                                }
                                            }
                                        }

                                    }
                                }
                            }




                            /*
                            }
                            else
                            {
                                Console.Out.WriteLine("@@@ field value not found.");
                            }
                            */
                        }

                        if (null != fillerRuns && fillerRuns.Count > 0)
                        {
                            foreach (var fillerRun in fillerRuns)
                            {
                                if (null != fillerRun)//&& fillerRun is typeof(Run))
                                    ((Run)fillerRun).Remove();
                            }
                        }
                        Console.Out.WriteLine("*** Field " + j.ToString() + " ends *********************************");

                        //prevRun = pivotRun;
                        //prevBegin = pivotBegin;
                        //prevEnd = pivotEnd;



                    }




                    setCustomProperty(wordDocument, WopiOptions.Value.ProcessedFlag, WopiOptions.Value.ApplicationName, CustomPropertyTypes.Text);

                    /*Remove VBA part
                        var docPart = wordDocument.MainDocumentPart;

                        // Look for the vbaProject part. If it is there, delete it.
                        var vbaPart = docPart.VbaProjectPart;
                        if (vbaPart != null)
                        {
                            // Delete the vbaProject part and then save the document.
                            docPart.DeletePart(vbaPart);
                            docPart.Document.Save();

                            // Change the document type to
                            // not macro-enabled.
                            wordDocument.ChangeDocumentType(WordprocessingDocumentType.Document);

                            // Track that the document has been changed.

                        }
                        //changeCompatibilityModeOfDocumentPart(wordDocument.MainDocumentPart);
                    */
                    /*}
                    catch (Exception ax)
                    {
                        throw new WordWebException("Error in running macro", ax);
                    }
                    finally
                    {*/
                    wordDocument.Save();
                    wordDocument.Close();
                    //}

                } // end if
            } // end using

        }









        public void runMacroX2(string newFileName, Dictionary<string, string> fieldMap, bool processedByWWA = false, bool normalize = false, string userRole = null)
        {
            Regex fileVer = new Regex(@"\s*(?<userID>)_(?<attachmentID>)_(?<versionNo>)\..*");
            var result = fileVer.Matches(newFileName);
            string userID = null;
            string attachmentID = null;
            string versionNo = null;

            /*if (Int32.TryParse(versionNo, out int v))
            {
                if (v > 1) return;
            }*/
            //if (processedByWWA) return;

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(newFileName, true))
            {

                if (null != wordDocument)
                {
                    /*for (var field in wordDocument.MainDocumentPart.RootElement.Descendants<FieldCode>())
                    {

                    }*/

                    /*List<Picture> pictures = new List<Picture>(wordDocument.MainDocumentPart.RootElement.Descendants<Picture>());

                    foreach (Picture p in pictures)
                    {
                        p.Remove();
                    }*/

                    //headerPart.DeleteParts(imagePartList);

                    foreach (Paragraph p in wordDocument.MainDocumentPart.Document.Body.Descendants<Paragraph>().Where<Paragraph>(p => p.InnerText.Contains(WopiOptions.Value.ConversionEngine)))
                    {
                        // Do something with the Paragraphs.
                        p.Remove();
                    }

                    if (wordDocument.MainDocumentPart.HeaderParts.Count() > 0)
                    {
                        foreach (HeaderPart headerPart in wordDocument.MainDocumentPart.HeaderParts)
                        {
                            foreach (Paragraph p in headerPart.Header.Descendants<Paragraph>().Where<Paragraph>(p => p.InnerText.Contains(WopiOptions.Value.ConversionEngine)))
                            {
                                p.Remove();
                            }
                        }
                    }

                    if (wordDocument.MainDocumentPart.FooterParts.Count() > 0)
                    {
                        foreach (FooterPart footerPart in wordDocument.MainDocumentPart.FooterParts)
                        {
                            foreach (Paragraph p in footerPart.Footer.Descendants<Paragraph>().Where<Paragraph>(p => p.InnerText.Contains(WopiOptions.Value.ConversionEngine)))
                            {
                                p.Remove();
                            }
                        }
                    }

                    if (WopiOptions?.Value?.DocumentProtection == TRUE)
                    {
                        if (getIDPList().Contains("all") || getIDPList().Contains(userRole) || userRole == null)
                        {
                            ArrayList docProtClassList = new ArrayList();
                            ArrayList docProtTypeList = new ArrayList();

                            if (null != WopiOptions.Value.DocumentProtectionClass && WopiOptions.Value.DocumentProtectionClass.Length > 0)
                                docProtClassList.AddRange(WopiOptions.Value.DocumentProtectionClass);

                            if (null != WopiOptions.Value.DocumentProtectionType && WopiOptions.Value.DocumentProtectionType.Length > 0)
                                docProtTypeList.AddRange(WopiOptions.Value.DocumentProtectionType);


                            var dsp = wordDocument.MainDocumentPart.DocumentSettingsPart;
                            foreach (DocumentProtection dp in wordDocument.MainDocumentPart.DocumentSettingsPart.Settings.ChildElements.OfType<DocumentProtection>())
                            {
                                if (!(String.IsNullOrEmpty(WopiOptions.Value.DocumentProtectionFlag)))
                                {
                                    if (null != dp.Enforcement)
                                    {

                                        if (dp.Enforcement == new OnOffValue(true))
                                        {
                                            //setCustomProperty(wordDocument, WopiOptions.Value.DocumentProtectionFlag, WopiOptions.Value.DocumentProtectionClass + "," + WopiOptions.Value.DocumentProtectionType, CustomPropertyTypes.Text);
                                            string docProtString = null;
                                            string docProtClass = null;
                                            string docProtType = null;
                                            if (docProtClassList.Contains("edit") && null != dp.Edit)
                                            {
                                                docProtClass = "edit";

                                                /*//
                                                // Summary:
                                                //     No Editing Restrictions.
                                                //     When the item is serialized out as xml, its value is "none".
                                                None = 0,
                                                //
                                                // Summary:
                                                //     Allow No Editing.
                                                //     When the item is serialized out as xml, its value is "readOnly".
                                                ReadOnly = 1,
                                                //
                                                // Summary:
                                                //     Allow Editing of Comments.
                                                //     When the item is serialized out as xml, its value is "comments".
                                                Comments = 2,
                                                //
                                                // Summary:
                                                //     Allow Editing With Revision Tracking.
                                                //     When the item is serialized out as xml, its value is "trackedChanges".
                                                TrackedChanges = 3,
                                                //
                                                // Summary:
                                                //     Allow Editing of Form Fields.
                                                //     When the item is serialized out as xml, its value is "forms".
                                                Forms = 4 */
                                                if (dp.Edit == DocumentProtectionValues.ReadOnly) docProtType = "1";
                                                if (dp.Edit == DocumentProtectionValues.Comments) docProtType = "2";
                                                if (dp.Edit == DocumentProtectionValues.TrackedChanges) docProtType = "3";
                                                if (dp.Edit == DocumentProtectionValues.Forms) docProtType = "4";
                                                if (dp.Edit == DocumentProtectionValues.None) docProtType = "0";

                                                var dpnew = new DocumentProtection()
                                                {
                                                    Edit = dp.Edit,
                                                    Enforcement = new OnOffValue(false),
                                                    Formatting = dp.Formatting
                                                    //CryptographicProviderType = CryptProviderValues.RsaFull,
                                                    //CryptographicAlgorithmClass = CryptAlgorithmClassValues.Hash,
                                                    //CryptographicAlgorithmType = CryptAlgorithmValues.TypeAny,
                                                    //CryptographicAlgorithmSid = 4,
                                                    //CryptographicSpinCount = 100000U,
                                                    //Hash = "2krUoz1qWd0WBeXqVrOq81l8xpk=",
                                                    //Salt = "9kIgmDDYtt2r5U2idCOwMA=="
                                                };


                                                if (!(null == dsp || null == dsp.Settings))
                                                {
                                                    dsp.Settings.ReplaceChild(dpnew, dp);
                                                }
                                            }

                                            // handles other types of protection here


                                            setCustomProperty(wordDocument, WopiOptions.Value.DocumentProtectionFlag, docProtClass + "," + docProtType, CustomPropertyTypes.Text);

                                            var docProtKey = Path.GetFileNameWithoutExtension(newFileName);
                                            if (null != docProtKey && docProtection.ContainsKey(docProtKey))
                                            {
                                                docProtection[docProtKey] = TRUE;
                                            }
                                            else
                                            {
                                                docProtection.Add(docProtKey, TRUE);
                                            }


                                        }
                                    }
                                }
                                //dp.Remove();




                                /*
                                < w:documentProtection w:edit = "forms"
                                  w: formatting = "1"
                                  w: enforcement = "1" />
                                */

                            }
                        }
                    }


                    //const string FieldDelimeter = @" MERGEFIELD ";
                    string FieldDelimeter = @" DOCPROPERTY ";
                    List<string> listeChamps = new List<string>();


                    if (normalize)
                    {
                        normalizeMarkup(wordDocument);
                        normalizeFieldCodesRuns(wordDocument);
                    }

                    Run prevRun = null;
                    Run prevBegin = null;
                    Run prevEnd = null;

                    string[] nonEditable = WopiOptions.Value.BookmarksNonEditable;

                    /* new string[] {
                      "LetterDate",
                     "NameAndAddressLn1",
                     "NameAndAddressLn2",
                     "NameAndAddressLn3",
                     "NameAndAddressLn4",
                     "NameAndAddressLn5",
                     "NameAndAddressLn6",
                     "NameAndAddressLn7",
                     "(P)",
                     "\"CC\"",
                     " CC ",
                     "COPYIND",
                     "CCTOKENLIST",
                     "IPTOKENLIST",
                     "PRIMARYRECIPCONTACTNAME" }; */



                    ArrayList nonEditableArray = new ArrayList();
                    for (var k = 0; k < nonEditable.Length; k++)
                    {
                        nonEditableArray.Add(nonEditable[k].ToUpper());
                    }

                    //nonEditableArray.AddRange(nonEditable);



                    int j = 0;
                    var fieldList = wordDocument.MainDocumentPart.RootElement.Descendants<FieldCode>();
                    foreach (var field2 in fieldList)
                    {
                        j++;
                        Console.Out.WriteLine("***Field " + j.ToString() + ">>>" + "Field Type:" + field2.ToString() + "<<<Field Text:>>>" + field2.Text.ToString() + "<<<");
                        if (null != field2.InnerXml) Console.Out.WriteLine("***** internal XML:" + field2.InnerXml + " *****");
                    }

                    j = 0;

                    foreach (var field in fieldList)
                    {
                        j++;
                        Console.Out.WriteLine("*** Field " + j.ToString() + " starts *********************************");
                        Console.Out.WriteLine("Field Type:" + field.ToString());
                        Console.Out.WriteLine("Field Text:>>>" + field.Text.ToString() + "<<<");
                        bool isFormField = false;
                        var fieldId = field.Text.Trim().ToString();
                        if (null != fieldId)
                        {
                            if (fieldId.Contains("FORMTEXT"))
                            {
                                isFormField = true;
                            }
                            Console.Out.WriteLine("fieldId:" + fieldId.ToString());
                        }

                        /*if (field.Text.ToString() == " DOCPROPERTY  WorkerName  \\* CHARFORMAT ")
                         {
                             field.Text = new string("REPLACED WORKER NAME");


                             Console.Out.WriteLine("Replaceing Field Text:>>>" + field.Text.ToString() + "<<<");
                             Console.Out.WriteLine("with:>>>REPLACED_WORKER_NAME<<<");
                         }*/


                        Run xxxfield = null;
                        Run rBegin = null;
                        Run rSep = null;
                        Run rText = null;
                        Run rEnd = null;
                        Text t = null;

                        Run nFormat = null;
                        Run nBegin = null;
                        Run nTag = null;
                        Run nSep = null;
                        Run nText = null;
                        Run nEnd = null;
                        Text nt = null;

                        Run nnFormat = null;
                        Run nnBegin = null;
                        Run nnTag = null;
                        Run nnSep = null;
                        Run nnText = null;
                        Run nnEnd = null;
                        Text nnt = null;


                        Run pivotRun = null;
                        Run pivotEnd = null;
                        Run pivotBegin = null;
                        Run rAdd = null;

                        xxxfield = (Run)field.Parent;
                        if (null != xxxfield) rBegin = xxxfield.PreviousSibling<Run>();
                        if (null != xxxfield) rSep = xxxfield.NextSibling<Run>();
                        if (null != rSep) rText = rSep.NextSibling<Run>();
                        if (null != rText) rEnd = rText.NextSibling<Run>();
                        if (null != rText) t = rText.GetFirstChild<Text>();

                        pivotRun = xxxfield;
                        pivotBegin = rBegin;
                        pivotEnd = rEnd;
                        bool found = false;

                        Console.Out.WriteLine("@@@ Checking field Id...");
                        if (null != xxxfield)
                        {
                            if (null != xxxfield.InnerText) Console.Out.WriteLine("xxxfield:" + xxxfield.ToString() + ">>>" + xxxfield.InnerText);
                        }
                        if (null != rBegin)
                        {
                            if (null != rBegin.InnerText) Console.Out.WriteLine("rBegin:" + rBegin.ToString() + ">>>" + rBegin.InnerText);
                        }
                        if (null != rText)
                        {
                            if (null != rText.InnerText) Console.Out.WriteLine("rText:" + rText.ToString() + ">>>" + rText.InnerText);
                        }
                        if (null != rEnd)
                        {
                            if (null != rEnd.InnerText) Console.Out.WriteLine("rEnd:" + rEnd.ToString() + ">>>" + rEnd.InnerText);
                        }
                        if (null != t)
                        {
                            if (null != t.InnerText) Console.Out.WriteLine("t:" + t.ToString() + ">>>" + t.InnerText);
                        }


                        if (isFormField)
                        {
                            if (null != t)
                            {
                                if (null != t.Text)
                                {
                                    var replaceFormValue = "";
                                    Console.Out.WriteLine("****Substitute form field " + t.Text);
                                    if (String.IsNullOrWhiteSpace(t.Text.Trim()))
                                        replaceFormValue = "[ ]";
                                    else
                                        replaceFormValue = "[" + t.Text.Trim() + "]";
                                    t.Text = formatText(replaceFormValue);
                                    rEnd.Remove();
                                    rSep.Remove();
                                    xxxfield.Remove();
                                    rBegin.Remove();

                                    if (null != rText)
                                    {
                                        var rp = rText.RunProperties;
                                        Highlight highlight = new Highlight() { Val = HighlightColorValues.Yellow };
                                        rp.Append(highlight);
                                    }


                                }
                            }
                        }
                        else
                        {
                            found = true;
                            if (found)
                            {

                                Regex expr = new Regex(@"\s*(?<docProperty>\S+)\s+(?<aFieldName>\S+)\.*\s+(?<formatType>\S+)\s*");
                                var results = expr.Matches(fieldId);
                                string docProperty = null;
                                string aFieldName = null;
                                string formatType = null;

                                foreach (Match match in results)
                                {
                                    docProperty = match.Groups["docProperty"].Value;
                                    aFieldName = match.Groups["aFieldName"].Value;
                                    formatType = match.Groups["formatType"].Value;
                                }

                                if (null != docProperty) Console.Out.WriteLine("docProperty=" + docProperty);
                                if (null != aFieldName) Console.Out.WriteLine("aFieldName=" + aFieldName);
                                if (null != formatType) Console.Out.WriteLine("formatType=" + formatType);

                                if (null != aFieldName)
                                {
                                    var aFieldKey = aFieldName.Trim().ToUpper();
                                    string aFieldValue = null;
                                    if (fieldMap.ContainsKey(aFieldKey))
                                    {
                                        if (null != fieldMap[aFieldKey])
                                        {
                                            aFieldValue = fieldMap[aFieldKey].Trim();
                                        }
                                    }
                                    if (String.IsNullOrEmpty(aFieldValue)) aFieldValue = new string(" ");

                                    if (t != null)
                                    {
                                        if (t.Text != null && aFieldKey != null && fieldMap.ContainsKey(aFieldKey))
                                        {
                                            if (!fieldMap.ContainsKey(aFieldKey) || (String.IsNullOrWhiteSpace(aFieldValue)
                                            || aFieldValue.Contains("BOOKMARK_UNDEFINED")))
                                            {
                                                Run rBegin2 = null;
                                                Run rBegin1 = null;
                                                Run rBegin0 = null;
                                                Paragraph rParent = null;
                                                Run rParentFirst = null;
                                                Run rParentLeft = null;
                                                Run rParentLeftFirst = null;
                                                if (null != rBegin) rBegin2 = rBegin.PreviousSibling<Run>();
                                                if (null != rBegin2) rBegin1 = rBegin2.PreviousSibling<Run>();
                                                if (null != rBegin1) rBegin0 = rBegin1.PreviousSibling<Run>();
                                                if (null != rText) rParent = (Paragraph)rText.Parent;
                                                /*if (null != rParent) rParentFirst = (Run)rParent.GetFirstChild<Run>();
                                                if (null != rParent) rParentLeft = rParent.PreviousSibling<Run>();
                                                if (null != rParentLeft) rParentLeftFirst = rParentLeft.GetFirstChild<Run>();*/
                                                //t.Text = " ";
                                                //if (null != rEnd) rAdd = rEnd.NextSibling<Run>();
                                                //if (null != rAdd)
                                                //    rAdd.AppendChild(new Text(" "));
                                                //rAdd.AppendChild(new Text(fieldMap[aFieldKey]));*/
                                                //if (null != rParent) rParent.InnerXml.Replace(fieldId, "");
                                                //if (null != rText) rText.RemoveAllChildren();
                                                //if (null != rText) rText.Remove();
                                                if (null != rEnd) rEnd.Remove();
                                                if (null != rSep) rSep.Remove();
                                                if (null != rBegin) rBegin.Remove();
                                                if (null != t) t.Text = aFieldValue;
                                                if (null != xxxfield) xxxfield.Remove();
                                                //if (null != rParent) rParent.RemoveAllChildren();
                                                //if (null != rParent) rParent.AppendChild<Run>(new Run(new Text("")));                                        
                                                /*if (null != rBegin2) rBegin2.RemoveAllChildren();
                                                if (null != rBegin2) rBegin2.Remove();
                                                if (null != rBegin1) rBegin1.RemoveAllChildren();
                                                if (null != rBegin1) rBegin1.Remove();
                                                if (null != rBegin0) rBegin0.RemoveAllChildren();
                                                if (null != rBegin0) rBegin0.Remove();*/
                                                //if (null != t.Text) Console.Out.WriteLine("****Substitute value " + t.Text + "with " + fieldMap[aFieldKey]);
                                                //rText.Remove();
                                            }
                                            else
                                            {
                                                Console.Out.WriteLine("****Substitute value " + t.Text + "with " + aFieldValue);
                                                if (fieldMap.ContainsKey(aFieldKey) && !(String.IsNullOrEmpty(aFieldValue)))
                                                {



                                                    if (nonEditableArray.Contains(aFieldKey))
                                                    {
                                                        t.Text = formatText(aFieldValue);
                                                    }
                                                    else
                                                    {


                                                        Run rBegin2 = null;
                                                        Run rBegin1 = null;
                                                        Run rBegin0 = null;
                                                        Paragraph rParent = null;
                                                        Run rParentFirst = null;
                                                        Run rParentLeft = null;
                                                        Run rParentLeftFirst = null;
                                                        if (null != rBegin) rBegin2 = rBegin.PreviousSibling<Run>();
                                                        if (null != rBegin2) rBegin1 = rBegin2.PreviousSibling<Run>();
                                                        if (null != rBegin1) rBegin0 = rBegin1.PreviousSibling<Run>();
                                                        if (null != rText) rParent = (Paragraph)rText.Parent;
                                                        /*if (null != rParent) rParentFirst = (Run)rParent.GetFirstChild<Run>();
                                                        if (null != rParent) rParentLeft = rParent.PreviousSibling<Run>();
                                                        if (null != rParentLeft) rParentLeftFirst = rParentLeft.GetFirstChild<Run>();*/
                                                        //t.Text = fieldMap[aFieldKey];
                                                        //if (null != rEnd) rAdd = rEnd.NextSibling<Run>();
                                                        //if (null != rAdd)
                                                        //    rAdd.AppendChild(new Text(fieldMap[aFieldKey]));
                                                        //rAdd = rText.NextSibling<Run>();
                                                        /*if (null != rAdd)
                                                            rAdd.AppendChild(new Text(fieldMap[aFieldKey]));*/
                                                        //if (null != rParent) rParent.InnerXml.Replace(fieldId, "");
                                                        //if (null != rText) rText.RemoveAllChildren();
                                                        //if (null != rText) rText.Remove();

                                                        string aaFieldKey = null;
                                                        string aaFieldValue = null;
                                                        string aaaFieldKey = null;
                                                        string aaaFieldValue = null;

                                                        if (null != rEnd) nFormat = rEnd.NextSibling<Run>();
                                                        if (null != nFormat) nBegin = nFormat.NextSibling<Run>();
                                                        if (null != nBegin) nTag = nBegin.NextSibling<Run>();
                                                        if (null != nTag) nSep = nTag.NextSibling<Run>();
                                                        if (null != nSep) nText = nSep.NextSibling<Run>();
                                                        if (null != nText) nEnd = nText.NextSibling<Run>();
                                                        if (null != nText) nt = nText.GetFirstChild<Text>();
                                                        if (null != nEnd) nnFormat = nEnd.NextSibling<Run>();
                                                        if (null != nnFormat) nnBegin = nnFormat.NextSibling<Run>();
                                                        if (null != nnBegin) nnTag = nnBegin.NextSibling<Run>();
                                                        if (null != nnTag) nnSep = nnTag.NextSibling<Run>();
                                                        if (null != nnSep) nnText = nnSep.NextSibling<Run>();
                                                        if (null != nnText) nnEnd = nnText.NextSibling<Run>();
                                                        if (null != nnText) nnt = nnText.GetFirstChild<Text>();


                                                        if (null != nnt)
                                                        {
                                                            if (null != nnt.Text && !String.IsNullOrWhiteSpace(nnt.Text))
                                                            {
                                                                aaaFieldKey = nnt.Text.Trim().ToUpper();
                                                                if (null != aaaFieldKey && fieldMap.ContainsKey(aaaFieldKey)) aaaFieldValue = fieldMap[aaaFieldKey];
                                                                if (null != aaaFieldValue && !String.IsNullOrWhiteSpace(aaaFieldValue) && aaaFieldValue.ToUpper() != "BOOKMARK_UNDEFINED")
                                                                {
                                                                    nnt.Text = formatText(aaaFieldValue);
                                                                    if (null != nnEnd) nnEnd.Remove();
                                                                    if (null != nnSep) nnSep.Remove();
                                                                    if (null != nnTag) nnTag.Remove();
                                                                    if (null != nnBegin) nnBegin.Remove();
                                                                    //if (null != rText) rText = new Run(new Text(aFieldValue));

                                                                }
                                                            }
                                                        }

                                                        if (null != nt)
                                                        {
                                                            if (null != nt.Text && !String.IsNullOrWhiteSpace(nt.Text))
                                                            {
                                                                aaFieldKey = nt.Text.Trim().ToUpper();
                                                                if (null != aaFieldKey && fieldMap.ContainsKey(aaFieldKey)) aaFieldValue = fieldMap[aaFieldKey];
                                                                if (null != aaFieldValue && !String.IsNullOrWhiteSpace(aaFieldValue) && aaFieldValue.ToUpper() != "BOOKMARK_UNDEFINED")
                                                                {
                                                                    nt.Text = formatText(aaFieldValue);
                                                                    if (null != nEnd) nEnd.Remove();
                                                                    if (null != nSep) nSep.Remove();
                                                                    if (null != nTag) nTag.Remove();
                                                                    if (null != nBegin) nBegin.Remove();
                                                                }
                                                            }
                                                        }







                                                        if (null != rEnd) rEnd.Remove();
                                                        if (null != rSep) rSep.Remove();
                                                        if (null != rBegin) rBegin.Remove();
                                                        if (null != t && null != t.Text) t.Text = formatText(aFieldValue);
                                                        //if (null != rText) rText = new Run(new Text(aFieldValue));
                                                        if (null != xxxfield) xxxfield.Remove();
                                                        //if (null != rParent) rParent.RemoveAllChildren();
                                                        //if (null != rParent) rParent.AppendChild<Run>(new Run(new Text(aFieldValue)));
                                                        /*if (null != rBegin2) rBegin2.RemoveAllChildren();
                                                        if (null != rBegin2) rBegin2.Remove();
                                                        if (null != rBegin1) rBegin1.RemoveAllChildren();
                                                        if (null != rBegin1) rBegin1.Remove();
                                                        if (null != rBegin0) rBegin0.RemoveAllChildren();
                                                        if (null != rBegin0) rBegin0.Remove();*/


                                                    }
                                                }
                                            }

                                        }
                                    }
                                    else //field name is CHARFORMAT or something
                                    {
                                        if (null != rEnd)
                                        {

                                            if (null != rEnd.InnerText)
                                            {
                                                if (!(String.IsNullOrWhiteSpace(rEnd.InnerText)) && fieldMap.ContainsKey(rEnd.InnerText.ToUpper()))
                                                {
                                                    //rEnd.SetText(fieldMap[rEnd.InnerText]);
                                                    if (null != fieldMap[rEnd.InnerText.ToUpper()])
                                                        rEnd.SetText(fieldMap[rEnd.InnerText.ToUpper()]);
                                                    else
                                                        rEnd.SetText(formatText(null));
                                                    /*Run rBegin2 = null;
                                                    Run rBegin1 = null;
                                                    Run rBegin0 = null;
                                                    Run rParent = null;
                                                    Run rParentFirst = null;
                                                    Run rParentLeft = null;
                                                    Run rParentLeftFirst = null;                                        
                                                    if (null != rBegin) rBegin2 = rBegin.PreviousSibling<Run>();
                                                    if (null != rBegin2) rBegin1 = rBegin2.PreviousSibling<Run>();
                                                    if (null != rBegin1) rBegin0 = rBegin1.PreviousSibling<Run>();
                                                    if (null != rEnd) rParent = (Run)rEnd.Parent;
                                                    if (null != rParent) rParentFirst = (Run)rParent.GetFirstChild<Run>();
                                                    if (null != rParent) rParentLeft = rParent.PreviousSibling<Run>();
                                                    if (null != rParentLeft) rParentLeftFirst = rParentLeft.GetFirstChild<Run>();
                                                    rAdd = rEnd.NextSibling<Run>();
                                                    if (null != rAdd)
                                                        rAdd.AppendChild(new Text(fieldMap[rEnd.InnerText]));
                                                    if (null != rEnd) rEnd.RemoveAllChildren();
                                                    if (null != rEnd) rEnd.Remove();
                                                    if (null != rText) rText.RemoveAllChildren();
                                                    if (null != rText) rText.Remove();
                                                    if (null != rSep) rSep.RemoveAllChildren();
                                                    if (null != rSep) rSep.Remove();
                                                    if (null != rBegin) rBegin.RemoveAllChildren();
                                                    if (null != rBegin) rBegin.Remove();
                                                    if (null != rBegin2) rBegin2.RemoveAllChildren();
                                                    if (null != rBegin2) rBegin2.Remove();
                                                    if (null != rBegin1) rBegin1.RemoveAllChildren();
                                                    if (null != rBegin1) rBegin1.Remove();
                                                    if (null != rBegin0) rBegin0.RemoveAllChildren();
                                                    if (null != rBegin0) rBegin0.Remove();
                                                    if (null != rParentFirst) rParentFirst.RemoveAllChildren();
                                                    if (null != rParentFirst) rParentFirst.Remove();
                                                    if (null != rParentLeftFirst) rParentLeftFirst.RemoveAllChildren();
                                                    if (null != rParentLeftFirst) rParentLeftFirst.Remove();
                                                    if (null != xxxfield) xxxfield.RemoveAllChildren();
                                                    if (null != xxxfield) xxxfield.Remove();*/
                                                }
                                                else
                                                {
                                                    rEnd.SetText(formatText(null));
                                                }
                                            }


                                        }
                                        else
                                        {
                                            // rEnd = nuill rEnd.SetText(" ");
                                            /* if (rEnd != null)
                                             {
                                                 Run rObject1 = null;
                                                 Run rObject2 = null;
                                                 Run rObject3 = null;
                                                 rObject1 = rEnd.GetFirstChild<Run>();
                                                 if (null != rObject1) rObject2 = rObject1.NextSibling<Run>();
                                                 if (null != rObject2) rObject3 = rObject2.NextSibling<Run>();

                                                 if (null != rObject1) 
                                                 {
                                                     if (null != rObject1 && null != rObject1.GetTextElement)
                                                         rObject1.Remove();
                                                     if (null != rObject2 && null != rObject2.GetTextElement)
                                                         rObject2.Remove();
                                                     if (null != rObject3)
                                                 }
                                             }*/
                                            /*t.Text = "";
                                            Run rBegin2 = null;
                                            Run rBegin1 = null;
                                            Run rBegin0 = null;
                                            Run rParent = null;
                                            Run rParentFirst = null;
                                            Run rParentLeft = null;
                                            Run rParentLeftFirst = null;
                                            if (null != rBegin) rBegin2 = rBegin.PreviousSibling<Run>();
                                            if (null != rBegin2) rBegin1 = rBegin2.PreviousSibling<Run>();
                                            if (null != rBegin1) rBegin0 = rBegin1.PreviousSibling<Run>();
                                            if (null != rEnd) rParent = (Run)rEnd.Parent;
                                            if (null != rParent) rParentFirst = (Run)rParent.GetFirstChild<Run>();
                                            if (null != rParent) rParentLeft = rParent.PreviousSibling<Run>();
                                            if (null != rParentLeft) rParentLeftFirst = rParentLeft.GetFirstChild<Run>();
                                            rAdd = rEnd.NextSibling<Run>();                                        
                                            if (null != rAdd) rAdd.AppendChild(new Text(""));
                                            if (null != rEnd) rEnd.RemoveAllChildren();
                                            if (null != rEnd) rEnd.Remove();                                        
                                            if (null != rText) rText.RemoveAllChildren();
                                            if (null != rText) rText.Remove();
                                            if (null != rSep) rSep.RemoveAllChildren();
                                            if (null != rSep) rSep.Remove();
                                            if (null != rBegin) rBegin.RemoveAllChildren();
                                            if (null != rBegin) rBegin.Remove();
                                            if (null != rBegin2) rBegin2.RemoveAllChildren();
                                            if (null != rBegin2) rBegin2.Remove();
                                            if (null != rBegin1) rBegin1.RemoveAllChildren(); 
                                            if (null != rBegin1) rBegin1.Remove();
                                            if (null != rBegin0) rBegin0.RemoveAllChildren(); 
                                            if (null != rBegin0) rBegin0.Remove();
                                            if (null != rParentFirst) rParentFirst.RemoveAllChildren();
                                            if (null != rParentFirst) rParentFirst.Remove();
                                            if (null != rParentLeftFirst) rParentLeftFirst.RemoveAllChildren();
                                            if (null != rParentLeftFirst) rParentLeftFirst.Remove();
                                            if (null != xxxfield) xxxfield.RemoveAllChildren();
                                            if (null != xxxfield) xxxfield.Remove();*/

                                        }
                                    }
                                }
                                else
                                {

                                }
                            }
                            else
                            {
                                Console.Out.WriteLine("@@@ field value not found.");
                            }
                        }


                        Console.Out.WriteLine("*** Field " + j.ToString() + " ends *********************************");





                        /* DocumentProperty cField = (custom[fieldId]);
                         if (null != cField)
                         {
                             Console.Out.WriteLine(">>> " + cField.Name + ": " + cField.Value);
                         }*/




                        /*int fieldNameStart = field.Text.LastIndexOf(FieldDelimeter, System.StringComparison.Ordinal);

                        if (fieldNameStart >= 0)
                        {
                            var fieldName = field.Text.Substring(fieldNameStart + FieldDelimeter.Length).Trim();

                            Run xxxfield = (Run)field.Parent;

                            Run rBegin = xxxfield.PreviousSibling<Run>();
                            Run rSep = xxxfield.NextSibling<Run>();
                            Run rText = rSep.NextSibling<Run>();
                            Run rEnd = rText.NextSibling<Run>();

                            if (null != xxxfield)
                            {

                                Text t = rText.GetFirstChild<Text>();
                               //custom.CustomHash;
                               Console.Out.WriteLine(t.ToString());

                            }
                        }*/

                    }




                    setCustomProperty(wordDocument, WopiOptions.Value.ProcessedFlag, WopiOptions.Value.ApplicationName, CustomPropertyTypes.Text);

                    /*Remove VBA part
                        var docPart = wordDocument.MainDocumentPart;

                        // Look for the vbaProject part. If it is there, delete it.
                        var vbaPart = docPart.VbaProjectPart;
                        if (vbaPart != null)
                        {
                            // Delete the vbaProject part and then save the document.
                            docPart.DeletePart(vbaPart);
                            docPart.Document.Save();

                            // Change the document type to
                            // not macro-enabled.
                            wordDocument.ChangeDocumentType(WordprocessingDocumentType.Document);

                            // Track that the document has been changed.

                        }
                        //changeCompatibilityModeOfDocumentPart(wordDocument.MainDocumentPart);
                    */

                    wordDocument.Save();
                    wordDocument.Close();

                }
            } // end using

        }


        public void runMacroX(string newFileName, Dictionary<string, string> fieldMap, bool processedByWWA = false, bool normalize = false, string userRole = null)
        {
            Regex fileVer = new Regex(@"\s*(?<userID>)_(?<attachmentID>)_(?<versionNo>)\..*");
            var result = fileVer.Matches(newFileName);
            string userID = null;
            string attachmentID = null;
            string versionNo = null;

            /*if (Int32.TryParse(versionNo, out int v))
            {
                if (v > 1) return;
            }*/
            //if (processedByWWA) return;

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(newFileName, true))
            {

                if (null != wordDocument)
                {
                    /*for (var field in wordDocument.MainDocumentPart.RootElement.Descendants<FieldCode>())
                    {

                    }*/

                    /*List<Picture> pictures = new List<Picture>(wordDocument.MainDocumentPart.RootElement.Descendants<Picture>());

                    foreach (Picture p in pictures)
                    {
                        p.Remove();
                    }*/

                    //headerPart.DeleteParts(imagePartList);

                    foreach (Paragraph p in wordDocument.MainDocumentPart.Document.Body.Descendants<Paragraph>().Where<Paragraph>(p => p.InnerText.Contains(WopiOptions.Value.ConversionEngine)))
                    {
                        // Do something with the Paragraphs.
                        p.Remove();
                    }

                    if (wordDocument.MainDocumentPart.HeaderParts.Count() > 0)
                    {
                        foreach (HeaderPart headerPart in wordDocument.MainDocumentPart.HeaderParts)
                        {
                            foreach (Paragraph p in headerPart.Header.Descendants<Paragraph>().Where<Paragraph>(p => p.InnerText.Contains(WopiOptions.Value.ConversionEngine)))
                            {
                                p.Remove();
                            }
                        }
                    }

                    if (wordDocument.MainDocumentPart.FooterParts.Count() > 0)
                    {
                        foreach (FooterPart footerPart in wordDocument.MainDocumentPart.FooterParts)
                        {
                            foreach (Paragraph p in footerPart.Footer.Descendants<Paragraph>().Where<Paragraph>(p => p.InnerText.Contains(WopiOptions.Value.ConversionEngine)))
                            {
                                p.Remove();
                            }
                        }
                    }

                    if (WopiOptions?.Value?.DocumentProtection == TRUE)
                    {
                        if (getIDPList().Contains("all") || getIDPList().Contains(userRole) || userRole == null)
                        {
                            ArrayList docProtClassList = new ArrayList();
                            ArrayList docProtTypeList = new ArrayList();

                            if (null != WopiOptions.Value.DocumentProtectionClass && WopiOptions.Value.DocumentProtectionClass.Length > 0)
                                docProtClassList.AddRange(WopiOptions.Value.DocumentProtectionClass);

                            if (null != WopiOptions.Value.DocumentProtectionType && WopiOptions.Value.DocumentProtectionType.Length > 0)
                                docProtTypeList.AddRange(WopiOptions.Value.DocumentProtectionType);


                            var dsp = wordDocument.MainDocumentPart.DocumentSettingsPart;
                            foreach (DocumentProtection dp in wordDocument.MainDocumentPart.DocumentSettingsPart.Settings.ChildElements.OfType<DocumentProtection>())
                            {
                                if (!(String.IsNullOrEmpty(WopiOptions.Value.DocumentProtectionFlag)))
                                {
                                    if (null != dp.Enforcement)
                                    {

                                        if (dp.Enforcement == new OnOffValue(true))
                                        {
                                            //setCustomProperty(wordDocument, WopiOptions.Value.DocumentProtectionFlag, WopiOptions.Value.DocumentProtectionClass + "," + WopiOptions.Value.DocumentProtectionType, CustomPropertyTypes.Text);
                                            string docProtString = null;
                                            string docProtClass = null;
                                            string docProtType = null;
                                            if (docProtClassList.Contains("edit") && null != dp.Edit)
                                            {
                                                docProtClass = "edit";

                                                /*//
                                                // Summary:
                                                //     No Editing Restrictions.
                                                //     When the item is serialized out as xml, its value is "none".
                                                None = 0,
                                                //
                                                // Summary:
                                                //     Allow No Editing.
                                                //     When the item is serialized out as xml, its value is "readOnly".
                                                ReadOnly = 1,
                                                //
                                                // Summary:
                                                //     Allow Editing of Comments.
                                                //     When the item is serialized out as xml, its value is "comments".
                                                Comments = 2,
                                                //
                                                // Summary:
                                                //     Allow Editing With Revision Tracking.
                                                //     When the item is serialized out as xml, its value is "trackedChanges".
                                                TrackedChanges = 3,
                                                //
                                                // Summary:
                                                //     Allow Editing of Form Fields.
                                                //     When the item is serialized out as xml, its value is "forms".
                                                Forms = 4 */
                                                if (dp.Edit == DocumentProtectionValues.ReadOnly) docProtType = "1";
                                                if (dp.Edit == DocumentProtectionValues.Comments) docProtType = "2";
                                                if (dp.Edit == DocumentProtectionValues.TrackedChanges) docProtType = "3";
                                                if (dp.Edit == DocumentProtectionValues.Forms) docProtType = "4";
                                                if (dp.Edit == DocumentProtectionValues.None) docProtType = "0";

                                                var dpnew = new DocumentProtection()
                                                {
                                                    Edit = dp.Edit,
                                                    Enforcement = new OnOffValue(false),
                                                    Formatting = dp.Formatting
                                                    //CryptographicProviderType = CryptProviderValues.RsaFull,
                                                    //CryptographicAlgorithmClass = CryptAlgorithmClassValues.Hash,
                                                    //CryptographicAlgorithmType = CryptAlgorithmValues.TypeAny,
                                                    //CryptographicAlgorithmSid = 4,
                                                    //CryptographicSpinCount = 100000U,
                                                    //Hash = "2krUoz1qWd0WBeXqVrOq81l8xpk=",
                                                    //Salt = "9kIgmDDYtt2r5U2idCOwMA=="
                                                };


                                                if (!(null == dsp || null == dsp.Settings))
                                                {
                                                    dsp.Settings.ReplaceChild(dpnew, dp);
                                                }
                                            }

                                            // handles other types of protection here


                                            setCustomProperty(wordDocument, WopiOptions.Value.DocumentProtectionFlag, docProtClass + "," + docProtType, CustomPropertyTypes.Text);

                                            var docProtKey = Path.GetFileNameWithoutExtension(newFileName);
                                            if (null != docProtKey && docProtection.ContainsKey(docProtKey))
                                            {
                                                docProtection[docProtKey] = TRUE;
                                            }
                                            else
                                            {
                                                docProtection.Add(docProtKey, TRUE);
                                            }


                                        }
                                    }
                                }
                                //dp.Remove();




                                /*
                                < w:documentProtection w:edit = "forms"
                                  w: formatting = "1"
                                  w: enforcement = "1" />
                                */

                            }
                        }
                    }


                    //const string FieldDelimeter = @" MERGEFIELD ";
                    string FieldDelimeter = @" DOCPROPERTY ";
                    List<string> listeChamps = new List<string>();


                    if (normalize)
                    {
                        normalizeMarkup(wordDocument);
                        normalizeFieldCodesRuns(wordDocument);
                    }

                    Run prevRun = null;
                    Run prevBegin = null;
                    Run prevEnd = null;

                    string[] nonEditable = WopiOptions.Value.BookmarksNonEditable;

                    /* new string[] {
                      "LetterDate",
                     "NameAndAddressLn1",
                     "NameAndAddressLn2",
                     "NameAndAddressLn3",
                     "NameAndAddressLn4",
                     "NameAndAddressLn5",
                     "NameAndAddressLn6",
                     "NameAndAddressLn7",
                     "(P)",
                     "\"CC\"",
                     " CC ",
                     "COPYIND",
                     "CCTOKENLIST",
                     "IPTOKENLIST",
                     "PRIMARYRECIPCONTACTNAME" }; */



                    ArrayList nonEditableArray = new ArrayList();
                    for (var k = 0; k < nonEditable.Length; k++)
                    {
                        nonEditableArray.Add(nonEditable[k].ToUpper());
                    }

                    //nonEditableArray.AddRange(nonEditable);



                    int j = 0;
                    foreach (var field in wordDocument.MainDocumentPart.RootElement.Descendants<FieldCode>())
                    {
                        j++;
                        Console.Out.WriteLine("***Field " + j.ToString() + ">>>" + "Field Type:" + field.ToString() + "<<<Field Text:>>>" + field.Text.ToString() + "<<<");
                        if (null != field.InnerXml) Console.Out.WriteLine("***** internal XML:" + field.InnerXml + " *****");
                    }

                    j = 0;

                    foreach (var field in wordDocument.MainDocumentPart.RootElement.Descendants<FieldCode>())
                    {
                        j++;
                        Console.Out.WriteLine("*** Field " + j.ToString() + " starts *********************************");
                        Console.Out.WriteLine("Field Type:" + field.ToString());
                        Console.Out.WriteLine("Field Text:>>>" + field.Text.ToString() + "<<<");
                        bool isFormField = false;
                        var fieldId = field.Text.Trim().ToString();
                        if (null != fieldId)
                        {
                            if (fieldId.Contains("FORMTEXT"))
                            {
                                isFormField = true;
                            }
                            Console.Out.WriteLine("fieldId:" + fieldId.ToString());
                        }

                        /*if (field.Text.ToString() == " DOCPROPERTY  WorkerName  \\* CHARFORMAT ")
                         {
                             field.Text = new string("REPLACED WORKER NAME");


                             Console.Out.WriteLine("Replaceing Field Text:>>>" + field.Text.ToString() + "<<<");
                             Console.Out.WriteLine("with:>>>REPLACED_WORKER_NAME<<<");
                         }*/


                        Run xxxfield = null;
                        Run rBegin = null;
                        Run rSep = null;
                        Run rText = null;
                        Run rEnd = null;
                        Text t = null;

                        Run nFormat = null;
                        Run nBegin = null;
                        Run nTag = null;
                        Run nSep = null;
                        Run nText = null;
                        Run nEnd = null;
                        Text nt = null;

                        Run nnFormat = null;
                        Run nnBegin = null;
                        Run nnTag = null;
                        Run nnSep = null;
                        Run nnText = null;
                        Run nnEnd = null;
                        Text nnt = null;


                        Run pivotRun = null;
                        Run pivotEnd = null;
                        Run pivotBegin = null;
                        Run rAdd = null;

                        xxxfield = (Run)field.Parent;
                        if (null != xxxfield) rBegin = xxxfield.PreviousSibling<Run>();
                        if (null != xxxfield) rSep = xxxfield.NextSibling<Run>();
                        if (null != rSep) rText = rSep.NextSibling<Run>();
                        if (null != rText) rEnd = rText.NextSibling<Run>();
                        if (null != rText) t = rText.GetFirstChild<Text>();

                        pivotRun = xxxfield;
                        pivotBegin = rBegin;
                        pivotEnd = rEnd;
                        bool found = false;

                        Console.Out.WriteLine("@@@ Checking field Id...");
                        if (null != xxxfield)
                        {
                            if (null != xxxfield.InnerText) Console.Out.WriteLine("xxxfield:" + xxxfield.ToString() + ">>>" + xxxfield.InnerText);
                        }
                        if (null != rBegin)
                        {
                            if (null != rBegin.InnerText) Console.Out.WriteLine("rBegin:" + rBegin.ToString() + ">>>" + rBegin.InnerText);
                        }
                        if (null != rText)
                        {
                            if (null != rText.InnerText) Console.Out.WriteLine("rText:" + rText.ToString() + ">>>" + rText.InnerText);
                        }
                        if (null != rEnd)
                        {
                            if (null != rEnd.InnerText) Console.Out.WriteLine("rEnd:" + rEnd.ToString() + ">>>" + rEnd.InnerText);
                        }
                        if (null != t)
                        {
                            if (null != t.InnerText) Console.Out.WriteLine("t:" + t.ToString() + ">>>" + t.InnerText);
                        }



                        //t.SetText("Vincent");

                        if (!isFormField && (null == t || null == t.InnerText || String.IsNullOrWhiteSpace(t.InnerText)))
                        //if (!fieldId.Contains("DOCPROPERTY"))
                        {
                            // check the previous sibling to see if it is there
                            Console.Out.WriteLine("@@@ Checking previous sibling...");
                            xxxfield = rBegin;
                            if (null != xxxfield) rBegin = xxxfield.PreviousSibling<Run>();
                            if (null != xxxfield) rSep = xxxfield.NextSibling<Run>();
                            if (null != rSep) rText = rSep.NextSibling<Run>();
                            if (null != rText) rEnd = rText.NextSibling<Run>();
                            if (null != rText) t = rText.GetFirstChild<Text>();
                            if (null != t)
                            {
                                if (null != t.Text && String.IsNullOrWhiteSpace(t.Text))
                                {
                                    Console.Out.WriteLine("@@@ Not found in previous sibling...");
                                }
                                else
                                {
                                    if (null != t.Text && t.Text.Length > 0)
                                    {
                                        Console.Out.WriteLine("@@@ Found in previous sibling... >>>" + t.Text + "<<<");
                                        if (null != xxxfield)
                                        {
                                            if (null != xxxfield.InnerText) Console.Out.WriteLine("xxxfield:" + xxxfield.ToString() + ">>>" + xxxfield.InnerText);
                                        }
                                        if (null != rBegin)
                                        {
                                            if (null != rBegin.InnerText) Console.Out.WriteLine("rBegin:" + rBegin.ToString() + ">>>" + rBegin.InnerText);
                                        }
                                        if (null != rText)
                                        {
                                            if (null != rText.InnerText) Console.Out.WriteLine("rText:" + rText.ToString() + ">>>" + rText.InnerText);
                                        }
                                        if (null != rEnd)
                                        {
                                            if (null != rEnd.InnerText) Console.Out.WriteLine("rEnd:" + rEnd.ToString() + ">>>" + rEnd.InnerText);
                                        }
                                        if (null != t)
                                        {
                                            if (null != t.InnerText) Console.Out.WriteLine("t:" + t.ToString() + ">>>" + t.InnerText);
                                        }
                                        found = true;
                                    }
                                }
                            }
                            // t is null
                            // not found in previous

                            //(t is null)
                            //check next

                            if (!found)
                            {
                                Console.Out.WriteLine("@@@ Checking next sibling...");
                                xxxfield = pivotEnd;
                                if (null != xxxfield) rBegin = xxxfield.PreviousSibling<Run>();
                                if (null != xxxfield) rSep = xxxfield.NextSibling<Run>();
                                if (null != rSep) rText = rSep.NextSibling<Run>();
                                if (null != rText) rEnd = rText.NextSibling<Run>();
                                if (null != rText) t = rText.GetFirstChild<Text>();
                                if (null != t)
                                {
                                    if (null != t.Text && String.IsNullOrWhiteSpace(t.Text))
                                    {
                                        Console.Out.WriteLine("@@@ Not found in next sibling... giving up");
                                    }
                                    else
                                    {
                                        if (null != t.Text && t.Text.Length > 0)
                                        {
                                            Console.Out.WriteLine("@@@ Found in next sibling...>>>" + t.Text + "<<<");
                                            if (null != xxxfield)
                                            {
                                                if (null != xxxfield.InnerText) Console.Out.WriteLine("xxxfield:" + xxxfield.ToString() + ">>>" + xxxfield.InnerText);
                                            }
                                            if (null != rBegin)
                                            {
                                                if (null != rBegin.InnerText) Console.Out.WriteLine("rBegin:" + rBegin.ToString() + ">>>" + rBegin.InnerText);
                                            }
                                            if (null != rText)
                                            {
                                                if (null != rText.InnerText) Console.Out.WriteLine("rText:" + rText.ToString() + ">>>" + rText.InnerText);
                                            }
                                            if (null != rEnd)
                                            {
                                                if (null != rEnd.InnerText) Console.Out.WriteLine("rEnd:" + rEnd.ToString() + ">>>" + rEnd.InnerText);
                                            }
                                            if (null != t)
                                            {
                                                if (null != t.InnerText) Console.Out.WriteLine("t:" + t.ToString() + ">>>" + t.InnerText);
                                            }
                                            found = true;
                                        }

                                    }
                                }
                            } // if !found

                        }
                        else
                        {
                            found = true;
                        }



                        if (!isFormField && !fieldId.Contains("DOCPROPERTY"))
                        {
                            if (null != rBegin)
                            {
                                if (null != rBegin.InnerText && rBegin.InnerText.Contains("DOCPROPERTY"))
                                {
                                    fieldId = rBegin.InnerText.Trim();
                                }
                            }
                            else
                            {
                                if (null != rEnd)
                                {
                                    if (null != rEnd.InnerText && rEnd.InnerText.Contains("DOCPROPERTY"))
                                    {
                                        fieldId = rEnd.InnerText.Trim();
                                    }
                                }
                                else
                                {
                                    if (null != prevBegin)
                                    {
                                        if (null != prevBegin.InnerText && prevBegin.InnerText.Contains("DOCPROPERTY"))
                                        {
                                            if (null != t)
                                            {
                                                if (null != t.Text && !String.IsNullOrWhiteSpace(t.Text))
                                                {
                                                    Console.Out.WriteLine("##### " + prevBegin.InnerText);
                                                    fieldId = prevBegin.InnerText.Trim();
                                                }
                                            }
                                        }

                                    }
                                }

                            }
                        }


                        /*if (!fieldId.Contains("FORMTEXT"))
                        {
                            if (null != rBegin)
                            {
                                if (null != rBegin.InnerText && rBegin.InnerText.Contains("DOCPROPERTY"))
                                {
                                    fieldId = rBegin.InnerText.Trim();
                                }
                            }
                            else
                            {
                                if (null != rEnd)
                                {
                                    if (null != rEnd.InnerText && rEnd.InnerText.Contains("DOCPROPERTY"))
                                    {
                                        fieldId = rEnd.InnerText.Trim();
                                    }
                                }
                                else
                                {
                                    if (null != prevBegin)
                                    {
                                        if (null != prevBegin.InnerText && prevBegin.InnerText.Contains("DOCPROPERTY"))
                                        {
                                            if (null != t)
                                            {
                                                if (null != t.Text && !String.IsNullOrWhiteSpace(t.Text))
                                                {
                                                    Console.Out.WriteLine("##### " + prevBegin.InnerText);
                                                    fieldId = prevBegin.InnerText.Trim();
                                                }
                                            }
                                        }

                                    }
                                }

                            }
                        } */


                        if (isFormField)
                        {
                            if (null != t)
                            {
                                if (null != t.Text)
                                {
                                    var replaceFormValue = "";
                                    Console.Out.WriteLine("****Substitute form field " + t.Text);
                                    if (String.IsNullOrWhiteSpace(t.Text.Trim()))
                                        replaceFormValue = "[ ]";
                                    else
                                        replaceFormValue = "[" + t.Text.Trim() + "]";
                                    t.Text = formatText(replaceFormValue);
                                    rEnd.Remove();
                                    rSep.Remove();
                                    xxxfield.Remove();
                                    rBegin.Remove();

                                    if (null != rText)
                                    {
                                        var rp = rText.RunProperties;
                                        Highlight highlight = new Highlight() { Val = HighlightColorValues.Yellow };
                                        rp.Append(highlight);
                                    }


                                }
                            }
                        }
                        else
                        {
                            if (found)
                            {

                                Regex expr = new Regex(@"\s*(?<docProperty>\S+)\s+(?<aFieldName>\S+)\.*\s+(?<formatType>\S+)\s*");
                                var results = expr.Matches(fieldId);
                                string docProperty = null;
                                string aFieldName = null;
                                string formatType = null;

                                foreach (Match match in results)
                                {
                                    docProperty = match.Groups["docProperty"].Value;
                                    aFieldName = match.Groups["aFieldName"].Value;
                                    formatType = match.Groups["formatType"].Value;
                                }

                                if (null != docProperty) Console.Out.WriteLine("docProperty=" + docProperty);
                                if (null != aFieldName) Console.Out.WriteLine("aFieldName=" + aFieldName);
                                if (null != formatType) Console.Out.WriteLine("formatType=" + formatType);

                                if (null != aFieldName)
                                {
                                    var aFieldKey = aFieldName.Trim().ToUpper();
                                    string aFieldValue = null;
                                    if (fieldMap.ContainsKey(aFieldKey))
                                    {
                                        if (null != fieldMap[aFieldKey])
                                        {
                                            aFieldValue = fieldMap[aFieldKey].Trim();
                                        }
                                    }
                                    if (String.IsNullOrEmpty(aFieldValue)) aFieldValue = new string(" ");

                                    if (t != null)
                                    {
                                        if (t.Text != null && aFieldKey != null && fieldMap.ContainsKey(aFieldKey))
                                        {
                                            if (!fieldMap.ContainsKey(aFieldKey) || (String.IsNullOrWhiteSpace(aFieldValue)
                                            || aFieldValue.Contains("BOOKMARK_UNDEFINED")))
                                            {
                                                Run rBegin2 = null;
                                                Run rBegin1 = null;
                                                Run rBegin0 = null;
                                                Paragraph rParent = null;
                                                Run rParentFirst = null;
                                                Run rParentLeft = null;
                                                Run rParentLeftFirst = null;
                                                if (null != rBegin) rBegin2 = rBegin.PreviousSibling<Run>();
                                                if (null != rBegin2) rBegin1 = rBegin2.PreviousSibling<Run>();
                                                if (null != rBegin1) rBegin0 = rBegin1.PreviousSibling<Run>();
                                                if (null != rText) rParent = (Paragraph)rText.Parent;
                                                /*if (null != rParent) rParentFirst = (Run)rParent.GetFirstChild<Run>();
                                                if (null != rParent) rParentLeft = rParent.PreviousSibling<Run>();
                                                if (null != rParentLeft) rParentLeftFirst = rParentLeft.GetFirstChild<Run>();*/
                                                //t.Text = " ";
                                                //if (null != rEnd) rAdd = rEnd.NextSibling<Run>();
                                                //if (null != rAdd)
                                                //    rAdd.AppendChild(new Text(" "));
                                                //rAdd.AppendChild(new Text(fieldMap[aFieldKey]));*/
                                                //if (null != rParent) rParent.InnerXml.Replace(fieldId, "");
                                                //if (null != rText) rText.RemoveAllChildren();
                                                //if (null != rText) rText.Remove();
                                                if (null != rEnd) rEnd.Remove();
                                                if (null != rSep) rSep.Remove();
                                                if (null != rBegin) rBegin.Remove();
                                                if (null != t) t.Text = aFieldValue;
                                                if (null != xxxfield) xxxfield.Remove();
                                                //if (null != rParent) rParent.RemoveAllChildren();
                                                //if (null != rParent) rParent.AppendChild<Run>(new Run(new Text("")));                                        
                                                /*if (null != rBegin2) rBegin2.RemoveAllChildren();
                                                if (null != rBegin2) rBegin2.Remove();
                                                if (null != rBegin1) rBegin1.RemoveAllChildren();
                                                if (null != rBegin1) rBegin1.Remove();
                                                if (null != rBegin0) rBegin0.RemoveAllChildren();
                                                if (null != rBegin0) rBegin0.Remove();*/
                                                //if (null != t.Text) Console.Out.WriteLine("****Substitute value " + t.Text + "with " + fieldMap[aFieldKey]);
                                                //rText.Remove();
                                            }
                                            else
                                            {
                                                Console.Out.WriteLine("****Substitute value " + t.Text + "with " + aFieldValue);
                                                if (fieldMap.ContainsKey(aFieldKey) && !(String.IsNullOrEmpty(aFieldValue)))
                                                {



                                                    if (nonEditableArray.Contains(aFieldKey))
                                                    {
                                                        t.Text = formatText(aFieldValue);
                                                    }
                                                    else
                                                    {


                                                        Run rBegin2 = null;
                                                        Run rBegin1 = null;
                                                        Run rBegin0 = null;
                                                        Paragraph rParent = null;
                                                        Run rParentFirst = null;
                                                        Run rParentLeft = null;
                                                        Run rParentLeftFirst = null;
                                                        if (null != rBegin) rBegin2 = rBegin.PreviousSibling<Run>();
                                                        if (null != rBegin2) rBegin1 = rBegin2.PreviousSibling<Run>();
                                                        if (null != rBegin1) rBegin0 = rBegin1.PreviousSibling<Run>();
                                                        if (null != rText) rParent = (Paragraph)rText.Parent;
                                                        /*if (null != rParent) rParentFirst = (Run)rParent.GetFirstChild<Run>();
                                                        if (null != rParent) rParentLeft = rParent.PreviousSibling<Run>();
                                                        if (null != rParentLeft) rParentLeftFirst = rParentLeft.GetFirstChild<Run>();*/
                                                        //t.Text = fieldMap[aFieldKey];
                                                        //if (null != rEnd) rAdd = rEnd.NextSibling<Run>();
                                                        //if (null != rAdd)
                                                        //    rAdd.AppendChild(new Text(fieldMap[aFieldKey]));
                                                        //rAdd = rText.NextSibling<Run>();
                                                        /*if (null != rAdd)
                                                            rAdd.AppendChild(new Text(fieldMap[aFieldKey]));*/
                                                        //if (null != rParent) rParent.InnerXml.Replace(fieldId, "");
                                                        //if (null != rText) rText.RemoveAllChildren();
                                                        //if (null != rText) rText.Remove();

                                                        string aaFieldKey = null;
                                                        string aaFieldValue = null;
                                                        string aaaFieldKey = null;
                                                        string aaaFieldValue = null;

                                                        if (null != rEnd) nFormat = rEnd.NextSibling<Run>();
                                                        if (null != nFormat) nBegin = nFormat.NextSibling<Run>();
                                                        if (null != nBegin) nTag = nBegin.NextSibling<Run>();
                                                        if (null != nTag) nSep = nTag.NextSibling<Run>();
                                                        if (null != nSep) nText = nSep.NextSibling<Run>();
                                                        if (null != nText) nEnd = nText.NextSibling<Run>();
                                                        if (null != nText) nt = nText.GetFirstChild<Text>();
                                                        if (null != nEnd) nnFormat = nEnd.NextSibling<Run>();
                                                        if (null != nnFormat) nnBegin = nnFormat.NextSibling<Run>();
                                                        if (null != nnBegin) nnTag = nnBegin.NextSibling<Run>();
                                                        if (null != nnTag) nnSep = nnTag.NextSibling<Run>();
                                                        if (null != nnSep) nnText = nnSep.NextSibling<Run>();
                                                        if (null != nnText) nnEnd = nnText.NextSibling<Run>();
                                                        if (null != nnText) nnt = nnText.GetFirstChild<Text>();


                                                        if (null != nnt)
                                                        {
                                                            if (null != nnt.Text && !String.IsNullOrWhiteSpace(nnt.Text))
                                                            {
                                                                aaaFieldKey = nnt.Text.Trim().ToUpper();
                                                                if (null != aaaFieldKey && fieldMap.ContainsKey(aaaFieldKey)) aaaFieldValue = fieldMap[aaaFieldKey];
                                                                if (null != aaaFieldValue && !String.IsNullOrWhiteSpace(aaaFieldValue) && aaaFieldValue.ToUpper() != "BOOKMARK_UNDEFINED")
                                                                {
                                                                    nnt.Text = formatText(aaaFieldValue);
                                                                    if (null != nnEnd) nnEnd.Remove();
                                                                    if (null != nnSep) nnSep.Remove();
                                                                    if (null != nnTag) nnTag.Remove();
                                                                    if (null != nnBegin) nnBegin.Remove();
                                                                    //if (null != rText) rText = new Run(new Text(aFieldValue));

                                                                }
                                                            }
                                                        }

                                                        if (null != nt)
                                                        {
                                                            if (null != nt.Text && !String.IsNullOrWhiteSpace(nt.Text))
                                                            {
                                                                aaFieldKey = nt.Text.Trim().ToUpper();
                                                                if (null != aaFieldKey && fieldMap.ContainsKey(aaFieldKey)) aaFieldValue = fieldMap[aaFieldKey];
                                                                if (null != aaFieldValue && !String.IsNullOrWhiteSpace(aaFieldValue) && aaFieldValue.ToUpper() != "BOOKMARK_UNDEFINED")
                                                                {
                                                                    nt.Text = formatText(aaFieldValue);
                                                                    if (null != nEnd) nEnd.Remove();
                                                                    if (null != nSep) nSep.Remove();
                                                                    if (null != nTag) nTag.Remove();
                                                                    if (null != nBegin) nBegin.Remove();
                                                                }
                                                            }
                                                        }







                                                        if (null != rEnd) rEnd.Remove();
                                                        if (null != rSep) rSep.Remove();
                                                        if (null != rBegin) rBegin.Remove();
                                                        if (null != t && null != t.Text) t.Text = formatText(aFieldValue);
                                                        //if (null != rText) rText = new Run(new Text(aFieldValue));
                                                        if (null != xxxfield) xxxfield.Remove();
                                                        //if (null != rParent) rParent.RemoveAllChildren();
                                                        //if (null != rParent) rParent.AppendChild<Run>(new Run(new Text(aFieldValue)));
                                                        /*if (null != rBegin2) rBegin2.RemoveAllChildren();
                                                        if (null != rBegin2) rBegin2.Remove();
                                                        if (null != rBegin1) rBegin1.RemoveAllChildren();
                                                        if (null != rBegin1) rBegin1.Remove();
                                                        if (null != rBegin0) rBegin0.RemoveAllChildren();
                                                        if (null != rBegin0) rBegin0.Remove();*/


                                                    }
                                                }
                                            }

                                        }
                                    }
                                    else //field name is CHARFORMAT or something
                                    {
                                        if (null != rEnd)
                                        {

                                            if (null != rEnd.InnerText)
                                            {
                                                if (!(String.IsNullOrWhiteSpace(rEnd.InnerText)) && fieldMap.ContainsKey(rEnd.InnerText.ToUpper()))
                                                {
                                                    //rEnd.SetText(fieldMap[rEnd.InnerText]);
                                                    if (null != fieldMap[rEnd.InnerText.ToUpper()])
                                                        rEnd.SetText(fieldMap[rEnd.InnerText.ToUpper()]);
                                                    else
                                                        rEnd.SetText(formatText(null));
                                                    /*Run rBegin2 = null;
                                                    Run rBegin1 = null;
                                                    Run rBegin0 = null;
                                                    Run rParent = null;
                                                    Run rParentFirst = null;
                                                    Run rParentLeft = null;
                                                    Run rParentLeftFirst = null;                                        
                                                    if (null != rBegin) rBegin2 = rBegin.PreviousSibling<Run>();
                                                    if (null != rBegin2) rBegin1 = rBegin2.PreviousSibling<Run>();
                                                    if (null != rBegin1) rBegin0 = rBegin1.PreviousSibling<Run>();
                                                    if (null != rEnd) rParent = (Run)rEnd.Parent;
                                                    if (null != rParent) rParentFirst = (Run)rParent.GetFirstChild<Run>();
                                                    if (null != rParent) rParentLeft = rParent.PreviousSibling<Run>();
                                                    if (null != rParentLeft) rParentLeftFirst = rParentLeft.GetFirstChild<Run>();
                                                    rAdd = rEnd.NextSibling<Run>();
                                                    if (null != rAdd)
                                                        rAdd.AppendChild(new Text(fieldMap[rEnd.InnerText]));
                                                    if (null != rEnd) rEnd.RemoveAllChildren();
                                                    if (null != rEnd) rEnd.Remove();
                                                    if (null != rText) rText.RemoveAllChildren();
                                                    if (null != rText) rText.Remove();
                                                    if (null != rSep) rSep.RemoveAllChildren();
                                                    if (null != rSep) rSep.Remove();
                                                    if (null != rBegin) rBegin.RemoveAllChildren();
                                                    if (null != rBegin) rBegin.Remove();
                                                    if (null != rBegin2) rBegin2.RemoveAllChildren();
                                                    if (null != rBegin2) rBegin2.Remove();
                                                    if (null != rBegin1) rBegin1.RemoveAllChildren();
                                                    if (null != rBegin1) rBegin1.Remove();
                                                    if (null != rBegin0) rBegin0.RemoveAllChildren();
                                                    if (null != rBegin0) rBegin0.Remove();
                                                    if (null != rParentFirst) rParentFirst.RemoveAllChildren();
                                                    if (null != rParentFirst) rParentFirst.Remove();
                                                    if (null != rParentLeftFirst) rParentLeftFirst.RemoveAllChildren();
                                                    if (null != rParentLeftFirst) rParentLeftFirst.Remove();
                                                    if (null != xxxfield) xxxfield.RemoveAllChildren();
                                                    if (null != xxxfield) xxxfield.Remove();*/
                                                }
                                                else
                                                {
                                                    rEnd.SetText(formatText(null));
                                                }
                                            }


                                        }
                                        else
                                        {
                                            // rEnd = nuill rEnd.SetText(" ");
                                            /* if (rEnd != null)
                                             {
                                                 Run rObject1 = null;
                                                 Run rObject2 = null;
                                                 Run rObject3 = null;
                                                 rObject1 = rEnd.GetFirstChild<Run>();
                                                 if (null != rObject1) rObject2 = rObject1.NextSibling<Run>();
                                                 if (null != rObject2) rObject3 = rObject2.NextSibling<Run>();

                                                 if (null != rObject1) 
                                                 {
                                                     if (null != rObject1 && null != rObject1.GetTextElement)
                                                         rObject1.Remove();
                                                     if (null != rObject2 && null != rObject2.GetTextElement)
                                                         rObject2.Remove();
                                                     if (null != rObject3)
                                                 }
                                             }*/
                                            /*t.Text = "";
                                            Run rBegin2 = null;
                                            Run rBegin1 = null;
                                            Run rBegin0 = null;
                                            Run rParent = null;
                                            Run rParentFirst = null;
                                            Run rParentLeft = null;
                                            Run rParentLeftFirst = null;
                                            if (null != rBegin) rBegin2 = rBegin.PreviousSibling<Run>();
                                            if (null != rBegin2) rBegin1 = rBegin2.PreviousSibling<Run>();
                                            if (null != rBegin1) rBegin0 = rBegin1.PreviousSibling<Run>();
                                            if (null != rEnd) rParent = (Run)rEnd.Parent;
                                            if (null != rParent) rParentFirst = (Run)rParent.GetFirstChild<Run>();
                                            if (null != rParent) rParentLeft = rParent.PreviousSibling<Run>();
                                            if (null != rParentLeft) rParentLeftFirst = rParentLeft.GetFirstChild<Run>();
                                            rAdd = rEnd.NextSibling<Run>();                                        
                                            if (null != rAdd) rAdd.AppendChild(new Text(""));
                                            if (null != rEnd) rEnd.RemoveAllChildren();
                                            if (null != rEnd) rEnd.Remove();                                        
                                            if (null != rText) rText.RemoveAllChildren();
                                            if (null != rText) rText.Remove();
                                            if (null != rSep) rSep.RemoveAllChildren();
                                            if (null != rSep) rSep.Remove();
                                            if (null != rBegin) rBegin.RemoveAllChildren();
                                            if (null != rBegin) rBegin.Remove();
                                            if (null != rBegin2) rBegin2.RemoveAllChildren();
                                            if (null != rBegin2) rBegin2.Remove();
                                            if (null != rBegin1) rBegin1.RemoveAllChildren(); 
                                            if (null != rBegin1) rBegin1.Remove();
                                            if (null != rBegin0) rBegin0.RemoveAllChildren(); 
                                            if (null != rBegin0) rBegin0.Remove();
                                            if (null != rParentFirst) rParentFirst.RemoveAllChildren();
                                            if (null != rParentFirst) rParentFirst.Remove();
                                            if (null != rParentLeftFirst) rParentLeftFirst.RemoveAllChildren();
                                            if (null != rParentLeftFirst) rParentLeftFirst.Remove();
                                            if (null != xxxfield) xxxfield.RemoveAllChildren();
                                            if (null != xxxfield) xxxfield.Remove();*/

                                        }
                                    }
                                }
                                else
                                {

                                }
                            }
                            else
                            {
                                Console.Out.WriteLine("@@@ field value not found.");
                            }
                        }


                        Console.Out.WriteLine("*** Field " + j.ToString() + " ends *********************************");

                        prevRun = pivotRun;
                        prevBegin = pivotBegin;
                        prevEnd = pivotEnd;










                        /* DocumentProperty cField = (custom[fieldId]);
                         if (null != cField)
                         {
                             Console.Out.WriteLine(">>> " + cField.Name + ": " + cField.Value);
                         }*/




                        /*int fieldNameStart = field.Text.LastIndexOf(FieldDelimeter, System.StringComparison.Ordinal);

                        if (fieldNameStart >= 0)
                        {
                            var fieldName = field.Text.Substring(fieldNameStart + FieldDelimeter.Length).Trim();

                            Run xxxfield = (Run)field.Parent;

                            Run rBegin = xxxfield.PreviousSibling<Run>();
                            Run rSep = xxxfield.NextSibling<Run>();
                            Run rText = rSep.NextSibling<Run>();
                            Run rEnd = rText.NextSibling<Run>();

                            if (null != xxxfield)
                            {

                                Text t = rText.GetFirstChild<Text>();
                               //custom.CustomHash;
                               Console.Out.WriteLine(t.ToString());

                            }
                        }*/

                    }


                    /*MainDocumentPart mainPart = wordDocument.MainDocumentPart;
                    var fields = mainPart.Document.Body.Descendants<FieldCode>();

                    foreach (var field in fields)
                    {
                        //if (field.GetType() == typeof(FormFieldData))
                        //{

                            Console.Out.WriteLine("***"+ field.ToString());
                            Console.Out.WriteLine("***"+ field.GetType());
                        //Console.Out.WriteLine("***" + ((FieldCode)field.FirstChild).Val.InnerText);
                        if (((FieldCode)field.FirstChild).Val.InnerText.Equals("WorkerName"))
                            {
                                TextInput text = field.Descendants<TextInput>().First();
                                SetFormFieldValue(text, "Put some text inside the field");
                            }
                        //}
                    }*/

                    /*if (null != wordDocument)
                    {

                        string aFieldDelimeter = @" MERGEFIELD ";
                        List<string> alisteChamps = new List<string>();

                        //foreach (var footer in wordDocument.MainDocumentPart.Document)
                        //{

                            foreach (var field in wordDocument.MainDocumentPart.RootElement.Descendants<FieldCode>())
                            {

                                int fieldNameStart = field.Text.LastIndexOf(aFieldDelimeter, System.StringComparison.Ordinal);

                                if (fieldNameStart >= 0)
                                {
                                    var fieldName = field.Text.Substring(fieldNameStart + aFieldDelimeter.Length).Trim();
                                    Console.Out.WriteLine("******" + fieldName.ToString());

                                Run xxxfield = (Run)field.Parent;

                                    Run rBegin = xxxfield.PreviousSibling<Run>();
                                    Run rSep = xxxfield.NextSibling<Run>();
                                    Run rText = rSep.NextSibling<Run>();
                                    Run rEnd = rText.NextSibling<Run>();

                                    if (null != xxxfield)
                                    {

                                        Text t = rText.GetFirstChild<Text>();
                                    //t.Text = replacementText;
                                    Console.Out.WriteLine("*******" + t.Text.ToString());

                                    }
                                }

                            }


                        //}
                    }*/

                    setCustomProperty(wordDocument, WopiOptions.Value.ProcessedFlag, WopiOptions.Value.ApplicationName, CustomPropertyTypes.Text);

                    /*Remove VBA part
                        var docPart = wordDocument.MainDocumentPart;

                        // Look for the vbaProject part. If it is there, delete it.
                        var vbaPart = docPart.VbaProjectPart;
                        if (vbaPart != null)
                        {
                            // Delete the vbaProject part and then save the document.
                            docPart.DeletePart(vbaPart);
                            docPart.Document.Save();

                            // Change the document type to
                            // not macro-enabled.
                            wordDocument.ChangeDocumentType(WordprocessingDocumentType.Document);

                            // Track that the document has been changed.

                        }
                        //changeCompatibilityModeOfDocumentPart(wordDocument.MainDocumentPart);
                    */

                    wordDocument.Save();
                    wordDocument.Close();

                }
            } // end using

        }



        public void runMacroOriginal(string newFileName, Dictionary<string, string> fieldMap, bool processedByWWA = false, bool normalize = false, string userRole = null)
        {
            Regex fileVer = new Regex(@"\s*(?<userID>)_(?<attachmentID>)_(?<versionNo>)\..*");
            var result = fileVer.Matches(newFileName);
            string userID = null;
            string attachmentID = null;
            string versionNo = null;

            /*if (Int32.TryParse(versionNo, out int v))
            {
                if (v > 1) return;
            }*/
            //if (processedByWWA) return;

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(newFileName, true))
            {

                if (null != wordDocument)
                {
                    /*for (var field in wordDocument.MainDocumentPart.RootElement.Descendants<FieldCode>())
                    {

                    }*/

                    /*List<Picture> pictures = new List<Picture>(wordDocument.MainDocumentPart.RootElement.Descendants<Picture>());

                    foreach (Picture p in pictures)
                    {
                        p.Remove();
                    }*/

                    //headerPart.DeleteParts(imagePartList);

                    foreach (Paragraph p in wordDocument.MainDocumentPart.Document.Body.Descendants<Paragraph>().Where<Paragraph>(p => p.InnerText.Contains(WopiOptions.Value.ConversionEngine)))
                    {
                        // Do something with the Paragraphs.
                        p.Remove();
                    }

                    if (wordDocument.MainDocumentPart.HeaderParts.Count() > 0)
                    {
                        foreach (HeaderPart headerPart in wordDocument.MainDocumentPart.HeaderParts)
                        {
                            foreach (Paragraph p in headerPart.Header.Descendants<Paragraph>().Where<Paragraph>(p => p.InnerText.Contains(WopiOptions.Value.ConversionEngine)))
                            {
                                p.Remove();
                            }
                        }
                    }

                    if (wordDocument.MainDocumentPart.FooterParts.Count() > 0)
                    {
                        foreach (FooterPart footerPart in wordDocument.MainDocumentPart.FooterParts)
                        {
                            foreach (Paragraph p in footerPart.Footer.Descendants<Paragraph>().Where<Paragraph>(p => p.InnerText.Contains(WopiOptions.Value.ConversionEngine)))
                            {
                                p.Remove();
                            }
                        }
                    }


                    if (WopiOptions?.Value?.DocumentProtection == TRUE)
                    {
                        if (getIDPList().Contains("all") || getIDPList().Contains(userRole) || userRole == null)
                        {
                            ArrayList docProtClassList = new ArrayList();
                            ArrayList docProtTypeList = new ArrayList();

                            if (null != WopiOptions.Value.DocumentProtectionClass && WopiOptions.Value.DocumentProtectionClass.Length > 0)
                                docProtClassList.AddRange(WopiOptions.Value.DocumentProtectionClass);

                            if (null != WopiOptions.Value.DocumentProtectionType && WopiOptions.Value.DocumentProtectionType.Length > 0)
                                docProtTypeList.AddRange(WopiOptions.Value.DocumentProtectionType);


                            var dsp = wordDocument.MainDocumentPart.DocumentSettingsPart;
                            foreach (DocumentProtection dp in wordDocument.MainDocumentPart.DocumentSettingsPart.Settings.ChildElements.OfType<DocumentProtection>())
                            {
                                if (!(String.IsNullOrEmpty(WopiOptions.Value.DocumentProtectionFlag)))
                                {
                                    if (null != dp.Enforcement)
                                    {

                                        if (dp.Enforcement == new OnOffValue(true))
                                        {
                                            //setCustomProperty(wordDocument, WopiOptions.Value.DocumentProtectionFlag, WopiOptions.Value.DocumentProtectionClass + "," + WopiOptions.Value.DocumentProtectionType, CustomPropertyTypes.Text);
                                            string docProtString = null;
                                            string docProtClass = null;
                                            string docProtType = null;
                                            if (docProtClassList.Contains("edit") && null != dp.Edit)
                                            {
                                                docProtClass = "edit";

                                                /*//
                                                // Summary:
                                                //     No Editing Restrictions.
                                                //     When the item is serialized out as xml, its value is "none".
                                                None = 0,
                                                //
                                                // Summary:
                                                //     Allow No Editing.
                                                //     When the item is serialized out as xml, its value is "readOnly".
                                                ReadOnly = 1,
                                                //
                                                // Summary:
                                                //     Allow Editing of Comments.
                                                //     When the item is serialized out as xml, its value is "comments".
                                                Comments = 2,
                                                //
                                                // Summary:
                                                //     Allow Editing With Revision Tracking.
                                                //     When the item is serialized out as xml, its value is "trackedChanges".
                                                TrackedChanges = 3,
                                                //
                                                // Summary:
                                                //     Allow Editing of Form Fields.
                                                //     When the item is serialized out as xml, its value is "forms".
                                                Forms = 4 */
                                                if (dp.Edit == DocumentProtectionValues.ReadOnly) docProtType = "1";
                                                if (dp.Edit == DocumentProtectionValues.Comments) docProtType = "2";
                                                if (dp.Edit == DocumentProtectionValues.TrackedChanges) docProtType = "3";
                                                if (dp.Edit == DocumentProtectionValues.Forms) docProtType = "4";
                                                if (dp.Edit == DocumentProtectionValues.None) docProtType = "0";

                                                var dpnew = new DocumentProtection()
                                                {
                                                    Edit = dp.Edit,
                                                    Enforcement = new OnOffValue(false),
                                                    Formatting = dp.Formatting
                                                    //CryptographicProviderType = CryptProviderValues.RsaFull,
                                                    //CryptographicAlgorithmClass = CryptAlgorithmClassValues.Hash,
                                                    //CryptographicAlgorithmType = CryptAlgorithmValues.TypeAny,
                                                    //CryptographicAlgorithmSid = 4,
                                                    //CryptographicSpinCount = 100000U,
                                                    //Hash = "2krUoz1qWd0WBeXqVrOq81l8xpk=",
                                                    //Salt = "9kIgmDDYtt2r5U2idCOwMA=="
                                                };


                                                if (!(null == dsp || null == dsp.Settings))
                                                {
                                                    dsp.Settings.ReplaceChild(dpnew, dp);
                                                }
                                            }

                                            // handles other types of protection here


                                            setCustomProperty(wordDocument, WopiOptions.Value.DocumentProtectionFlag, docProtClass + "," + docProtType, CustomPropertyTypes.Text);

                                            var docProtKey = Path.GetFileNameWithoutExtension(newFileName);
                                            if (null != docProtKey && docProtection.ContainsKey(docProtKey))
                                            {
                                                docProtection[docProtKey] = TRUE;
                                            }
                                            else
                                            {
                                                docProtection.Add(docProtKey, TRUE);
                                            }


                                        }
                                    }
                                }
                                //dp.Remove();




                                /*
                                < w:documentProtection w:edit = "forms"
                                  w: formatting = "1"
                                  w: enforcement = "1" />
                                */

                            }
                        }
                    }



                    //const string FieldDelimeter = @" MERGEFIELD ";
                    string FieldDelimeter = @" DOCPROPERTY ";
                    List<string> listeChamps = new List<string>();


                    if (normalize)
                    {
                        normalizeMarkup(wordDocument);
                        normalizeFieldCodesRuns(wordDocument);
                    }

                    Run prevRun = null;
                    Run prevBegin = null;
                    Run prevEnd = null;

                    string[] nonEditable = WopiOptions.Value.BookmarksNonEditable;

                    /* new string[] {
                      "LetterDate",
                     "NameAndAddressLn1",
                     "NameAndAddressLn2",
                     "NameAndAddressLn3",
                     "NameAndAddressLn4",
                     "NameAndAddressLn5",
                     "NameAndAddressLn6",
                     "NameAndAddressLn7",
                     "(P)",
                     "\"CC\"",
                     " CC ",
                     "COPYIND",
                     "CCTOKENLIST",
                     "IPTOKENLIST",
                     "PRIMARYRECIPCONTACTNAME" }; */



                    ArrayList nonEditableArray = new ArrayList();
                    for (var k = 0; k < nonEditable.Length; k++)
                    {
                        nonEditableArray.Add(nonEditable[k].ToUpper());
                    }

                    //nonEditableArray.AddRange(nonEditable);



                    int j = 0;
                    foreach (var field in wordDocument.MainDocumentPart.RootElement.Descendants<FieldCode>())
                    {
                        j++;
                        Console.Out.WriteLine("***Field " + j.ToString() + ">>>" + "Field Type:" + field.ToString() + "<<<Field Text:>>>" + field.Text.ToString() + "<<<");
                        if (null != field.InnerXml) Console.Out.WriteLine("***** internal XML:" + field.InnerXml + " *****");
                    }

                    j = 0;

                    foreach (var field in wordDocument.MainDocumentPart.RootElement.Descendants<FieldCode>())
                    {
                        j++;
                        Console.Out.WriteLine("*** Field " + j.ToString() + " starts *********************************");
                        Console.Out.WriteLine("Field Type:" + field.ToString());
                        Console.Out.WriteLine("Field Text:>>>" + field.Text.ToString() + "<<<");
                        bool isFormField = false;
                        var fieldId = field.Text.Trim().ToString();
                        if (null != fieldId)
                        {
                            if (fieldId.Contains("FORMTEXT"))
                            {
                                isFormField = true;
                            }
                            Console.Out.WriteLine("fieldId:" + fieldId.ToString());
                        }

                        /*if (field.Text.ToString() == " DOCPROPERTY  WorkerName  \\* CHARFORMAT ")
                         {
                             field.Text = new string("REPLACED WORKER NAME");


                             Console.Out.WriteLine("Replaceing Field Text:>>>" + field.Text.ToString() + "<<<");
                             Console.Out.WriteLine("with:>>>REPLACED_WORKER_NAME<<<");
                         }*/


                        Run xxxfield = null;
                        Run rBegin = null;
                        Run rSep = null;
                        Run rText = null;
                        Run rEnd = null;
                        Text t = null;

                        Run nFormat = null;
                        Run nBegin = null;
                        Run nTag = null;
                        Run nSep = null;
                        Run nText = null;
                        Run nEnd = null;
                        Text nt = null;

                        Run nnFormat = null;
                        Run nnBegin = null;
                        Run nnTag = null;
                        Run nnSep = null;
                        Run nnText = null;
                        Run nnEnd = null;
                        Text nnt = null;


                        Run pivotRun = null;
                        Run pivotEnd = null;
                        Run pivotBegin = null;
                        Run rAdd = null;

                        xxxfield = (Run)field.Parent;
                        if (null != xxxfield) rBegin = xxxfield.PreviousSibling<Run>();
                        if (null != xxxfield) rSep = xxxfield.NextSibling<Run>();
                        if (null != rSep) rText = rSep.NextSibling<Run>();
                        if (null != rText) rEnd = rText.NextSibling<Run>();
                        if (null != rText) t = rText.GetFirstChild<Text>();

                        pivotRun = xxxfield;
                        pivotBegin = rBegin;
                        pivotEnd = rEnd;
                        bool found = false;

                        Console.Out.WriteLine("@@@ Checking field Id...");
                        if (null != xxxfield)
                        {
                            if (null != xxxfield.InnerText) Console.Out.WriteLine("xxxfield:" + xxxfield.ToString() + ">>>" + xxxfield.InnerText);
                        }
                        if (null != rBegin)
                        {
                            if (null != rBegin.InnerText) Console.Out.WriteLine("rBegin:" + rBegin.ToString() + ">>>" + rBegin.InnerText);
                        }
                        if (null != rText)
                        {
                            if (null != rText.InnerText) Console.Out.WriteLine("rText:" + rText.ToString() + ">>>" + rText.InnerText);
                        }
                        if (null != rEnd)
                        {
                            if (null != rEnd.InnerText) Console.Out.WriteLine("rEnd:" + rEnd.ToString() + ">>>" + rEnd.InnerText);
                        }
                        if (null != t)
                        {
                            if (null != t.InnerText) Console.Out.WriteLine("t:" + t.ToString() + ">>>" + t.InnerText);
                        }



                        //t.SetText("Vincent");

                        if (!isFormField && (null == t || null == t.InnerText || String.IsNullOrWhiteSpace(t.InnerText)))
                        //if (!fieldId.Contains("DOCPROPERTY"))
                        {
                            // check the previous sibling to see if it is there
                            Console.Out.WriteLine("@@@ Checking previous sibling...");
                            xxxfield = rBegin;
                            if (null != xxxfield) rBegin = xxxfield.PreviousSibling<Run>();
                            if (null != xxxfield) rSep = xxxfield.NextSibling<Run>();
                            if (null != rSep) rText = rSep.NextSibling<Run>();
                            if (null != rText) rEnd = rText.NextSibling<Run>();
                            if (null != rText) t = rText.GetFirstChild<Text>();
                            if (null != t)
                            {
                                if (null != t.Text && String.IsNullOrWhiteSpace(t.Text))
                                {
                                    Console.Out.WriteLine("@@@ Not found in previous sibling...");
                                }
                                else
                                {
                                    if (null != t.Text && t.Text.Length > 0)
                                    {
                                        Console.Out.WriteLine("@@@ Found in previous sibling... >>>" + t.Text + "<<<");
                                        if (null != xxxfield)
                                        {
                                            if (null != xxxfield.InnerText) Console.Out.WriteLine("xxxfield:" + xxxfield.ToString() + ">>>" + xxxfield.InnerText);
                                        }
                                        if (null != rBegin)
                                        {
                                            if (null != rBegin.InnerText) Console.Out.WriteLine("rBegin:" + rBegin.ToString() + ">>>" + rBegin.InnerText);
                                        }
                                        if (null != rText)
                                        {
                                            if (null != rText.InnerText) Console.Out.WriteLine("rText:" + rText.ToString() + ">>>" + rText.InnerText);
                                        }
                                        if (null != rEnd)
                                        {
                                            if (null != rEnd.InnerText) Console.Out.WriteLine("rEnd:" + rEnd.ToString() + ">>>" + rEnd.InnerText);
                                        }
                                        if (null != t)
                                        {
                                            if (null != t.InnerText) Console.Out.WriteLine("t:" + t.ToString() + ">>>" + t.InnerText);
                                        }
                                        found = true;
                                    }
                                }
                            }
                            // t is null
                            // not found in previous

                            //(t is null)
                            //check next

                            if (!found)
                            {
                                Console.Out.WriteLine("@@@ Checking next sibling...");
                                xxxfield = pivotEnd;
                                if (null != xxxfield) rBegin = xxxfield.PreviousSibling<Run>();
                                if (null != xxxfield) rSep = xxxfield.NextSibling<Run>();
                                if (null != rSep) rText = rSep.NextSibling<Run>();
                                if (null != rText) rEnd = rText.NextSibling<Run>();
                                if (null != rText) t = rText.GetFirstChild<Text>();
                                if (null != t)
                                {
                                    if (null != t.Text && String.IsNullOrWhiteSpace(t.Text))
                                    {
                                        Console.Out.WriteLine("@@@ Not found in next sibling... giving up");
                                    }
                                    else
                                    {
                                        if (null != t.Text && t.Text.Length > 0)
                                        {
                                            Console.Out.WriteLine("@@@ Found in next sibling...>>>" + t.Text + "<<<");
                                            if (null != xxxfield)
                                            {
                                                if (null != xxxfield.InnerText) Console.Out.WriteLine("xxxfield:" + xxxfield.ToString() + ">>>" + xxxfield.InnerText);
                                            }
                                            if (null != rBegin)
                                            {
                                                if (null != rBegin.InnerText) Console.Out.WriteLine("rBegin:" + rBegin.ToString() + ">>>" + rBegin.InnerText);
                                            }
                                            if (null != rText)
                                            {
                                                if (null != rText.InnerText) Console.Out.WriteLine("rText:" + rText.ToString() + ">>>" + rText.InnerText);
                                            }
                                            if (null != rEnd)
                                            {
                                                if (null != rEnd.InnerText) Console.Out.WriteLine("rEnd:" + rEnd.ToString() + ">>>" + rEnd.InnerText);
                                            }
                                            if (null != t)
                                            {
                                                if (null != t.InnerText) Console.Out.WriteLine("t:" + t.ToString() + ">>>" + t.InnerText);
                                            }
                                            found = true;
                                        }

                                    }
                                }
                            } // if !found

                        }
                        else
                        {
                            found = true;
                        }



                        if (!isFormField && !fieldId.Contains("DOCPROPERTY"))
                        {
                            if (null != rBegin)
                            {
                                if (null != rBegin.InnerText && rBegin.InnerText.Contains("DOCPROPERTY"))
                                {
                                    fieldId = rBegin.InnerText.Trim();
                                }
                            }
                            else
                            {
                                if (null != rEnd)
                                {
                                    if (null != rEnd.InnerText && rEnd.InnerText.Contains("DOCPROPERTY"))
                                    {
                                        fieldId = rEnd.InnerText.Trim();
                                    }
                                }
                                else
                                {
                                    if (null != prevBegin)
                                    {
                                        if (null != prevBegin.InnerText && prevBegin.InnerText.Contains("DOCPROPERTY"))
                                        {
                                            if (null != t)
                                            {
                                                if (null != t.Text && !String.IsNullOrWhiteSpace(t.Text))
                                                {
                                                    Console.Out.WriteLine("##### " + prevBegin.InnerText);
                                                    fieldId = prevBegin.InnerText.Trim();
                                                }
                                            }
                                        }

                                    }
                                }

                            }
                        }


                        /*if (!fieldId.Contains("FORMTEXT"))
                        {
                            if (null != rBegin)
                            {
                                if (null != rBegin.InnerText && rBegin.InnerText.Contains("DOCPROPERTY"))
                                {
                                    fieldId = rBegin.InnerText.Trim();
                                }
                            }
                            else
                            {
                                if (null != rEnd)
                                {
                                    if (null != rEnd.InnerText && rEnd.InnerText.Contains("DOCPROPERTY"))
                                    {
                                        fieldId = rEnd.InnerText.Trim();
                                    }
                                }
                                else
                                {
                                    if (null != prevBegin)
                                    {
                                        if (null != prevBegin.InnerText && prevBegin.InnerText.Contains("DOCPROPERTY"))
                                        {
                                            if (null != t)
                                            {
                                                if (null != t.Text && !String.IsNullOrWhiteSpace(t.Text))
                                                {
                                                    Console.Out.WriteLine("##### " + prevBegin.InnerText);
                                                    fieldId = prevBegin.InnerText.Trim();
                                                }
                                            }
                                        }

                                    }
                                }

                            }
                        } */


                        if (isFormField)
                        {
                            if (null != t)
                            {
                                if (null != t.Text)
                                {
                                    var replaceFormValue = "";
                                    Console.Out.WriteLine("****Substitute form field " + t.Text);
                                    if (String.IsNullOrWhiteSpace(t.Text.Trim()))
                                        replaceFormValue = "[ ]";
                                    else
                                        replaceFormValue = "[" + t.Text.Trim() + "]";
                                    t.Text = formatText(replaceFormValue);
                                    rEnd.Remove();
                                    rSep.Remove();
                                    xxxfield.Remove();
                                    rBegin.Remove();

                                    if (null != rText)
                                    {
                                        var rp = rText.RunProperties;
                                        Highlight highlight = new Highlight() { Val = HighlightColorValues.Yellow };
                                        rp.Append(highlight);
                                    }


                                }
                            }
                        }
                        else
                        {
                            if (found)
                            {

                                Regex expr = new Regex(@"\s*(?<docProperty>\S+)\s+(?<aFieldName>\S+)\.*\s+(?<formatType>\S+)\s*");
                                var results = expr.Matches(fieldId);
                                string docProperty = null;
                                string aFieldName = null;
                                string formatType = null;

                                foreach (Match match in results)
                                {
                                    docProperty = match.Groups["docProperty"].Value;
                                    aFieldName = match.Groups["aFieldName"].Value;
                                    formatType = match.Groups["formatType"].Value;
                                }

                                if (null != docProperty) Console.Out.WriteLine("docProperty=" + docProperty);
                                if (null != aFieldName) Console.Out.WriteLine("aFieldName=" + aFieldName);
                                if (null != formatType) Console.Out.WriteLine("formatType=" + formatType);

                                if (null != aFieldName)
                                {
                                    var aFieldKey = aFieldName.Trim().ToUpper();
                                    string aFieldValue = null;
                                    if (fieldMap.ContainsKey(aFieldKey))
                                    {
                                        if (null != fieldMap[aFieldKey])
                                        {
                                            aFieldValue = fieldMap[aFieldKey].Trim();
                                        }
                                    }
                                    if (String.IsNullOrEmpty(aFieldValue)) aFieldValue = new string(" ");

                                    if (t != null)
                                    {
                                        if (t.Text != null && aFieldKey != null && fieldMap.ContainsKey(aFieldKey))
                                        {
                                            if (!fieldMap.ContainsKey(aFieldKey) || (String.IsNullOrWhiteSpace(aFieldValue)
                                            || aFieldValue.Contains("BOOKMARK_UNDEFINED")))
                                            {
                                                Run rBegin2 = null;
                                                Run rBegin1 = null;
                                                Run rBegin0 = null;
                                                Paragraph rParent = null;
                                                Run rParentFirst = null;
                                                Run rParentLeft = null;
                                                Run rParentLeftFirst = null;
                                                if (null != rBegin) rBegin2 = rBegin.PreviousSibling<Run>();
                                                if (null != rBegin2) rBegin1 = rBegin2.PreviousSibling<Run>();
                                                if (null != rBegin1) rBegin0 = rBegin1.PreviousSibling<Run>();
                                                if (null != rText) rParent = (Paragraph)rText.Parent;
                                                /*if (null != rParent) rParentFirst = (Run)rParent.GetFirstChild<Run>();
                                                if (null != rParent) rParentLeft = rParent.PreviousSibling<Run>();
                                                if (null != rParentLeft) rParentLeftFirst = rParentLeft.GetFirstChild<Run>();*/
                                                //t.Text = " ";
                                                //if (null != rEnd) rAdd = rEnd.NextSibling<Run>();
                                                //if (null != rAdd)
                                                //    rAdd.AppendChild(new Text(" "));
                                                //rAdd.AppendChild(new Text(fieldMap[aFieldKey]));*/
                                                //if (null != rParent) rParent.InnerXml.Replace(fieldId, "");
                                                //if (null != rText) rText.RemoveAllChildren();
                                                //if (null != rText) rText.Remove();
                                                if (null != rEnd) rEnd.Remove();
                                                if (null != rSep) rSep.Remove();
                                                if (null != rBegin) rBegin.Remove();
                                                if (null != t) t.Text = aFieldValue;
                                                if (null != xxxfield) xxxfield.Remove();
                                                //if (null != rParent) rParent.RemoveAllChildren();
                                                //if (null != rParent) rParent.AppendChild<Run>(new Run(new Text("")));                                        
                                                /*if (null != rBegin2) rBegin2.RemoveAllChildren();
                                                if (null != rBegin2) rBegin2.Remove();
                                                if (null != rBegin1) rBegin1.RemoveAllChildren();
                                                if (null != rBegin1) rBegin1.Remove();
                                                if (null != rBegin0) rBegin0.RemoveAllChildren();
                                                if (null != rBegin0) rBegin0.Remove();*/
                                                //if (null != t.Text) Console.Out.WriteLine("****Substitute value " + t.Text + "with " + fieldMap[aFieldKey]);
                                                //rText.Remove();
                                            }
                                            else
                                            {
                                                Console.Out.WriteLine("****Substitute value " + t.Text + "with " + aFieldValue);
                                                if (fieldMap.ContainsKey(aFieldKey) && !(String.IsNullOrEmpty(aFieldValue)))
                                                {



                                                    if (nonEditableArray.Contains(aFieldKey))
                                                    {
                                                        t.Text = formatText(aFieldValue);
                                                    }
                                                    else
                                                    {


                                                        Run rBegin2 = null;
                                                        Run rBegin1 = null;
                                                        Run rBegin0 = null;
                                                        Paragraph rParent = null;
                                                        Run rParentFirst = null;
                                                        Run rParentLeft = null;
                                                        Run rParentLeftFirst = null;
                                                        if (null != rBegin) rBegin2 = rBegin.PreviousSibling<Run>();
                                                        if (null != rBegin2) rBegin1 = rBegin2.PreviousSibling<Run>();
                                                        if (null != rBegin1) rBegin0 = rBegin1.PreviousSibling<Run>();
                                                        if (null != rText) rParent = (Paragraph)rText.Parent;
                                                        /*if (null != rParent) rParentFirst = (Run)rParent.GetFirstChild<Run>();
                                                        if (null != rParent) rParentLeft = rParent.PreviousSibling<Run>();
                                                        if (null != rParentLeft) rParentLeftFirst = rParentLeft.GetFirstChild<Run>();*/
                                                        //t.Text = fieldMap[aFieldKey];
                                                        //if (null != rEnd) rAdd = rEnd.NextSibling<Run>();
                                                        //if (null != rAdd)
                                                        //    rAdd.AppendChild(new Text(fieldMap[aFieldKey]));
                                                        //rAdd = rText.NextSibling<Run>();
                                                        /*if (null != rAdd)
                                                            rAdd.AppendChild(new Text(fieldMap[aFieldKey]));*/
                                                        //if (null != rParent) rParent.InnerXml.Replace(fieldId, "");
                                                        //if (null != rText) rText.RemoveAllChildren();
                                                        //if (null != rText) rText.Remove();

                                                        string aaFieldKey = null;
                                                        string aaFieldValue = null;
                                                        string aaaFieldKey = null;
                                                        string aaaFieldValue = null;

                                                        if (null != rEnd) nFormat = rEnd.NextSibling<Run>();
                                                        if (null != nFormat) nBegin = nFormat.NextSibling<Run>();
                                                        if (null != nBegin) nTag = nBegin.NextSibling<Run>();
                                                        if (null != nTag) nSep = nTag.NextSibling<Run>();
                                                        if (null != nSep) nText = nSep.NextSibling<Run>();
                                                        if (null != nText) nEnd = nText.NextSibling<Run>();
                                                        if (null != nText) nt = nText.GetFirstChild<Text>();
                                                        if (null != nEnd) nnFormat = nEnd.NextSibling<Run>();
                                                        if (null != nnFormat) nnBegin = nnFormat.NextSibling<Run>();
                                                        if (null != nnBegin) nnTag = nnBegin.NextSibling<Run>();
                                                        if (null != nnTag) nnSep = nnTag.NextSibling<Run>();
                                                        if (null != nnSep) nnText = nnSep.NextSibling<Run>();
                                                        if (null != nnText) nnEnd = nnText.NextSibling<Run>();
                                                        if (null != nnText) nnt = nnText.GetFirstChild<Text>();


                                                        if (null != nnt)
                                                        {
                                                            if (null != nnt.Text && !String.IsNullOrWhiteSpace(nnt.Text))
                                                            {
                                                                aaaFieldKey = nnt.Text.Trim().ToUpper();
                                                                if (null != aaaFieldKey && fieldMap.ContainsKey(aaaFieldKey)) aaaFieldValue = fieldMap[aaaFieldKey];
                                                                if (null != aaaFieldValue && !String.IsNullOrWhiteSpace(aaaFieldValue) && aaaFieldValue.ToUpper() != "BOOKMARK_UNDEFINED")
                                                                {
                                                                    nnt.Text = formatText(aaaFieldValue);
                                                                    if (null != nnEnd) nnEnd.Remove();
                                                                    if (null != nnSep) nnSep.Remove();
                                                                    if (null != nnTag) nnTag.Remove();
                                                                    if (null != nnBegin) nnBegin.Remove();
                                                                    //if (null != rText) rText = new Run(new Text(aFieldValue));

                                                                }
                                                            }
                                                        }

                                                        if (null != nt)
                                                        {
                                                            if (null != nt.Text && !String.IsNullOrWhiteSpace(nt.Text))
                                                            {
                                                                aaFieldKey = nt.Text.Trim().ToUpper();
                                                                if (null != aaFieldKey && fieldMap.ContainsKey(aaFieldKey)) aaFieldValue = fieldMap[aaFieldKey];
                                                                if (null != aaFieldValue && !String.IsNullOrWhiteSpace(aaFieldValue) && aaFieldValue.ToUpper() != "BOOKMARK_UNDEFINED")
                                                                {
                                                                    nt.Text = formatText(aaFieldValue);
                                                                    if (null != nEnd) nEnd.Remove();
                                                                    if (null != nSep) nSep.Remove();
                                                                    if (null != nTag) nTag.Remove();
                                                                    if (null != nBegin) nBegin.Remove();
                                                                }
                                                            }
                                                        }







                                                        if (null != rEnd) rEnd.Remove();
                                                        if (null != rSep) rSep.Remove();
                                                        if (null != rBegin) rBegin.Remove();
                                                        if (null != t && null != t.Text) t.Text = formatText(aFieldValue);
                                                        //if (null != rText) rText = new Run(new Text(aFieldValue));
                                                        if (null != xxxfield) xxxfield.Remove();
                                                        //if (null != rParent) rParent.RemoveAllChildren();
                                                        //if (null != rParent) rParent.AppendChild<Run>(new Run(new Text(aFieldValue)));
                                                        /*if (null != rBegin2) rBegin2.RemoveAllChildren();
                                                        if (null != rBegin2) rBegin2.Remove();
                                                        if (null != rBegin1) rBegin1.RemoveAllChildren();
                                                        if (null != rBegin1) rBegin1.Remove();
                                                        if (null != rBegin0) rBegin0.RemoveAllChildren();
                                                        if (null != rBegin0) rBegin0.Remove();*/


                                                    }
                                                }
                                            }

                                        }
                                    }
                                    else //field name is CHARFORMAT or something
                                    {
                                        if (null != rEnd)
                                        {

                                            if (null != rEnd.InnerText)
                                            {
                                                if (!(String.IsNullOrWhiteSpace(rEnd.InnerText)) && fieldMap.ContainsKey(rEnd.InnerText.ToUpper()))
                                                {
                                                    //rEnd.SetText(fieldMap[rEnd.InnerText]);
                                                    if (null != fieldMap[rEnd.InnerText.ToUpper()])
                                                        rEnd.SetText(fieldMap[rEnd.InnerText.ToUpper()]);
                                                    else
                                                        rEnd.SetText(formatText(null));
                                                    /*Run rBegin2 = null;
                                                    Run rBegin1 = null;
                                                    Run rBegin0 = null;
                                                    Run rParent = null;
                                                    Run rParentFirst = null;
                                                    Run rParentLeft = null;
                                                    Run rParentLeftFirst = null;                                        
                                                    if (null != rBegin) rBegin2 = rBegin.PreviousSibling<Run>();
                                                    if (null != rBegin2) rBegin1 = rBegin2.PreviousSibling<Run>();
                                                    if (null != rBegin1) rBegin0 = rBegin1.PreviousSibling<Run>();
                                                    if (null != rEnd) rParent = (Run)rEnd.Parent;
                                                    if (null != rParent) rParentFirst = (Run)rParent.GetFirstChild<Run>();
                                                    if (null != rParent) rParentLeft = rParent.PreviousSibling<Run>();
                                                    if (null != rParentLeft) rParentLeftFirst = rParentLeft.GetFirstChild<Run>();
                                                    rAdd = rEnd.NextSibling<Run>();
                                                    if (null != rAdd)
                                                        rAdd.AppendChild(new Text(fieldMap[rEnd.InnerText]));
                                                    if (null != rEnd) rEnd.RemoveAllChildren();
                                                    if (null != rEnd) rEnd.Remove();
                                                    if (null != rText) rText.RemoveAllChildren();
                                                    if (null != rText) rText.Remove();
                                                    if (null != rSep) rSep.RemoveAllChildren();
                                                    if (null != rSep) rSep.Remove();
                                                    if (null != rBegin) rBegin.RemoveAllChildren();
                                                    if (null != rBegin) rBegin.Remove();
                                                    if (null != rBegin2) rBegin2.RemoveAllChildren();
                                                    if (null != rBegin2) rBegin2.Remove();
                                                    if (null != rBegin1) rBegin1.RemoveAllChildren();
                                                    if (null != rBegin1) rBegin1.Remove();
                                                    if (null != rBegin0) rBegin0.RemoveAllChildren();
                                                    if (null != rBegin0) rBegin0.Remove();
                                                    if (null != rParentFirst) rParentFirst.RemoveAllChildren();
                                                    if (null != rParentFirst) rParentFirst.Remove();
                                                    if (null != rParentLeftFirst) rParentLeftFirst.RemoveAllChildren();
                                                    if (null != rParentLeftFirst) rParentLeftFirst.Remove();
                                                    if (null != xxxfield) xxxfield.RemoveAllChildren();
                                                    if (null != xxxfield) xxxfield.Remove();*/
                                                }
                                                else
                                                {
                                                    rEnd.SetText(formatText(null));
                                                }
                                            }


                                        }
                                        else
                                        {
                                            // rEnd = nuill rEnd.SetText(" ");
                                            /* if (rEnd != null)
                                             {
                                                 Run rObject1 = null;
                                                 Run rObject2 = null;
                                                 Run rObject3 = null;
                                                 rObject1 = rEnd.GetFirstChild<Run>();
                                                 if (null != rObject1) rObject2 = rObject1.NextSibling<Run>();
                                                 if (null != rObject2) rObject3 = rObject2.NextSibling<Run>();

                                                 if (null != rObject1) 
                                                 {
                                                     if (null != rObject1 && null != rObject1.GetTextElement)
                                                         rObject1.Remove();
                                                     if (null != rObject2 && null != rObject2.GetTextElement)
                                                         rObject2.Remove();
                                                     if (null != rObject3)
                                                 }
                                             }*/
                                            /*t.Text = "";
                                            Run rBegin2 = null;
                                            Run rBegin1 = null;
                                            Run rBegin0 = null;
                                            Run rParent = null;
                                            Run rParentFirst = null;
                                            Run rParentLeft = null;
                                            Run rParentLeftFirst = null;
                                            if (null != rBegin) rBegin2 = rBegin.PreviousSibling<Run>();
                                            if (null != rBegin2) rBegin1 = rBegin2.PreviousSibling<Run>();
                                            if (null != rBegin1) rBegin0 = rBegin1.PreviousSibling<Run>();
                                            if (null != rEnd) rParent = (Run)rEnd.Parent;
                                            if (null != rParent) rParentFirst = (Run)rParent.GetFirstChild<Run>();
                                            if (null != rParent) rParentLeft = rParent.PreviousSibling<Run>();
                                            if (null != rParentLeft) rParentLeftFirst = rParentLeft.GetFirstChild<Run>();
                                            rAdd = rEnd.NextSibling<Run>();                                        
                                            if (null != rAdd) rAdd.AppendChild(new Text(""));
                                            if (null != rEnd) rEnd.RemoveAllChildren();
                                            if (null != rEnd) rEnd.Remove();                                        
                                            if (null != rText) rText.RemoveAllChildren();
                                            if (null != rText) rText.Remove();
                                            if (null != rSep) rSep.RemoveAllChildren();
                                            if (null != rSep) rSep.Remove();
                                            if (null != rBegin) rBegin.RemoveAllChildren();
                                            if (null != rBegin) rBegin.Remove();
                                            if (null != rBegin2) rBegin2.RemoveAllChildren();
                                            if (null != rBegin2) rBegin2.Remove();
                                            if (null != rBegin1) rBegin1.RemoveAllChildren(); 
                                            if (null != rBegin1) rBegin1.Remove();
                                            if (null != rBegin0) rBegin0.RemoveAllChildren(); 
                                            if (null != rBegin0) rBegin0.Remove();
                                            if (null != rParentFirst) rParentFirst.RemoveAllChildren();
                                            if (null != rParentFirst) rParentFirst.Remove();
                                            if (null != rParentLeftFirst) rParentLeftFirst.RemoveAllChildren();
                                            if (null != rParentLeftFirst) rParentLeftFirst.Remove();
                                            if (null != xxxfield) xxxfield.RemoveAllChildren();
                                            if (null != xxxfield) xxxfield.Remove();*/

                                        }
                                    }
                                }
                                else
                                {

                                }
                            }
                            else
                            {
                                Console.Out.WriteLine("@@@ field value not found.");
                            }
                        }


                        Console.Out.WriteLine("*** Field " + j.ToString() + " ends *********************************");

                        prevRun = pivotRun;
                        prevBegin = pivotBegin;
                        prevEnd = pivotEnd;










                        /* DocumentProperty cField = (custom[fieldId]);
                         if (null != cField)
                         {
                             Console.Out.WriteLine(">>> " + cField.Name + ": " + cField.Value);
                         }*/




                        /*int fieldNameStart = field.Text.LastIndexOf(FieldDelimeter, System.StringComparison.Ordinal);

                        if (fieldNameStart >= 0)
                        {
                            var fieldName = field.Text.Substring(fieldNameStart + FieldDelimeter.Length).Trim();

                            Run xxxfield = (Run)field.Parent;

                            Run rBegin = xxxfield.PreviousSibling<Run>();
                            Run rSep = xxxfield.NextSibling<Run>();
                            Run rText = rSep.NextSibling<Run>();
                            Run rEnd = rText.NextSibling<Run>();

                            if (null != xxxfield)
                            {

                                Text t = rText.GetFirstChild<Text>();
                               //custom.CustomHash;
                               Console.Out.WriteLine(t.ToString());

                            }
                        }*/

                    }


                    /*MainDocumentPart mainPart = wordDocument.MainDocumentPart;
                    var fields = mainPart.Document.Body.Descendants<FieldCode>();

                    foreach (var field in fields)
                    {
                        //if (field.GetType() == typeof(FormFieldData))
                        //{

                            Console.Out.WriteLine("***"+ field.ToString());
                            Console.Out.WriteLine("***"+ field.GetType());
                        //Console.Out.WriteLine("***" + ((FieldCode)field.FirstChild).Val.InnerText);
                        if (((FieldCode)field.FirstChild).Val.InnerText.Equals("WorkerName"))
                            {
                                TextInput text = field.Descendants<TextInput>().First();
                                SetFormFieldValue(text, "Put some text inside the field");
                            }
                        //}
                    }*/

                    /*if (null != wordDocument)
                    {

                        string aFieldDelimeter = @" MERGEFIELD ";
                        List<string> alisteChamps = new List<string>();

                        //foreach (var footer in wordDocument.MainDocumentPart.Document)
                        //{

                            foreach (var field in wordDocument.MainDocumentPart.RootElement.Descendants<FieldCode>())
                            {

                                int fieldNameStart = field.Text.LastIndexOf(aFieldDelimeter, System.StringComparison.Ordinal);

                                if (fieldNameStart >= 0)
                                {
                                    var fieldName = field.Text.Substring(fieldNameStart + aFieldDelimeter.Length).Trim();
                                    Console.Out.WriteLine("******" + fieldName.ToString());

                                Run xxxfield = (Run)field.Parent;

                                    Run rBegin = xxxfield.PreviousSibling<Run>();
                                    Run rSep = xxxfield.NextSibling<Run>();
                                    Run rText = rSep.NextSibling<Run>();
                                    Run rEnd = rText.NextSibling<Run>();

                                    if (null != xxxfield)
                                    {

                                        Text t = rText.GetFirstChild<Text>();
                                    //t.Text = replacementText;
                                    Console.Out.WriteLine("*******" + t.Text.ToString());

                                    }
                                }

                            }


                        //}
                    }*/

                    setCustomProperty(wordDocument, WopiOptions.Value.ProcessedFlag, WopiOptions.Value.ApplicationName, CustomPropertyTypes.Text);

                    /*Remove VBA part
                        var docPart = wordDocument.MainDocumentPart;

                        // Look for the vbaProject part. If it is there, delete it.
                        var vbaPart = docPart.VbaProjectPart;
                        if (vbaPart != null)
                        {
                            // Delete the vbaProject part and then save the document.
                            docPart.DeletePart(vbaPart);
                            docPart.Document.Save();

                            // Change the document type to
                            // not macro-enabled.
                            wordDocument.ChangeDocumentType(WordprocessingDocumentType.Document);

                            // Track that the document has been changed.

                        }
                        //changeCompatibilityModeOfDocumentPart(wordDocument.MainDocumentPart);
                    */

                    wordDocument.Save();
                    wordDocument.Close();

                }
            } // end using

        }



        protected void runMacroY(string newFileName, bool restoreDocProtection = true)
        {
            Regex fileVer = new Regex(@"\s*(?<userID>)_(?<attachmentID>)_(?<versionNo>)\..*");
            var result = fileVer.Matches(newFileName);
            string userID = null;
            string attachmentID = null;
            string versionNo = null;

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(newFileName, true))
            {

                if (null != wordDocument)
                {

                    /*var properties = new Dictionary<string, string>();
                    foreach (var property in wordDocument.CustomFilePropertiesPart.Properties.Elements<DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty>())
                    {
                        properties.Add(property.Name, property.VTLPWSTR.Text);
                    }*/

                    try
                    {
                        ArrayList docProtClassList = new ArrayList();
                        ArrayList docProtTypeList = new ArrayList();
                        var docProt = getCustomProperty(wordDocument, WopiOptions.Value.DocumentProtectionFlag);
                        if (null != WopiOptions.Value.DocumentProtectionClass && WopiOptions.Value.DocumentProtectionClass.Length > 0)
                            docProtClassList.AddRange(WopiOptions.Value.DocumentProtectionClass);

                        if (null != WopiOptions.Value.DocumentProtectionType && WopiOptions.Value.DocumentProtectionType.Length > 0)
                            docProtTypeList.AddRange(WopiOptions.Value.DocumentProtectionType);

                        if (!String.IsNullOrWhiteSpace(docProt))
                        {
                            string[] docProtParams = docProt.Split(",");
                            string docProtClass = "edit";
                            string docProtType = "forms";
                            EnumValue<DocumentProtectionValues> editValue = DocumentProtectionValues.None;

                            if (null != docProtParams)
                            {
                                if (docProtParams.Length == 2)
                                {
                                    docProtClass = docProtParams[0];
                                    docProtType = docProtParams[1];
                                }
                            }

                            //if (docProtType != "forms")
                            //{
                            if (docProtType == "0") editValue = DocumentProtectionValues.None;
                            if (docProtType == "1") editValue = DocumentProtectionValues.ReadOnly;
                            if (docProtType == "2") editValue = DocumentProtectionValues.Comments;
                            if (docProtType == "3") editValue = DocumentProtectionValues.TrackedChanges;
                            if (docProtType == "4") editValue = DocumentProtectionValues.Forms;

                            //}

                            /*foreach (DocumentProtection dp in wordDocument.MainDocumentPart.DocumentSettingsPart.Settings.ChildElements.OfType<DocumentProtection>())
                            {
                                //dp.Remove();                           
                                < w:documentProtection w:edit = "forms"
                                  w: formatting = "1"
                                  w: enforcement = "1" />
                            
                            }*/


                            if (docProtClassList.Contains(docProtClass))
                            {

                                var dp = new DocumentProtection()
                                {
                                    Edit = editValue,
                                    Enforcement = new OnOffValue(true),
                                    Formatting = new OnOffValue(true)
                                    //CryptographicProviderType = CryptProviderValues.RsaFull,
                                    //CryptographicAlgorithmClass = CryptAlgorithmClassValues.Hash,
                                    //CryptographicAlgorithmType = CryptAlgorithmValues.TypeAny,
                                    //CryptographicAlgorithmSid = 4,
                                    //CryptographicSpinCount = 100000U,
                                    //Hash = "2krUoz1qWd0WBeXqVrOq81l8xpk=",
                                    //Salt = "9kIgmDDYtt2r5U2idCOwMA=="
                                };

                                if (null != dp)
                                    dp.Edit = editValue;

                                var dsp = wordDocument.MainDocumentPart.DocumentSettingsPart;

                                var oldDp = dsp.Settings.FirstOrDefault(s => s.GetType() == typeof(DocumentProtection));

                                if (oldDp == null)
                                {
                                    dsp.Settings.AppendChild(dp);
                                }
                                else
                                {
                                    dsp.Settings.ReplaceChild(dp, oldDp);
                                }

                                var docProtKey = Path.GetFileNameWithoutExtension(newFileName);
                                if (null != docProtKey)
                                {
                                    if (docProtection.ContainsKey(docProtKey))
                                    {
                                        docProtection[docProtKey] = TRUE;
                                    }
                                    else
                                    {
                                        docProtection.Add(docProtKey, TRUE);
                                    }

                                }

                            }


                            //wordDocument.MainDocumentPart.DocumentSettingsPart.Settings.ChildElements.OfType<DocumentProtection>()

                        }
                        else
                        {
                            var docProtKey = Path.GetFileNameWithoutExtension(newFileName);
                            if (docProtection.ContainsKey(docProtKey)) docProtection.Remove(docProtKey);
                        }


                        //setCustomProperty(wordDocument, WopiOptions.Value.ProcessedFlag, WopiOptions.Value.ApplicationName, CustomPropertyTypes.Text);

                        /*Remove VBA part
                            var docPart = wordDocument.MainDocumentPart;

                            // Look for the vbaProject part. If it is there, delete it.
                            var vbaPart = docPart.VbaProjectPart;
                            if (vbaPart != null)
                            {
                                // Delete the vbaProject part and then save the document.
                                docPart.DeletePart(vbaPart);
                                docPart.Document.Save();

                                // Change the document type to
                                // not macro-enabled.
                                wordDocument.ChangeDocumentType(WordprocessingDocumentType.Document);

                                // Track that the document has been changed.

                            }
                            //changeCompatibilityModeOfDocumentPart(wordDocument.MainDocumentPart);
                        */
                    }
                    catch (Exception x)
                    {
                        Console.Out.WriteLine("Problem occurred in restoring document protection.");
                    }

                    wordDocument.Save();
                    wordDocument.Close();

                }
            } // end using

        }


        protected void runMacroZ(string newFileName)
        {
            Regex fileVer = new Regex(@"\s*(?<userID>)_(?<attachmentID>)_(?<versionNo>)\..*");
            var result = fileVer.Matches(newFileName);
            string userID = null;
            string attachmentID = null;
            string versionNo = null;

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(newFileName, true))
            {

                if (null != wordDocument)
                {

                    /*var properties = new Dictionary<string, string>();
                    foreach (var property in wordDocument.CustomFilePropertiesPart.Properties.Elements<DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty>())
                    {
                        properties.Add(property.Name, property.VTLPWSTR.Text);
                    }*/

                    try
                    {
                        var docProt = getCustomProperty(wordDocument, WopiOptions.Value.DocumentProtectionFlag);

                        if (!String.IsNullOrWhiteSpace(docProt))
                        {
                            string[] docProtParams = docProt.Split(",");
                            string docProtClass = "edit";
                            string docProtType = "forms";

                            if (null != docProtParams)
                            {
                                if (docProtParams.Length == 2)
                                {
                                    docProtClass = docProtParams[0];
                                    docProtType = docProtParams[1];
                                }
                            }

                            /*foreach (DocumentProtection dp in wordDocument.MainDocumentPart.DocumentSettingsPart.Settings.ChildElements.OfType<DocumentProtection>())
                            {
                                //dp.Remove();                           
                                < w:documentProtection w:edit = "forms"
                                  w: formatting = "1"
                                  w: enforcement = "1" />
                            
                            }*/


                            if (docProtClass != null)
                            {

                                var dp = new DocumentProtection()
                                {
                                    Edit = DocumentProtectionValues.Forms,
                                    Enforcement = true,
                                    Formatting = true
                                    //CryptographicProviderType = CryptProviderValues.RsaFull,
                                    //CryptographicAlgorithmClass = CryptAlgorithmClassValues.Hash,
                                    //CryptographicAlgorithmType = CryptAlgorithmValues.TypeAny,
                                    //CryptographicAlgorithmSid = 4,
                                    //CryptographicSpinCount = 100000U,
                                    //Hash = "2krUoz1qWd0WBeXqVrOq81l8xpk=",
                                    //Salt = "9kIgmDDYtt2r5U2idCOwMA=="
                                };

                                var dsp = wordDocument.MainDocumentPart.DocumentSettingsPart;

                                var oldDp = dsp.Settings.FirstOrDefault(s => s.GetType() == typeof(DocumentProtection));

                                if (oldDp == null)
                                {
                                    dsp.Settings.AppendChild(dp);
                                }
                                else
                                {
                                    dsp.Settings.ReplaceChild(dp, oldDp);
                                }

                            }


                            //wordDocument.MainDocumentPart.DocumentSettingsPart.Settings.ChildElements.OfType<DocumentProtection>()

                        }


                        //setCustomProperty(wordDocument, WopiOptions.Value.ProcessedFlag, WopiOptions.Value.ApplicationName, CustomPropertyTypes.Text);

                        /*Remove VBA part
                            var docPart = wordDocument.MainDocumentPart;

                            // Look for the vbaProject part. If it is there, delete it.
                            var vbaPart = docPart.VbaProjectPart;
                            if (vbaPart != null)
                            {
                                // Delete the vbaProject part and then save the document.
                                docPart.DeletePart(vbaPart);
                                docPart.Document.Save();

                                // Change the document type to
                                // not macro-enabled.
                                wordDocument.ChangeDocumentType(WordprocessingDocumentType.Document);

                                // Track that the document has been changed.

                            }
                            //changeCompatibilityModeOfDocumentPart(wordDocument.MainDocumentPart);
                        */
                    }
                    catch (Exception x)
                    {
                        Console.Out.WriteLine("Problem occurred in restoring document protection.");
                    }

                    wordDocument.Save();
                    wordDocument.Close();

                }
            } // end using

        }



        public void runMacro(string newFileName, Dictionary<string, string> fieldMap, bool processedByWWA = false, bool normalize = false)
        {
            Regex fileVer = new Regex(@"\s*(?<userID>)_(?<attachmentID>)_(?<versionNo>)\..*");
            var result = fileVer.Matches(newFileName);
            string userID = null;
            string attachmentID = null;
            string versionNo = null;

            /*if (Int32.TryParse(versionNo, out int v))
            {
                if (v > 1) return;
            }*/
            //if (processedByWWA) return;

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(newFileName, true))
            {

                if (null != wordDocument)
                {
                    /*for (var field in wordDocument.MainDocumentPart.RootElement.Descendants<FieldCode>())
                    {

                    }*/

                    //const string FieldDelimeter = @" MERGEFIELD ";
                    string FieldDelimeter = @" DOCPROPERTY ";
                    List<string> listeChamps = new List<string>();


                    if (normalize)
                    {
                        normalizeMarkup(wordDocument);
                        normalizeFieldCodesRuns(wordDocument);
                    }

                    Run prevRun = null;
                    Run prevBegin = null;
                    Run prevEnd = null;


                    int j = 0;
                    foreach (var field in wordDocument.MainDocumentPart.RootElement.Descendants<FieldCode>())
                    {
                        j++;
                        Console.Out.WriteLine("*** Field " + j.ToString() + " starts *********************************");
                        Console.Out.WriteLine("Field Type:" + field.ToString());
                        Console.Out.WriteLine("Field Text:>>>" + field.Text.ToString() + "<<<");
                        var fieldId = field.Text.Trim().ToString();
                        if (null != fieldId) Console.Out.WriteLine("fieldId:" + fieldId.ToString());

                        /*if (field.Text.ToString() == " DOCPROPERTY  WorkerName  \\* CHARFORMAT ")
                         {
                             field.Text = new string("REPLACED WORKER NAME");


                             Console.Out.WriteLine("Replaceing Field Text:>>>" + field.Text.ToString() + "<<<");
                             Console.Out.WriteLine("with:>>>REPLACED_WORKER_NAME<<<");
                         }*/

                        Run xxxfield = null;
                        Run rBegin = null;
                        Run rSep = null;
                        Run rText = null;
                        Run rEnd = null;
                        Text t = null;
                        Run pivotRun = null;
                        Run pivotEnd = null;
                        Run pivotBegin = null;
                        Run rAdd = null;

                        xxxfield = (Run)field.Parent;
                        if (null != xxxfield) rBegin = xxxfield.PreviousSibling<Run>();
                        if (null != xxxfield) rSep = xxxfield.NextSibling<Run>();
                        if (null != rSep) rText = rSep.NextSibling<Run>();
                        if (null != rText) rEnd = rText.NextSibling<Run>();
                        if (null != rText) t = rText.GetFirstChild<Text>();

                        pivotRun = xxxfield;
                        pivotBegin = rBegin;
                        pivotEnd = rEnd;
                        bool found = false;

                        Console.Out.WriteLine("@@@ Checking field Id...");
                        if (null != xxxfield) Console.Out.WriteLine("xxxfield:" + xxxfield.ToString() + ">>>" + xxxfield.InnerText);
                        if (null != rBegin) Console.Out.WriteLine("rBegin:" + rBegin.ToString() + ">>>" + rBegin.InnerText);
                        if (null != rText) Console.Out.WriteLine("rText:" + rText.ToString() + ">>>" + rText.InnerText);
                        if (null != rEnd) Console.Out.WriteLine("rEnd:" + rEnd.ToString() + ">>>" + rEnd.InnerText);
                        if (null != t) Console.Out.WriteLine("t:" + t.ToString() + ">>>" + t.InnerText);


                        //t.SetText("Vincent");

                        if (null == t || null == t.InnerText || String.IsNullOrWhiteSpace(t.InnerText))
                        //if (!fieldId.Contains("DOCPROPERTY"))
                        {
                            // check the previous sibling to see if it is there
                            Console.Out.WriteLine("@@@ Checking previous sibling...");
                            xxxfield = rBegin;
                            if (null != xxxfield) rBegin = xxxfield.PreviousSibling<Run>();
                            if (null != xxxfield) rSep = xxxfield.NextSibling<Run>();
                            if (null != rSep) rText = rSep.NextSibling<Run>();
                            if (null != rText) rEnd = rText.NextSibling<Run>();
                            if (null != rText) t = rText.GetFirstChild<Text>();
                            if (null != t)
                            {
                                if (null != t.Text && String.IsNullOrWhiteSpace(t.Text))
                                {
                                    Console.Out.WriteLine("@@@ Not found in previous sibling...");
                                }
                                else
                                {
                                    if (null != t.Text && t.Text.Length > 0)
                                    {
                                        Console.Out.WriteLine("@@@ Found in previous sibling... >>>" + t.Text + "<<<");
                                        if (null != xxxfield) Console.Out.WriteLine("xxxfield:" + xxxfield.ToString() + ">>>" + xxxfield.InnerText);
                                        if (null != rBegin) Console.Out.WriteLine("rBegin:" + rBegin.ToString() + ">>>" + rBegin.InnerText);
                                        if (null != rText) Console.Out.WriteLine("rText:" + rText.ToString() + ">>>" + rText.InnerText);
                                        if (null != rEnd) Console.Out.WriteLine("rEnd:" + rEnd.ToString() + ">>>" + rEnd.InnerText);
                                        if (null != t) Console.Out.WriteLine("t:" + t.ToString() + ">>>" + t.InnerText);
                                        found = true;
                                    }
                                }
                            }
                            // t is null
                            // not found in previous

                            //(t is null)
                            //check next
                            if (!found)
                            {
                                Console.Out.WriteLine("@@@ Checking next sibling...");
                                xxxfield = pivotEnd;
                                if (null != xxxfield) rBegin = xxxfield.PreviousSibling<Run>();
                                if (null != xxxfield) rSep = xxxfield.NextSibling<Run>();
                                if (null != rSep) rText = rSep.NextSibling<Run>();
                                if (null != rText) rEnd = rText.NextSibling<Run>();
                                if (null != rText) t = rText.GetFirstChild<Text>();
                                if (null != t)
                                {
                                    if (null != t.Text && String.IsNullOrWhiteSpace(t.Text))
                                    {
                                        Console.Out.WriteLine("@@@ Not found in next sibling... giving up");
                                    }
                                    else
                                    {
                                        if (null != t.Text && t.Text.Length > 0)
                                        {
                                            Console.Out.WriteLine("@@@ Found in next sibling...>>>" + t.Text + "<<<");
                                            if (null != xxxfield) Console.Out.WriteLine("xxxfield:" + xxxfield.ToString() + ">>>" + xxxfield.InnerText);
                                            if (null != rBegin) Console.Out.WriteLine("rBegin:" + rBegin.ToString() + ">>>" + rBegin.InnerText);
                                            if (null != rText) Console.Out.WriteLine("rText:" + rText.ToString() + ">>>" + rText.InnerText);
                                            if (null != rEnd) Console.Out.WriteLine("rEnd:" + rEnd.ToString() + ">>>" + rEnd.InnerText);
                                            if (null != t) Console.Out.WriteLine("t:" + t.ToString() + ">>>" + t.InnerText);
                                            found = true;
                                        }

                                    }
                                }
                            }
                        }
                        else
                        {
                            found = true;
                        }


                        if (!fieldId.Contains("DOCPROPERTY"))
                        {
                            if (null != rBegin && null != rBegin.InnerText && rBegin.InnerText.Contains("DOCPROPERTY"))
                            {
                                fieldId = rBegin.InnerText.Trim();
                            }
                            else
                            {
                                if (null != rEnd && null != rEnd.InnerText && rEnd.InnerText.Contains("DOCPROPERTY"))
                                    fieldId = rEnd.InnerText.Trim();
                                else
                                {
                                    if (null != prevBegin && null != prevBegin.InnerText && prevBegin.InnerText.Contains("DOCPROPERTY"))
                                    {
                                        if (null != t && null != t.Text && !String.IsNullOrWhiteSpace(t.Text))
                                        {
                                            Console.Out.WriteLine("##### " + prevBegin.InnerText);
                                            fieldId = prevBegin.InnerText.Trim();
                                        }

                                    }
                                }

                            }
                        }





                        if (found)
                        {

                            Regex expr = new Regex(@"\s*(?<docProperty>\S+)\s+(?<aFieldName>\S+)\.*\s+(?<formatType>\S+)\s*");
                            var results = expr.Matches(fieldId);
                            string docProperty = null;
                            string aFieldName = null;
                            string formatType = null;

                            foreach (Match match in results)
                            {
                                docProperty = match.Groups["docProperty"].Value;
                                aFieldName = match.Groups["aFieldName"].Value;
                                formatType = match.Groups["formatType"].Value;
                            }

                            if (null != docProperty) Console.Out.WriteLine("docProperty=" + docProperty);
                            if (null != aFieldName) Console.Out.WriteLine("aFieldName=" + aFieldName);
                            if (null != formatType) Console.Out.WriteLine("formatType=" + formatType);

                            if (null != aFieldName)
                            {
                                var aFieldKey = aFieldName.Trim().ToUpper();

                                if (t != null && t.Text != null && aFieldName != null && fieldMap.ContainsKey(aFieldKey))
                                {
                                    if (!fieldMap.ContainsKey(aFieldKey) || (String.IsNullOrWhiteSpace(fieldMap[aFieldKey])
                                    || fieldMap[aFieldKey].Contains("BOOKMARK_UNDEFINED")))
                                    {
                                        Run rBegin2 = null;
                                        Run rBegin1 = null;
                                        Run rBegin0 = null;
                                        Paragraph rParent = null;
                                        Run rParentFirst = null;
                                        Run rParentLeft = null;
                                        Run rParentLeftFirst = null;
                                        if (null != rBegin) rBegin2 = rBegin.PreviousSibling<Run>();
                                        if (null != rBegin2) rBegin1 = rBegin2.PreviousSibling<Run>();
                                        if (null != rBegin1) rBegin0 = rBegin1.PreviousSibling<Run>();
                                        if (null != rText) rParent = (Paragraph)rText.Parent;
                                        /*if (null != rParent) rParentFirst = (Run)rParent.GetFirstChild<Run>();
                                        if (null != rParent) rParentLeft = rParent.PreviousSibling<Run>();
                                        if (null != rParentLeft) rParentLeftFirst = rParentLeft.GetFirstChild<Run>();*/
                                        t.Text = "";
                                        rAdd = rText.NextSibling<Run>();
                                        /*if (null != rAdd)
                                            rAdd.AppendChild(new Text(fieldMap[aFieldKey]));*/
                                        //if (null != rParent) rParent.InnerXml.Replace(fieldId, "");
                                        //if (null != rText) rText.RemoveAllChildren();
                                        //if (null != rText) rText.Remove();
                                        //if (null != rSep) rSep.Remove();
                                        //if (null != rBegin) rBegin.Remove();
                                        //if (null != rParent) rParent.RemoveAllChildren();
                                        //if (null != rParent) rParent.AppendChild<Run>(new Run(new Text(fieldMap[aFieldKey])));
                                        /*if (null != rBegin2) rBegin2.RemoveAllChildren();
                                        if (null != rBegin2) rBegin2.Remove();
                                        if (null != rBegin1) rBegin1.RemoveAllChildren();
                                        if (null != rBegin1) rBegin1.Remove();
                                        if (null != rBegin0) rBegin0.RemoveAllChildren();
                                        if (null != rBegin0) rBegin0.Remove();*/
                                        //if (null != t.Text) Console.Out.WriteLine("****Substitute value " + t.Text + "with " + fieldMap[aFieldKey]);
                                        //rText.Remove();

                                    }
                                    else
                                    {
                                        Console.Out.WriteLine("****Substitute value " + t.Text + "with " + fieldMap[aFieldKey]);
                                        if (null != t.Text)
                                        {
                                            Run rBegin2 = null;
                                            Run rBegin1 = null;
                                            Run rBegin0 = null;
                                            Paragraph rParent = null;
                                            Run rParentFirst = null;
                                            Run rParentLeft = null;
                                            Run rParentLeftFirst = null;
                                            if (null != rBegin) rBegin2 = rBegin.PreviousSibling<Run>();
                                            if (null != rBegin2) rBegin1 = rBegin2.PreviousSibling<Run>();
                                            if (null != rBegin1) rBegin0 = rBegin1.PreviousSibling<Run>();
                                            if (null != rText) rParent = (Paragraph)rText.Parent;
                                            /*if (null != rParent) rParentFirst = (Run)rParent.GetFirstChild<Run>();
                                            if (null != rParent) rParentLeft = rParent.PreviousSibling<Run>();
                                            if (null != rParentLeft) rParentLeftFirst = rParentLeft.GetFirstChild<Run>();*/
                                            t.Text = fieldMap[aFieldKey];
                                            rAdd = rText.NextSibling<Run>();
                                            /*if (null != rAdd)
                                                rAdd.AppendChild(new Text(fieldMap[aFieldKey]));*/
                                            //if (null != rParent) rParent.InnerXml.Replace(fieldId, "");
                                            //if (null != rText) rText.RemoveAllChildren();
                                            //if (null != rText) rText.Remove();
                                            //if (null != rSep) rSep.Remove();
                                            //if (null != rBegin) rBegin.Remove();
                                            //if (null != rParent) rParent.RemoveAllChildren();
                                            //if (null != rParent) rParent.AppendChild<Run>(new Run(new Text(fieldMap[aFieldKey]+"\r\n")));
                                            /*if (null != rBegin2) rBegin2.RemoveAllChildren();
                                            if (null != rBegin2) rBegin2.Remove();
                                            if (null != rBegin1) rBegin1.RemoveAllChildren();
                                            if (null != rBegin1) rBegin1.Remove();
                                            if (null != rBegin0) rBegin0.RemoveAllChildren();
                                            if (null != rBegin0) rBegin0.Remove();*/

                                        }
                                    }
                                }
                                else //field name is CHARFORMAT or something
                                {
                                    if (null != rEnd && null != rEnd.InnerText && !(String.IsNullOrWhiteSpace(rEnd.InnerText)) && fieldMap.ContainsKey(rEnd.InnerText.ToUpper()))
                                    {
                                        //rEnd.SetText(fieldMap[rEnd.InnerText]);
                                        rEnd.SetText(fieldMap[rEnd.InnerText.ToUpper()]);
                                        /*Run rBegin2 = null;
                                        Run rBegin1 = null;
                                        Run rBegin0 = null;
                                        Run rParent = null;
                                        Run rParentFirst = null;
                                        Run rParentLeft = null;
                                        Run rParentLeftFirst = null;                                        
                                        if (null != rBegin) rBegin2 = rBegin.PreviousSibling<Run>();
                                        if (null != rBegin2) rBegin1 = rBegin2.PreviousSibling<Run>();
                                        if (null != rBegin1) rBegin0 = rBegin1.PreviousSibling<Run>();
                                        if (null != rEnd) rParent = (Run)rEnd.Parent;
                                        if (null != rParent) rParentFirst = (Run)rParent.GetFirstChild<Run>();
                                        if (null != rParent) rParentLeft = rParent.PreviousSibling<Run>();
                                        if (null != rParentLeft) rParentLeftFirst = rParentLeft.GetFirstChild<Run>();
                                        rAdd = rEnd.NextSibling<Run>();
                                        if (null != rAdd)
                                            rAdd.AppendChild(new Text(fieldMap[rEnd.InnerText]));
                                        if (null != rEnd) rEnd.RemoveAllChildren();
                                        if (null != rEnd) rEnd.Remove();
                                        if (null != rText) rText.RemoveAllChildren();
                                        if (null != rText) rText.Remove();
                                        if (null != rSep) rSep.RemoveAllChildren();
                                        if (null != rSep) rSep.Remove();
                                        if (null != rBegin) rBegin.RemoveAllChildren();
                                        if (null != rBegin) rBegin.Remove();
                                        if (null != rBegin2) rBegin2.RemoveAllChildren();
                                        if (null != rBegin2) rBegin2.Remove();
                                        if (null != rBegin1) rBegin1.RemoveAllChildren();
                                        if (null != rBegin1) rBegin1.Remove();
                                        if (null != rBegin0) rBegin0.RemoveAllChildren();
                                        if (null != rBegin0) rBegin0.Remove();
                                        if (null != rParentFirst) rParentFirst.RemoveAllChildren();
                                        if (null != rParentFirst) rParentFirst.Remove();
                                        if (null != rParentLeftFirst) rParentLeftFirst.RemoveAllChildren();
                                        if (null != rParentLeftFirst) rParentLeftFirst.Remove();
                                        if (null != xxxfield) xxxfield.RemoveAllChildren();
                                        if (null != xxxfield) xxxfield.Remove();*/


                                    }
                                    else
                                    {
                                        rEnd.SetText(" ");
                                        /* if (rEnd != null)
                                         {
                                             Run rObject1 = null;
                                             Run rObject2 = null;
                                             Run rObject3 = null;
                                             rObject1 = rEnd.GetFirstChild<Run>();
                                             if (null != rObject1) rObject2 = rObject1.NextSibling<Run>();
                                             if (null != rObject2) rObject3 = rObject2.NextSibling<Run>();

                                             if (null != rObject1) 
                                             {
                                                 if (null != rObject1 && null != rObject1.GetTextElement)
                                                     rObject1.Remove();
                                                 if (null != rObject2 && null != rObject2.GetTextElement)
                                                     rObject2.Remove();
                                                 if (null != rObject3)
                                             }
                                         }*/
                                        /*t.Text = "";
                                        Run rBegin2 = null;
                                        Run rBegin1 = null;
                                        Run rBegin0 = null;
                                        Run rParent = null;
                                        Run rParentFirst = null;
                                        Run rParentLeft = null;
                                        Run rParentLeftFirst = null;
                                        if (null != rBegin) rBegin2 = rBegin.PreviousSibling<Run>();
                                        if (null != rBegin2) rBegin1 = rBegin2.PreviousSibling<Run>();
                                        if (null != rBegin1) rBegin0 = rBegin1.PreviousSibling<Run>();
                                        if (null != rEnd) rParent = (Run)rEnd.Parent;
                                        if (null != rParent) rParentFirst = (Run)rParent.GetFirstChild<Run>();
                                        if (null != rParent) rParentLeft = rParent.PreviousSibling<Run>();
                                        if (null != rParentLeft) rParentLeftFirst = rParentLeft.GetFirstChild<Run>();
                                        rAdd = rEnd.NextSibling<Run>();                                        
                                        if (null != rAdd) rAdd.AppendChild(new Text(""));
                                        if (null != rEnd) rEnd.RemoveAllChildren();
                                        if (null != rEnd) rEnd.Remove();                                        
                                        if (null != rText) rText.RemoveAllChildren();
                                        if (null != rText) rText.Remove();
                                        if (null != rSep) rSep.RemoveAllChildren();
                                        if (null != rSep) rSep.Remove();
                                        if (null != rBegin) rBegin.RemoveAllChildren();
                                        if (null != rBegin) rBegin.Remove();
                                        if (null != rBegin2) rBegin2.RemoveAllChildren();
                                        if (null != rBegin2) rBegin2.Remove();
                                        if (null != rBegin1) rBegin1.RemoveAllChildren(); 
                                        if (null != rBegin1) rBegin1.Remove();
                                        if (null != rBegin0) rBegin0.RemoveAllChildren(); 
                                        if (null != rBegin0) rBegin0.Remove();
                                        if (null != rParentFirst) rParentFirst.RemoveAllChildren();
                                        if (null != rParentFirst) rParentFirst.Remove();
                                        if (null != rParentLeftFirst) rParentLeftFirst.RemoveAllChildren();
                                        if (null != rParentLeftFirst) rParentLeftFirst.Remove();
                                        if (null != xxxfield) xxxfield.RemoveAllChildren();
                                        if (null != xxxfield) xxxfield.Remove();*/

                                    }
                                }
                            }
                        }
                        else
                        {
                            Console.Out.WriteLine("@@@ field value not found.");
                        }


                        Console.Out.WriteLine("*** Field " + j.ToString() + " ends *********************************");

                        prevRun = pivotRun;
                        prevBegin = pivotBegin;
                        prevEnd = pivotEnd;










                        /* DocumentProperty cField = (custom[fieldId]);
                         if (null != cField)
                         {
                             Console.Out.WriteLine(">>> " + cField.Name + ": " + cField.Value);
                         }*/




                        /*int fieldNameStart = field.Text.LastIndexOf(FieldDelimeter, System.StringComparison.Ordinal);

                        if (fieldNameStart >= 0)
                        {
                            var fieldName = field.Text.Substring(fieldNameStart + FieldDelimeter.Length).Trim();

                            Run xxxfield = (Run)field.Parent;

                            Run rBegin = xxxfield.PreviousSibling<Run>();
                            Run rSep = xxxfield.NextSibling<Run>();
                            Run rText = rSep.NextSibling<Run>();
                            Run rEnd = rText.NextSibling<Run>();

                            if (null != xxxfield)
                            {

                                Text t = rText.GetFirstChild<Text>();
                               //custom.CustomHash;
                               Console.Out.WriteLine(t.ToString());

                            }
                        }*/

                    }


                    /*MainDocumentPart mainPart = wordDocument.MainDocumentPart;
                    var fields = mainPart.Document.Body.Descendants<FieldCode>();

                    foreach (var field in fields)
                    {
                        //if (field.GetType() == typeof(FormFieldData))
                        //{

                            Console.Out.WriteLine("***"+ field.ToString());
                            Console.Out.WriteLine("***"+ field.GetType());
                        //Console.Out.WriteLine("***" + ((FieldCode)field.FirstChild).Val.InnerText);
                        if (((FieldCode)field.FirstChild).Val.InnerText.Equals("WorkerName"))
                            {
                                TextInput text = field.Descendants<TextInput>().First();
                                SetFormFieldValue(text, "Put some text inside the field");
                            }
                        //}
                    }*/

                    /*if (null != wordDocument)
                    {

                        string aFieldDelimeter = @" MERGEFIELD ";
                        List<string> alisteChamps = new List<string>();

                        //foreach (var footer in wordDocument.MainDocumentPart.Document)
                        //{

                            foreach (var field in wordDocument.MainDocumentPart.RootElement.Descendants<FieldCode>())
                            {

                                int fieldNameStart = field.Text.LastIndexOf(aFieldDelimeter, System.StringComparison.Ordinal);

                                if (fieldNameStart >= 0)
                                {
                                    var fieldName = field.Text.Substring(fieldNameStart + aFieldDelimeter.Length).Trim();
                                    Console.Out.WriteLine("******" + fieldName.ToString());

                                Run xxxfield = (Run)field.Parent;

                                    Run rBegin = xxxfield.PreviousSibling<Run>();
                                    Run rSep = xxxfield.NextSibling<Run>();
                                    Run rText = rSep.NextSibling<Run>();
                                    Run rEnd = rText.NextSibling<Run>();

                                    if (null != xxxfield)
                                    {

                                        Text t = rText.GetFirstChild<Text>();
                                    //t.Text = replacementText;
                                    Console.Out.WriteLine("*******" + t.Text.ToString());

                                    }
                                }

                            }


                        //}
                    }*/

                    setCustomProperty(wordDocument, WopiOptions.Value.ProcessedFlag, WopiOptions.Value.ApplicationName, CustomPropertyTypes.Text);

                    /*Remove VBA part
                        var docPart = wordDocument.MainDocumentPart;

                        // Look for the vbaProject part. If it is there, delete it.
                        var vbaPart = docPart.VbaProjectPart;
                        if (vbaPart != null)
                        {
                            // Delete the vbaProject part and then save the document.
                            docPart.DeletePart(vbaPart);
                            docPart.Document.Save();

                            // Change the document type to
                            // not macro-enabled.
                            wordDocument.ChangeDocumentType(WordprocessingDocumentType.Document);

                            // Track that the document has been changed.

                        }
                        //changeCompatibilityModeOfDocumentPart(wordDocument.MainDocumentPart);
                    */

                    wordDocument.Save();
                    wordDocument.Close();

                }
            } // end using

        }


        public void convertDocToDocxSpire(string path, Boolean useTemp = false, string userRole = null)
        {
            //bool processedByWWA = false;
            //Dictionary<string, string> fieldMap = new Dictionary<string, string>();
            if (path.ToLower().EndsWith(WopiOptions.Value.Word2010Ext))
            {

                var sourceFile = new FileInfo(path);
                bool processedByWWA = false;
                string newFileName = sourceFile.FullName.Replace(WopiOptions.Value.Word2010Ext, WopiOptions.Value.WordExt);
                Dictionary<string, string> fieldMap = new Dictionary<string, string>();
                /*OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
                {
                    //Compliance = OoxmlCompliance.Iso29500_2008_Strict,
                    SaveFormat = SaveFormat.Docx
                };*/

                using (var document = new Document())
                {


                    if (!System.IO.File.Exists(newFileName))
                    {
                        System.IO.File.Delete(newFileName);

                        //ComponentInfo.SetLicense("FREE-LIMITED-KEY");

                        // In order to convert Word to PDF, we just need to:
                        // 1. Load DOC or DOCX file into DocumentModel object.
                        // 2. Save DocumentModel object to PDF file.
                        //DocumentModel document = DocumentModel.Load(sourceFile.FullName);
                        //document.Save(newFileName);
                        //Document document = new Document();
                        //document.LoadFromFile(sourceFile.FullName, FileFormat.WordXml);
                        document.LoadFromFile(sourceFile.FullName);

                        CustomDocumentProperties custom = document.CustomDocumentProperties.Clone();

                        Console.Out.WriteLine(sourceFile.FullName + ">>> " + (custom.Count).ToString());
                        for (int i = 0; i < custom.Count; i++)
                        {
                            Console.Out.WriteLine(">>>Field " + i + ">>>" + custom[i].ToString());
                        }

                        foreach (KeyValuePair<string, DocumentProperty> entry in custom.CustomHash)
                        {
                            DocumentProperty fields = entry.Value;

                            Console.Out.WriteLine(entry.GetType() + ">>>" + entry.Key + ">>>" + fields.Name + "<<<:>>>" + fields.Value + "<<<");
                            fieldMap.Add(fields.Name.ToUpper(), fields.Value.ToString());
                            if (null != fields.Name && fields.Name.ToUpper() == WopiOptions.Value.ProcessedFlag)
                            {
                                if (null != fields.Value && fields.Value.ToString() == WopiOptions.Value.ApplicationName)
                                {
                                    processedByWWA = true;
                                }
                            }
                        }

                        //document.SaveToFile(newFileName, FileFormat.Docm2010);
                        //document.ClearMacros();


                        document.SaveToFile(newFileName, FileFormat.Docm);
                        //document.Close();

                        //document.ClearMacros();
                        //document.SaveToFile(newFileName, FileFormat.Docx);


                        var preMacroFile = newFileName.Replace(WopiOptions.Value.WordExt, WopiOptions.Value.PreMacroSuffix) + WopiOptions.Value.WordExt;
                        document.SaveToFile(preMacroFile, FileFormat.Docm);
                        //document.SaveToFile(preMacroFile, FileFormat.Docx);
                        //document.SaveToFile(newFileName, FileFormat.Docm);


                        /*if (WopiOptions.Value.RunMacro.ToLower() == TRUE)
                        {
                            runMacroX(newFileName, fieldMap, processedByWWA, true, userRole);
                        }*/

                        if (WopiOptions.Value.RunMacro.ToLower() == TRUE)
                        {
                            var callRunMacro = "runMacro";

                            if (!String.IsNullOrEmpty(WopiOptions.Value.RunMacroVersion))
                            {
                                callRunMacro += WopiOptions.Value.RunMacroVersion;
                            }

                            MethodInfo runMacroMethod = this.GetType().GetMethod(callRunMacro);
                            object result = null;

                            try
                            {
                                runMacroMethod.Invoke(this, new object[] { newFileName, fieldMap, processedByWWA, true, userRole });
                            }
                            catch (Exception exc)
                            {
                                Console.Out.WriteLine("******** Error running macro method " + callRunMacro);
                            }

                        }

                        //document.SaveToFile(newFileName, FileFormat.Docm);

                        /*
                        document = new Document();
                        //document.LoadFromFile(sourceFile.FullName, FileFormat.WordXml);
                        document.LoadFromFile(newFileName);
                        document.SaveToFile(newFileName, FileFormat.OOXML);
                        */

                        //var newFileName_macro_removed = newFileName.Replace(Path.GetFileNameWithoutExtension(newFileName), Path.GetFileNameWithoutExtension(newFileName)+"_macro_removed");

                        /*
                        var newFileName_auto = newFileName.Replace(Path.GetFileNameWithoutExtension(newFileName), Path.GetFileNameWithoutExtension(newFileName) + "_auto");
                        var newFileName_doc = newFileName.Replace(Path.GetFileNameWithoutExtension(newFileName), Path.GetFileNameWithoutExtension(newFileName) + "_doc");
                        var newFileName_docm = newFileName.Replace(Path.GetFileNameWithoutExtension(newFileName), Path.GetFileNameWithoutExtension(newFileName) + "_docm");
                        var newFileName_docm2010 = newFileName.Replace(Path.GetFileNameWithoutExtension(newFileName), Path.GetFileNameWithoutExtension(newFileName) + "_docm2010");
                        var newFileName_docm2013 = newFileName.Replace(Path.GetFileNameWithoutExtension(newFileName), Path.GetFileNameWithoutExtension(newFileName) + "_docm2013");
                        var newFileName_docx = newFileName.Replace(Path.GetFileNameWithoutExtension(newFileName), Path.GetFileNameWithoutExtension(newFileName) + "_docx");
                        var newFileName_docx2010 = newFileName.Replace(Path.GetFileNameWithoutExtension(newFileName), Path.GetFileNameWithoutExtension(newFileName) + "_docx2010");
                        var newFileName_docx2013 = newFileName.Replace(Path.GetFileNameWithoutExtension(newFileName), Path.GetFileNameWithoutExtension(newFileName) + "_docx2013");
                        var newFileName_odt = newFileName.Replace(Path.GetFileNameWithoutExtension(newFileName), Path.GetFileNameWithoutExtension(newFileName) + "_odt");
                        var newFileName_ooxml = newFileName.Replace(Path.GetFileNameWithoutExtension(newFileName), Path.GetFileNameWithoutExtension(newFileName) + "_ooxml");
                        var newFileName_rtf = newFileName.Replace(Path.GetFileNameWithoutExtension(newFileName), Path.GetFileNameWithoutExtension(newFileName) + "_rtf");
                        var newFileName_wordml = newFileName.Replace(Path.GetFileNameWithoutExtension(newFileName), Path.GetFileNameWithoutExtension(newFileName) + "_wordml");
                        var newFileName_wordxml = newFileName.Replace(Path.GetFileNameWithoutExtension(newFileName), Path.GetFileNameWithoutExtension(newFileName) + "_wordxml");
                        var newFileName_html = newFileName.Replace(Path.GetFileNameWithoutExtension(newFileName), Path.GetFileNameWithoutExtension(newFileName) + "_html");
                        var newFileName_pdf = newFileName.Replace(Path.GetFileNameWithoutExtension(newFileName), Path.GetFileNameWithoutExtension(newFileName) + "_pdf");



                        //var newFileName3 = newFileName.Replace(Path.GetFileNameWithoutExtension(newFileName), Path.GetFileNameWithoutExtension(newFileName)+ "_");
                        //newFileName2 = newFileName2.Replace("docx", "doc");
                        document.SaveToFile(newFileName, FileFormat.OOXML);

                        /*
                        document.SaveToFile(newFileName_auto, FileFormat.Auto);
                        document.SaveToFile(newFileName_doc.Replace(".docx",".doc"), FileFormat.Doc);
                        document.SaveToFile(newFileName_docm.Replace(".docx", ".docm"), FileFormat.Docm);
                        document.SaveToFile(newFileName_docm2010.Replace(".docx", ".docm"), FileFormat.Docm2010);
                        document.SaveToFile(newFileName_docm2013.Replace(".docx",".docm"), FileFormat.Docm);
                        document.SaveToFile(newFileName_docx, FileFormat.Docx);
                        document.SaveToFile(newFileName_docx2010, FileFormat.Docx2010);
                        document.SaveToFile(newFileName_docx2013, FileFormat.Docx2013);
                        document.SaveToFile(newFileName_odt.Replace(".docx", ".odt"), FileFormat.Odt);
                        document.SaveToFile(newFileName_ooxml.Replace(".docx", ".doc"), FileFormat.OOXML);
                        document.SaveToFile(newFileName_rtf.Replace(".docx", ".rtf"), FileFormat.Rtf);
                        document.SaveToFile(newFileName_wordml.Replace(".docx", ".doc"), FileFormat.WordML);
                        document.SaveToFile(newFileName_wordxml, FileFormat.WordXml);
                        document.SaveToFile(newFileName_html.Replace(".docx",".htm"), FileFormat.Html);
                        document.SaveToFile(newFileName_pdf.Replace(".docx", ".pdf"), FileFormat.PDF);
                        */

                        //document.SaveToFile(newFileName_macro_removed, FileFormat.Docx2013);



                        var postEditFile = newFileName.Replace(WopiOptions.Value.WordExt, WopiOptions.Value.PostMacroSuffix) + WopiOptions.Value.WordExt;
                        if (System.IO.File.Exists(postEditFile)) System.IO.File.Delete(postEditFile);
                        System.IO.File.Copy(newFileName, postEditFile);

                        document.ClearMacros();
                        var newFileName_macro_removed = newFileName.Replace(Path.GetFileNameWithoutExtension(newFileName), Path.GetFileNameWithoutExtension(newFileName) + "_macro_removed");
                        document.SaveToFile(newFileName_macro_removed, FileFormat.Docx2013);

                        /*using (OpenXMLTemplates.Documents.TemplateDocument doc = new OpenXMLTemplates.Documents.TemplateDocument(newFileName2))
                        {

                            try
                            {
                                //System.IO.File.Copy(newFileName, newFileName2);
                                //using (OpenXMLTemplates.Documents.TemplateDocument doc = new OpenXMLTemplates.Documents.TemplateDocument(newFileName2)

                                //string data = System.IO.File.ReadAllText("macrodata.json");
                                //string data = System.IO.File.ReadAllText("macrodata.json");

                                string data = "{ \"WorkerName\": \"MYWORKER\" }";
                                var src = new VariableSource(data);

                                var engine = new DefaultOpenXmlTemplateEngine();

                                //Call the ReplaceAll method on the engine using the document and the variable source

                                engine.ReplaceAll(doc, src);


                                //src.LoadDataFromJson(data);

                                //var replacer = new VariableControlReplacer();

                                //replacer.ReplaceAll(doc, src);


                                doc.SaveAs(newFileName3);
                                //doc.Close();
                                //doc.Dispose();

                            }
                            catch (Exception x)
                            {
                                Console.Out.WriteLine("Errors in substitution");
                            }
                            finally
                            {
                                doc.Close();
                                doc.Dispose();
                            }
                        }*/




                        try
                        {
                            var current = DateTime.Now;
                            System.IO.File.SetCreationTime(newFileName, current);
                            System.IO.File.SetLastWriteTime(newFileName, current);
                            System.IO.File.SetLastAccessTime(newFileName, current);

                            Console.WriteLine(System.IO.File.GetCreationTime(newFileName));
                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }

                    }
                    else
                    {
                        //the file exists
                        if (!useTemp)
                        {
                            System.IO.File.Delete(newFileName);
                            //var document = word.Documents.Open(sourceFile.FullName);

                            //var project = document.VBProject;
                            //var module = project.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
                            // module.CodeModule.AddFromString("CMSUpdateFields");
                            //word.Run("CMSUpdateFields");
                            //word.ActiveDocument.ActiveWindow.View.ShowFieldCodes = false;
                            //word.ActiveDocument.ActiveWindow.View.ShowFieldCodes = false;
                            //word.ActiveDocument.StoryRanges.Fields.Update();
                            //word.ActiveDocument.StoryRanges.Fields.Update();


                            //document.SaveAs2(newFileName, WdSaveFormat.wdFormatXMLDocument,
                            //             CompatibilityMode: WdCompatibilityMode.wdWord2010);

                            //word.ActiveDocument.Close();
                            //word.Quit();

                            //ComponentInfo.SetLicense("FREE-LIMITED-KEY");

                            // In order to convert Word to PDF, we just need to:
                            // 1. Load DOC or DOCX file into DocumentModel object.
                            // 2. Save DocumentModel object to PDF file.
                            //DocumentModel document = DocumentModel.Load(sourceFile.FullName);
                            //document.Save(newFileName);
                            // Document document = new Document();
                            document.LoadFromFile(sourceFile.FullName);

                            CustomDocumentProperties custom = document.CustomDocumentProperties;

                            Console.Out.WriteLine(sourceFile.FullName + ">>> " + (custom.Count).ToString());
                            for (int i = 0; i < custom.Count; i++)
                            {
                                Console.Out.WriteLine(">>>Field " + i + ">>>" + custom[i].ToString());
                            }

                            foreach (KeyValuePair<string, DocumentProperty> entry in custom.CustomHash)
                            {
                                DocumentProperty fields = entry.Value;
                                Console.Out.WriteLine(entry.Key + ">>>" + fields.Name + "<<<:>>>" + fields.Value + "<<<");
                                fieldMap.Add(fields.Name.ToUpper(), fields.Value.ToString());


                                if (null != fields.Name && fields.Name.ToUpper() == WopiOptions.Value.ProcessedFlag)
                                {
                                    if (null != fields.Value && fields.Value.ToString() == WopiOptions.Value.ApplicationName)
                                    {
                                        processedByWWA = true;
                                    }
                                }
                            }


                            //document.SaveToFile(newFileName, FileFormat.Docm2010);
                            //document.ClearMacros();
                            //document.SaveToFile(newFileName, FileFormat.Docm);


                            //document.ClearMacros();
                            //document.SaveToFile(newFileName, FileFormat.Docx);


                            document.SaveToFile(newFileName, FileFormat.Docm);
                            //document.Close();

                            var preMacroFile = newFileName.Replace(WopiOptions.Value.WordExt, WopiOptions.Value.PreMacroSuffix) + WopiOptions.Value.WordExt;
                            document.SaveToFile(preMacroFile, FileFormat.Docm);
                            //document.SaveToFile(preMacroFile, FileFormat.Docx);


                            /* if (WopiOptions.Value.RunMacro.ToLower() == TRUE)
                            {
                                runMacroX(newFileName, fieldMap, processedByWWA, true, userRole);
                            } */
                            if (WopiOptions.Value.RunMacro.ToLower() == TRUE)
                            {
                                var callRunMacro = "runMacro";

                                if (!String.IsNullOrEmpty(WopiOptions.Value.RunMacroVersion))
                                {
                                    callRunMacro += WopiOptions.Value.RunMacroVersion;
                                }

                                MethodInfo runMacroMethod = this.GetType().GetMethod(callRunMacro);
                                object result = null;

                                try
                                {
                                    runMacroMethod.Invoke(this, new object[] { newFileName, fieldMap, processedByWWA, true, userRole });
                                }
                                catch (Exception exc)
                                {
                                    Console.Out.WriteLine("******** Error running macro method " + callRunMacro);
                                }

                            }


                            //runMacro(newFileName, fieldMap);


                            //var newFileName2 = newFileName.Replace(Path.GetFileNameWithoutExtension(newFileName), Path.GetFileNameWithoutExtension(newFileName) + "_2");
                            //var newFileName3 = newFileName.Replace(Path.GetFileNameWithoutExtension(newFileName), Path.GetFileNameWithoutExtension(newFileName) + "_3");
                            //newFileName2 = newFileName2.Replace("docx", "doc");
                            //document.SaveToFile(newFileName2, FileFormat.Docx2013);
                            //document.Close();


                            var postEditFile = newFileName.Replace(WopiOptions.Value.WordExt, WopiOptions.Value.PostMacroSuffix) + WopiOptions.Value.WordExt;
                            if (System.IO.File.Exists(postEditFile)) System.IO.File.Delete(postEditFile);
                            System.IO.File.Copy(newFileName, postEditFile);

                            document.ClearMacros();
                            var newFileName_macro_removed = newFileName.Replace(Path.GetFileNameWithoutExtension(newFileName), Path.GetFileNameWithoutExtension(newFileName) + "_macro_removed");
                            document.SaveToFile(newFileName_macro_removed, FileFormat.Docx2013);


                            try
                            {
                                var current = DateTime.Now;
                                System.IO.File.SetCreationTime(newFileName, current);
                                System.IO.File.SetLastWriteTime(newFileName, current);
                                System.IO.File.SetLastAccessTime(newFileName, current);

                                Console.WriteLine(System.IO.File.GetCreationTime(newFileName));
                            }
                            catch (Exception ex)
                            {
                                throw ex;
                            }

                        }
                        else
                        {
                            //leave it alone
                        }




                    } // else



                    var tempfile = path.Replace(WopiOptions.Value.Word2010Ext, WopiOptions.Value.RetreivedSuffix) + WopiOptions.Value.Word2010Ext;
                    if (System.IO.File.Exists(tempfile)) System.IO.File.Delete(tempfile);
                    System.IO.File.Move(path, tempfile);

                    //document.Close();


                }//using
            }
            else
            {
                //TODO
                ;
            }
        }

        /*public void convertDocToDocxSpire(string path, Boolean useTemp = false, string userRole = null)
        {
            bool processedByWWA = false;
            Dictionary<string, string> fieldMap = new Dictionary<string, string>();
            if (path.ToLower().EndsWith(WopiOptions.Value.Word2010Ext))
            {

                var sourceFile = new FileInfo(path);
                string newFileName = sourceFile.FullName.Replace(WopiOptions.Value.Word2010Ext, WopiOptions.Value.WordExt);

                if (!System.IO.File.Exists(newFileName))
                {
                    System.IO.File.Delete(newFileName);

                    //ComponentInfo.SetLicense("FREE-LIMITED-KEY");

                    // In order to convert Word to PDF, we just need to:
                    // 1. Load DOC or DOCX file into DocumentModel object.
                    // 2. Save DocumentModel object to PDF file.
                    //DocumentModel document = DocumentModel.Load(sourceFile.FullName);
                    //document.Save(newFileName);
                    Document document = new Document();
                    //document.LoadFromFile(sourceFile.FullName, FileFormat.WordXml);
                    document.LoadFromFile(sourceFile.FullName);

                    CustomDocumentProperties custom = document.CustomDocumentProperties.Clone();

                    Console.Out.WriteLine(sourceFile.FullName + ">>> " + (custom.Count).ToString());
                    for (int i = 0; i < custom.Count; i++)
                    {
                        Console.Out.WriteLine(">>>Field " + i + ">>>" + custom[i].ToString());
                    }

                    foreach (KeyValuePair<string, DocumentProperty> entry in custom.CustomHash)
                    {
                        DocumentProperty fields = entry.Value;

                        Console.Out.WriteLine(entry.GetType() + ">>>" + entry.Key + ">>>" + fields.Name + "<<<:>>>" + fields.Value + "<<<");
                        fieldMap.Add(fields.Name.ToUpper(), fields.Value.ToString());
                        if (null != fields.Name && fields.Name.ToUpper() == WopiOptions.Value.ProcessedFlag)
                        {
                            if (null != fields.Value && fields.Value.ToString() == WopiOptions.Value.ApplicationName)
                            {
                                processedByWWA = true;
                            }
                        }
                    }

                    //document.SaveToFile(newFileName, FileFormat.Docm2010);
                    //document.ClearMacros();
                    document.SaveToFile(newFileName, FileFormat.Docm);
                    //document.Close();

                    /*if (WopiOptions.Value.RunMacro.ToLower() == TRUE)
                    {
                        runMacroX(newFileName, fieldMap, processedByWWA, true, userRole);
                    }*/
        /*
                    if (WopiOptions.Value.RunMacro.ToLower() == TRUE)
                    {
                        var callRunMacro = "runMacro";

                        if (!String.IsNullOrEmpty(WopiOptions.Value.RunMacroVersion))
                        {
                            callRunMacro += WopiOptions.Value.RunMacroVersion;
                        }

                        MethodInfo runMacroMethod = this.GetType().GetMethod(callRunMacro);
                        object result = null;

                        try
                        {
                            runMacroMethod.Invoke(this, new object[] { newFileName, fieldMap, processedByWWA, true, userRole });
                        }
                        catch (Exception exc)
                        {
                            Console.Out.WriteLine("******** Error running macro method " + callRunMacro);
                        }

                    }

                    //document.SaveToFile(newFileName, FileFormat.Docm);

                    /*
                    document = new Document();
                    //document.LoadFromFile(sourceFile.FullName, FileFormat.WordXml);
                    document.LoadFromFile(newFileName);
                    document.SaveToFile(newFileName, FileFormat.OOXML);
                    */

        //var newFileName_macro_removed = newFileName.Replace(Path.GetFileNameWithoutExtension(newFileName), Path.GetFileNameWithoutExtension(newFileName)+"_macro_removed");

        /*
        var newFileName_auto = newFileName.Replace(Path.GetFileNameWithoutExtension(newFileName), Path.GetFileNameWithoutExtension(newFileName) + "_auto");
        var newFileName_doc = newFileName.Replace(Path.GetFileNameWithoutExtension(newFileName), Path.GetFileNameWithoutExtension(newFileName) + "_doc");
        var newFileName_docm = newFileName.Replace(Path.GetFileNameWithoutExtension(newFileName), Path.GetFileNameWithoutExtension(newFileName) + "_docm");
        var newFileName_docm2010 = newFileName.Replace(Path.GetFileNameWithoutExtension(newFileName), Path.GetFileNameWithoutExtension(newFileName) + "_docm2010");
        var newFileName_docm2013 = newFileName.Replace(Path.GetFileNameWithoutExtension(newFileName), Path.GetFileNameWithoutExtension(newFileName) + "_docm2013");
        var newFileName_docx = newFileName.Replace(Path.GetFileNameWithoutExtension(newFileName), Path.GetFileNameWithoutExtension(newFileName) + "_docx");
        var newFileName_docx2010 = newFileName.Replace(Path.GetFileNameWithoutExtension(newFileName), Path.GetFileNameWithoutExtension(newFileName) + "_docx2010");
        var newFileName_docx2013 = newFileName.Replace(Path.GetFileNameWithoutExtension(newFileName), Path.GetFileNameWithoutExtension(newFileName) + "_docx2013");
        var newFileName_odt = newFileName.Replace(Path.GetFileNameWithoutExtension(newFileName), Path.GetFileNameWithoutExtension(newFileName) + "_odt");
        var newFileName_ooxml = newFileName.Replace(Path.GetFileNameWithoutExtension(newFileName), Path.GetFileNameWithoutExtension(newFileName) + "_ooxml");
        var newFileName_rtf = newFileName.Replace(Path.GetFileNameWithoutExtension(newFileName), Path.GetFileNameWithoutExtension(newFileName) + "_rtf");
        var newFileName_wordml = newFileName.Replace(Path.GetFileNameWithoutExtension(newFileName), Path.GetFileNameWithoutExtension(newFileName) + "_wordml");
        var newFileName_wordxml = newFileName.Replace(Path.GetFileNameWithoutExtension(newFileName), Path.GetFileNameWithoutExtension(newFileName) + "_wordxml");
        var newFileName_html = newFileName.Replace(Path.GetFileNameWithoutExtension(newFileName), Path.GetFileNameWithoutExtension(newFileName) + "_html");
        var newFileName_pdf = newFileName.Replace(Path.GetFileNameWithoutExtension(newFileName), Path.GetFileNameWithoutExtension(newFileName) + "_pdf");



        //var newFileName3 = newFileName.Replace(Path.GetFileNameWithoutExtension(newFileName), Path.GetFileNameWithoutExtension(newFileName)+ "_");
        //newFileName2 = newFileName2.Replace("docx", "doc");
        document.SaveToFile(newFileName, FileFormat.OOXML);

        /*
        document.SaveToFile(newFileName_auto, FileFormat.Auto);
        document.SaveToFile(newFileName_doc.Replace(".docx",".doc"), FileFormat.Doc);
        document.SaveToFile(newFileName_docm.Replace(".docx", ".docm"), FileFormat.Docm);
        document.SaveToFile(newFileName_docm2010.Replace(".docx", ".docm"), FileFormat.Docm2010);
        document.SaveToFile(newFileName_docm2013.Replace(".docx",".docm"), FileFormat.Docm);
        document.SaveToFile(newFileName_docx, FileFormat.Docx);
        document.SaveToFile(newFileName_docx2010, FileFormat.Docx2010);
        document.SaveToFile(newFileName_docx2013, FileFormat.Docx2013);
        document.SaveToFile(newFileName_odt.Replace(".docx", ".odt"), FileFormat.Odt);
        document.SaveToFile(newFileName_ooxml.Replace(".docx", ".doc"), FileFormat.OOXML);
        document.SaveToFile(newFileName_rtf.Replace(".docx", ".rtf"), FileFormat.Rtf);
        document.SaveToFile(newFileName_wordml.Replace(".docx", ".doc"), FileFormat.WordML);
        document.SaveToFile(newFileName_wordxml, FileFormat.WordXml);
        document.SaveToFile(newFileName_html.Replace(".docx",".htm"), FileFormat.Html);
        document.SaveToFile(newFileName_pdf.Replace(".docx", ".pdf"), FileFormat.PDF);
        */

        //document.SaveToFile(newFileName_macro_removed, FileFormat.Docx2013);
        /*var newFileName_macro_removed = newFileName.Replace(Path.GetFileNameWithoutExtension(newFileName), Path.GetFileNameWithoutExtension(newFileName) + "_macro_removed");
        document.SaveToFile(newFileName_macro_removed, FileFormat.Docx2013);
        document.Close();*/



        /*using (OpenXMLTemplates.Documents.TemplateDocument doc = new OpenXMLTemplates.Documents.TemplateDocument(newFileName2))
        {

            try
            {
                //System.IO.File.Copy(newFileName, newFileName2);
                //using (OpenXMLTemplates.Documents.TemplateDocument doc = new OpenXMLTemplates.Documents.TemplateDocument(newFileName2)

                //string data = System.IO.File.ReadAllText("macrodata.json");
                //string data = System.IO.File.ReadAllText("macrodata.json");

                string data = "{ \"WorkerName\": \"MYWORKER\" }";
                var src = new VariableSource(data);

                var engine = new DefaultOpenXmlTemplateEngine();

                //Call the ReplaceAll method on the engine using the document and the variable source

                engine.ReplaceAll(doc, src);


                //src.LoadDataFromJson(data);

                //var replacer = new VariableControlReplacer();

                //replacer.ReplaceAll(doc, src);


                doc.SaveAs(newFileName3);
                //doc.Close();
                //doc.Dispose();

            }
            catch (Exception x)
            {
                Console.Out.WriteLine("Errors in substitution");
            }
            finally
            {
                doc.Close();
                doc.Dispose();
            }
        }*/
        /*



        try
        {
            var current = DateTime.Now;
            System.IO.File.SetCreationTime(newFileName, current);
            System.IO.File.SetLastWriteTime(newFileName, current);
            System.IO.File.SetLastAccessTime(newFileName, current);

            Console.WriteLine(System.IO.File.GetCreationTime(newFileName));
        }
        catch (Exception ex)
        {
            throw ex;
        }

    }
    else
    {
        //the file exists
        if (!useTemp)
        {
            System.IO.File.Delete(newFileName);
            //var document = word.Documents.Open(sourceFile.FullName);

            //var project = document.VBProject;
            //var module = project.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
            // module.CodeModule.AddFromString("CMSUpdateFields");
            //word.Run("CMSUpdateFields");
            //word.ActiveDocument.ActiveWindow.View.ShowFieldCodes = false;
            //word.ActiveDocument.ActiveWindow.View.ShowFieldCodes = false;
            //word.ActiveDocument.StoryRanges.Fields.Update();
            //word.ActiveDocument.StoryRanges.Fields.Update();


            //document.SaveAs2(newFileName, WdSaveFormat.wdFormatXMLDocument,
            //             CompatibilityMode: WdCompatibilityMode.wdWord2010);

            //word.ActiveDocument.Close();
            //word.Quit();

            //ComponentInfo.SetLicense("FREE-LIMITED-KEY");

            // In order to convert Word to PDF, we just need to:
            // 1. Load DOC or DOCX file into DocumentModel object.
            // 2. Save DocumentModel object to PDF file.
            //DocumentModel document = DocumentModel.Load(sourceFile.FullName);
            //document.Save(newFileName);
            Document document = new Document();
            document.LoadFromFile(sourceFile.FullName);

            CustomDocumentProperties custom = document.CustomDocumentProperties;

            Console.Out.WriteLine(sourceFile.FullName + ">>> " + (custom.Count).ToString());
            for (int i = 0; i < custom.Count; i++)
            {
                Console.Out.WriteLine(">>>Field " + i + ">>>" + custom[i].ToString());
            }

            foreach (KeyValuePair<string, DocumentProperty> entry in custom.CustomHash)
            {
                DocumentProperty fields = entry.Value;
                Console.Out.WriteLine(entry.Key + ">>>" + fields.Name + "<<<:>>>" + fields.Value + "<<<");
                fieldMap.Add(fields.Name, fields.Value.ToString());
            }





            //document.SaveToFile(newFileName, FileFormat.Docm2010);
            //document.ClearMacros();
            document.SaveToFile(newFileName, FileFormat.Docm);


            /* if (WopiOptions.Value.RunMacro.ToLower() == TRUE)
            {
                runMacroX(newFileName, fieldMap, processedByWWA, true, userRole);
            } */
        /*
            if (WopiOptions.Value.RunMacro.ToLower() == TRUE)
            {
                var callRunMacro = "runMacro";

                if (!String.IsNullOrEmpty(WopiOptions.Value.RunMacroVersion))
                {
                    callRunMacro += WopiOptions.Value.RunMacroVersion;
                }

                MethodInfo runMacroMethod = this.GetType().GetMethod(callRunMacro);
                object result = null;

                try
                {
                    runMacroMethod.Invoke(this, new object[] { newFileName, fieldMap, processedByWWA, true, userRole });
                }
                catch (Exception exc)
                {
                    Console.Out.WriteLine("******** Error running macro method " + callRunMacro);
                }

            }


            //runMacro(newFileName, fieldMap);


            var newFileName2 = newFileName.Replace(Path.GetFileNameWithoutExtension(newFileName), Path.GetFileNameWithoutExtension(newFileName) + "_2");
            var newFileName3 = newFileName.Replace(Path.GetFileNameWithoutExtension(newFileName), Path.GetFileNameWithoutExtension(newFileName) + "_3");
            //newFileName2 = newFileName2.Replace("docx", "doc");
            document.SaveToFile(newFileName2, FileFormat.Docx2013);
            document.Close();





            try
            {
                var current = DateTime.Now;
                System.IO.File.SetCreationTime(newFileName, current);
                System.IO.File.SetLastWriteTime(newFileName, current);
                System.IO.File.SetLastAccessTime(newFileName, current);

                Console.WriteLine(System.IO.File.GetCreationTime(newFileName));
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }
        else
        {
            //leave it alone
        }

    }

    var tempfile = path.Replace(WopiOptions.Value.Word2010Ext, "_retrieved") + WopiOptions.Value.Word2010Ext;
    if (System.IO.File.Exists(tempfile)) System.IO.File.Delete(tempfile);
    System.IO.File.Move(path, tempfile);

}
else
{
    //TODO
    ;
}
}*/

        private static void SetFormFieldValue(DocumentFormat.OpenXml.Wordprocessing.TextInput textInput, string value)
        {

            if (value == null) // Reset formfield using default if set.
            {
                if (textInput.DefaultTextBoxFormFieldString != null && textInput.DefaultTextBoxFormFieldString.Val.HasValue)
                    value = textInput.DefaultTextBoxFormFieldString.Val.Value;
            }

            // Enforce max length.
            short maxLength = 0; // Unlimited
            if (textInput.MaxLength != null && textInput.MaxLength.Val.HasValue)
                maxLength = textInput.MaxLength.Val.Value;
            if (value != null && maxLength > 0 && value.Length > maxLength)
                value = value.Substring(0, maxLength);

            // Not enforcing TextBoxFormFieldType (read documentation...).
            // Just note that the Word instance may modify the value of a formfield when user leave it based on TextBoxFormFieldType and Format.
            // A curious example:
            // Type Number, format "# ##0,00".
            // Set value to "2016 was the warmest year ever, at least since 1999.".
            // Open the document and select the field then tab out of it.
            // Value now is "2 016 tht,tt" (the logic behind this escapes me).

            // Format value. (Only able to handle formfields with textboxformfieldtype regular.)
            if (textInput.TextBoxFormFieldType != null
            && textInput.TextBoxFormFieldType.Val.HasValue
            && textInput.TextBoxFormFieldType.Val.Value != TextBoxFormFieldValues.Regular)
                throw new ApplicationException("SetFormField: Unsupported textboxformfieldtype, only regular is handled.\r\n" + textInput.Parent.OuterXml);
            if (!string.IsNullOrWhiteSpace(value)
            && textInput.Format != null
            && textInput.Format.Val.HasValue)
            {
                switch (textInput.Format.Val.Value)
                {
                    case "Uppercase":
                        value = value.ToUpperInvariant();
                        break;
                    case "Lowercase":
                        value = value.ToLowerInvariant();
                        break;
                    case "First capital":
                        value = value[0].ToString().ToUpperInvariant() + value.Substring(1);
                        break;
                    case "Title case":
                        value = System.Globalization.CultureInfo.InvariantCulture.TextInfo.ToTitleCase(value);
                        break;
                    default: // ignoring any other values (not supposed to be any)
                        break;
                }
            }

            // Find run containing "separate" fieldchar.
            Run rTextInput = textInput.Ancestors<Run>().FirstOrDefault();
            if (rTextInput == null) throw new ApplicationException("SetFormField: Did not find run containing textinput.\r\n" + textInput.Parent.OuterXml);
            Run rSeparate = rTextInput.ElementsAfter().FirstOrDefault(ru =>
               ru.GetType() == typeof(Run)
               && ru.Elements<FieldChar>().FirstOrDefault(fc =>
                  fc.FieldCharType == FieldCharValues.Separate)
                  != null) as Run;
            if (rSeparate == null) throw new ApplicationException("SetFormField: Did not find run containing separate.\r\n" + textInput.Parent.OuterXml);

            // Find run containg "end" fieldchar.
            Run rEnd = rTextInput.ElementsAfter().FirstOrDefault(ru =>
               ru.GetType() == typeof(Run)
               && ru.Elements<FieldChar>().FirstOrDefault(fc =>
                  fc.FieldCharType == FieldCharValues.End)
                  != null) as Run;
            if (rEnd == null) // Formfield value contains paragraph(s)
            {
                Paragraph p = rSeparate.Parent as Paragraph;
                Paragraph pEnd = p.ElementsAfter().FirstOrDefault(pa =>
                pa.GetType() == typeof(Paragraph)
                && pa.Elements<Run>().FirstOrDefault(ru =>
                   ru.Elements<FieldChar>().FirstOrDefault(fc =>
                      fc.FieldCharType == FieldCharValues.End)
                      != null)
                   != null) as Paragraph;
                if (pEnd == null) throw new ApplicationException("SetFormField: Did not find paragraph containing end.\r\n" + textInput.Parent.OuterXml);
                rEnd = pEnd.Elements<Run>().FirstOrDefault(ru =>
                   ru.Elements<FieldChar>().FirstOrDefault(fc =>
                      fc.FieldCharType == FieldCharValues.End)
                      != null);
            }

            // Remove any existing value.

            Run rFirst = rSeparate.NextSibling<Run>();
            if (rFirst == null || rFirst == rEnd)
            {
                RunProperties rPr = rTextInput.GetFirstChild<RunProperties>();
                if (rPr != null) rPr = rPr.CloneNode(true) as RunProperties;
                rFirst = rSeparate.InsertAfterSelf<Run>(new Run(new[] { rPr }));
            }
            rFirst.RemoveAllChildren<Text>();

            Run r = rFirst.NextSibling<Run>();
            while (r != rEnd)
            {
                if (r != null)
                {
                    r.Remove();
                    r = rFirst.NextSibling<Run>();
                }
                else // next paragraph
                {
                    Paragraph p = rFirst.Parent.NextSibling<Paragraph>();
                    if (p == null) throw new ApplicationException("SetFormField: Did not find next paragraph prior to or containing end.\r\n" + textInput.Parent.OuterXml);
                    r = p.GetFirstChild<Run>();
                    if (r == null)
                    {
                        // No runs left in paragraph, move other content to end of paragraph containing "separate" fieldchar.
                        p.Remove();
                        while (p.FirstChild != null)
                        {
                            OpenXmlElement oxe = p.FirstChild;
                            oxe.Remove();
                            if (oxe.GetType() == typeof(ParagraphProperties)) continue;
                            rSeparate.Parent.AppendChild(oxe);
                        }
                    }
                }
            }
            if (rEnd.Parent != rSeparate.Parent)
            {
                // Merge paragraph containing "end" fieldchar with paragraph containing "separate" fieldchar.
                Paragraph p = rEnd.Parent as Paragraph;
                p.Remove();
                while (p.FirstChild != null)
                {
                    OpenXmlElement oxe = p.FirstChild;
                    oxe.Remove();
                    if (oxe.GetType() == typeof(ParagraphProperties)) continue;
                    rSeparate.Parent.AppendChild(oxe);
                }
            }

            // Set new value.

            if (value != null)
            {
                // Word API use \v internally for newline and \r for para. We treat \v, \r\n, and \n as newline (Break).
                string[] lines = value.Replace("\r\n", "\n").Split(new char[] { '\v', '\n', '\r' });
                string line = lines[0];
                Text text = rFirst.AppendChild<Text>(new Text(line));
                if (line.StartsWith(" ") || line.EndsWith(" ")) text.SetAttribute(new OpenXmlAttribute("xml:space", null, "preserve"));
                for (int i = 1; i < lines.Length; i++)
                {
                    rFirst.AppendChild<Break>(new Break());
                    line = lines[i];
                    text = rFirst.AppendChild<Text>(new Text(lines[i]));
                    if (line.StartsWith(" ") || line.EndsWith(" ")) text.SetAttribute(new OpenXmlAttribute("xml:space", null, "preserve"));
                }
            }
            else
            { // An empty formfield of type textinput got char 8194 times 5 or maxlength if maxlength is in the range 1 to 4.
                short length = maxLength;
                if (length == 0 || length > 5) length = 5;
                rFirst.AppendChild(new Text(((char)8194).ToString()));
                r = rFirst;
                for (int i = 1; i < length; i++) r = r.InsertAfterSelf<Run>(r.CloneNode(true) as Run);
            }
        }

        protected void normalizeMarkup(WordprocessingDocument document)
        {
            SimplifyMarkupSettings settings = new SimplifyMarkupSettings
            {
                RemoveComments = false,
                RemoveContentControls = false,
                RemoveEndAndFootNotes = false,
                //RemoveFieldCodes = false,
                RemoveFieldCodes = false,
                RemoveLastRenderedPageBreak = false,
                RemovePermissions = false,
                RemoveProof = false,
                RemoveRsidInfo = false,
                RemoveSmartTags = false,
                RemoveSoftHyphens = false,
                ReplaceTabsWithSpaces = false,
            };
            MarkupSimplifier.SimplifyMarkup(document, settings);
        }

        protected void normalizeFieldCodesRuns(WordprocessingDocument document)
        {
            var fieldMasks = new string[] {
                DocFieldCodes.CATEGORY0,
                DocFieldCodes.CC,
                DocFieldCodes.CLAIM_NUMBER,
                DocFieldCodes.CREATE_DATE,
                DocFieldCodes.CUSTOMER_CARE_NUMBER,
                DocFieldCodes.DATE_OF_INJURY,
                DocFieldCodes.DESCRIPTION0,
                DocFieldCodes.EMPLOYER_ACCOUNT_NAME,
                DocFieldCodes.EMPLOYER_ACCOUNT_NUMBER,
                DocFieldCodes.EMPLOYER_CU,
                DocFieldCodes.HIDDEN,
                DocFieldCodes.INJURY_DESCRIPTION,
                DocFieldCodes.LETTER_DATE,
                DocFieldCodes.NAME_AND_ADDRESS_LN1,
                DocFieldCodes.NAME_AND_ADDRESS_LN2,
                DocFieldCodes.NAME_AND_ADDRESS_LN3,
                DocFieldCodes.NAME_AND_ADDRESS_LN4,
                DocFieldCodes.NAME_AND_ADDRESS_LN5,
                DocFieldCodes.NAME_AND_ADDRESS_LN6,
                DocFieldCodes.NAME_AND_ADDRESS_LN7,
                DocFieldCodes.PHONE_MEMO_CONTACT_NUMBER,
                DocFieldCodes.PHONE_MEMO_CONTACT_ORG,
                DocFieldCodes.PHONE_MEMO_CONTACT_PERSON,
                DocFieldCodes.PHONE_MEMO_CONTACT_TYPE,
                DocFieldCodes.PRIMARY_RECIPIENT_NAME,
                DocFieldCodes.REGARDING,
                DocFieldCodes.USERNAME,
                DocFieldCodes.USER_DEPARTMENT,
                DocFieldCodes.USER_INITIALS
            };
            normalizeFieldCodesInElement(document.MainDocumentPart.RootElement, fieldMasks);

            foreach (var headerPart in document.MainDocumentPart.HeaderParts)
            {
                normalizeFieldCodesInElement(headerPart.Header, fieldMasks);
            }

            foreach (var footerPart in document.MainDocumentPart.FooterParts)
            {
                normalizeFieldCodesInElement(footerPart.Footer, fieldMasks);
            }

        }

        protected void normalizeFieldCodesInElement(OpenXmlElement element, string[] regexpMasks)
        {
            foreach (var run in element.Descendants<Run>()
                .Select(item => (Run)item)
                .ToList())
            {
                var fieldChar = run.Descendants<FieldChar>().FirstOrDefault();
                if (fieldChar != null && fieldChar.FieldCharType == FieldCharValues.Begin)
                {
                    string fieldContent = "";
                    List<Run> runsInFieldCode = new List<Run>();

                    var currentRun = run.NextSibling();
                    while ((currentRun is Run) && currentRun.Descendants<FieldCode>().FirstOrDefault() != null)
                    {
                        var currentRunFieldCode = currentRun.Descendants<FieldCode>().FirstOrDefault();
                        fieldContent += currentRunFieldCode.InnerText;
                        runsInFieldCode.Add((Run)currentRun);
                        currentRun = currentRun.NextSibling();
                    }

                    // If there is more than one Run for the FieldCode, and is one we must change, set the complete text in the first Run and remove the rest
                    if (runsInFieldCode.Count > 1)
                    {
                        // Check fielcode to know it's one that we must simplify (for not to change TOC, PAGEREF, etc.)
                        bool applyTransform = false;
                        foreach (string regexpMask in regexpMasks)
                        {
                            Regex regex = new Regex(regexpMask);
                            Match match = regex.Match(fieldContent);
                            if (match.Success)
                            {
                                applyTransform = true;
                                Console.Out.WriteLine("In normalizeFieldCodesInElement(), fieldContent=" + fieldContent);
                                break;
                            }
                        }

                        if (applyTransform)
                        {
                            var currentRunFieldCode = runsInFieldCode[0].Descendants<FieldCode>().FirstOrDefault();
                            currentRunFieldCode.Text = fieldContent;
                            runsInFieldCode.RemoveAt(0);

                            foreach (Run runToRemove in runsInFieldCode)
                            {
                                runToRemove.Remove();
                            }
                        }
                    }
                }
            }
        }

        /*
                public void startupWord(Microsoft.Office.Interop.Word.Application wordPointer)
                { 
                        wordPointer = word;
                        word.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone;
                        word.Visible = false;
                        word.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;
                }
        */


        /*        public void convertAndRunMacro(string path, bool runMacro)
                {
                    bool wordExit = true;
                    try
                    {
                        //var word = new Microsoft.Office.Interop.Word.Application();
                        wordPointer = word;
                        word.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone;
                        word.Visible = false;
                        word.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;



                        if (path.ToLower().EndsWith(WopiOptions.Value.Word2010Ext))
                        {

                            var sourceFile = new FileInfo(path);
                            string newFileName = sourceFile.FullName.Replace(WopiOptions.Value.Word2010Ext, WopiOptions.Value.WordExt);

                            if (!System.IO.File.Exists(newFileName))
                            {
                                System.IO.File.Delete(newFileName);
                                var document = word.Documents.Open(sourceFile.FullName);

                                //var project = document.VBProject;
                                //var module = project.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
                                // module.CodeModule.AddFromString("CMSUpdateFields");
                                try
                                {
                                    word.Run("CMSUpdateFields");
                                }
                                catch (Exception macroEx)
                                {
                                    //throw macroEx;
                                }
                                word.ActiveDocument.ActiveWindow.View.ShowFieldCodes = false;
                                //word.ActiveDocument.ActiveWindow.View.ShowFieldCodes = false;
                                //word.ActiveDocument.StoryRanges.Fields.Update();
                                //word.ActiveDocument.StoryRanges.Fields.Update();


                                document.SaveAs2(newFileName, WdSaveFormat.wdFormatXMLDocument,
                                             CompatibilityMode: WdCompatibilityMode.wdWord2010);
                                word.ActiveDocument.Close();

                                try
                                {
                                var current = DateTime.Now;
                                System.IO.File.SetCreationTime(newFileName, current);
                                System.IO.File.SetLastWriteTime(newFileName, current);
                                System.IO.File.SetLastAccessTime(newFileName, current);
                                
                                    Console.WriteLine(System.IO.File.GetCreationTime(newFileName));
                                }
                                catch (Exception ex)
                                {
                                    throw ex;
                                }

                            }
                            else
                            {
                                //the file exists
                                if (true)
                                {
                                    System.IO.File.Delete(newFileName);
                                    var document = word.Documents.Open(sourceFile.FullName);

                                    //var project = document.VBProject;
                                    //var module = project.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
                                    // module.CodeModule.AddFromString("CMSUpdateFields");
                                    try
                                    {
                                        word.Run("CMSUpdateFields");
                                    }
                                    catch (Exception macroEx)
                                    {
                                        //;
                                    }
                                    word.ActiveDocument.ActiveWindow.View.ShowFieldCodes = false;
                                    //word.ActiveDocument.ActiveWindow.View.ShowFieldCodes = false;
                                    //word.ActiveDocument.StoryRanges.Fields.Update();
                                    //word.ActiveDocument.StoryRanges.Fields.Update();


                                    document.SaveAs2(newFileName, WdSaveFormat.wdFormatXMLDocument,
                                                 CompatibilityMode: WdCompatibilityMode.wdWord2010);

                                    word.ActiveDocument.Close();
                                    word.Quit();
                                    wordExit = false;

                                    try
                                    {
                                       var current = DateTime.Now;
                                System.IO.File.SetCreationTime(newFileName, current);
                                System.IO.File.SetLastWriteTime(newFileName, current);
                                System.IO.File.SetLastAccessTime(newFileName, current);
                                
                                        Console.WriteLine(System.IO.File.GetCreationTime(newFileName));
                                    }
                                    catch (Exception ex)
                                    {
                                        throw ex;
                                    }

                                }
                                else
                                {
                                    //leave it alone
                                }

                            }


                            System.IO.File.Delete(path);

                        }
                        else
                        {
                            //TODO
                            ;
                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                    finally
                    {
                        if (!wordExit && wordPointer != null)
                        {
                            //if (wordPointer.get_IsObjectValid(wordPointer.ActiveDocument) && wordPointer.ActiveDocument != null) wordPointer.ActiveDocument.Close();
                            wordPointer.Quit();
                        }

                    }
                }

        */


        public void convertDOCMtoDOCX(string path, Boolean useTemp = false)
        {
            bool fileChanged = false;

            var sourceFile = new FileInfo(path);
            var fileName = sourceFile.FullName;
            var newFileName = sourceFile.FullName.Replace(WopiOptions.Value.Word2010Ext, WopiOptions.Value.WordExt);

            using (WordprocessingDocument document =
                WordprocessingDocument.Open(fileName, true))
            {
                // Access the main document part.
                // var docPart = document.MainDocumentPart;

                // Look for the vbaProject part. If it is there, delete it.
                //var vbaPart = docPart.VbaProjectPart;
                //if (vbaPart != null)
                //{
                // Delete the vbaProject part and then save the document.
                //docPart.DeletePart(vbaPart);
                //docPart.Document.Save();

                // Change the document type to
                // not macro-enabled.
                document.ChangeDocumentType(
                    WordprocessingDocumentType.Document);

                // Track that the document has been changed.
                fileChanged = true;
                //}
            }

            // If anything goes wrong in this file handling,
            // the code will raise an exception back to the caller.
            if (fileChanged)
            {
                // Create the new .docx filename.
                //var newFileName = Path.ChangeExtension(fileName, ".docx");

                // If it already exists, it will be deleted!
                if (System.IO.File.Exists(newFileName))
                {
                    System.IO.File.Delete(newFileName);
                }

                // Rename the file.
                System.IO.File.Copy(fileName, newFileName);
            }
        }




        public void convertDocToDocx(string path, Boolean useTemp = false)
        {
            //Microsoft.Office.Interop.Word.Application wordPointer = null;
            bool wordExit = true;
            try
            {
                var word = new Microsoft.Office.Interop.Word.Application();
                wordPointer = word;
                word.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone;
                word.Visible = false;
                word.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;
                object TRUE_VALUE = true;
                object FALSE_VALUE = false;
                object MISSING_VALUE = System.Reflection.Missing.Value;



                if (path.ToLower().EndsWith(WopiOptions.Value.Word2010Ext))
                {

                    Microsoft.Office.Interop.Word.Document document = null;
                    var sourceFile = new FileInfo(path);
                    string newFileName = sourceFile.FullName.Replace(WopiOptions.Value.Word2010Ext, WopiOptions.Value.WordExt);

                    if (!System.IO.File.Exists(newFileName))
                    {
                        System.IO.File.Delete(newFileName);

                        if (System.IO.File.Exists(sourceFile.FullName))
                        {
                            System.IO.File.Copy(sourceFile.FullName, sourceFile.FullName + ".temp");
                            try
                            {
                                var theDoc = sourceFile.FullName;
                                //document = word.Documents.Open(theDoc, ref FALSE_VALUE, ref TRUE_VALUE, ref FALSE_VALUE, ref MISSING_VALUE, ref MISSING_VALUE, ref MISSING_VALUE, ref MISSING_VALUE, ref MISSING_VALUE, ref MISSING_VALUE, ref MISSING_VALUE, ref FALSE_VALUE, ref TRUE_VALUE, ref MISSING_VALUE, ref TRUE_VALUE, ref MISSING_VALUE);
                                new System.Threading.Tasks.Task(() => document = word.Documents.Open(theDoc)).Start();

                            }
                            catch (Exception x)
                            {
                                throw new Exception(sourceFile.FullName + " cannot be loaded: " + x.StackTrace);
                            }
                        }
                        else
                        {
                            throw new Exception("tempFile" + sourceFile.FullName + "cannot be created");
                        }


                        //var project = document.VBProject;
                        //var module = project.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
                        // module.CodeModule.AddFromString("CMSUpdateFields");
                        if (document != null)
                        {
                            ;
                        }
                        else
                        {
                            throw new Exception(sourceFile.FullName + " cannot be loaded");
                        }
                        try
                        {
                            word.Run("CMSUpdateFields");
                        }
                        catch (Exception macroEx)
                        {
                            //throw macroEx;
                        }
                        //word.ActiveDocument.ActiveWindow.View.ShowFieldCodes = false;
                        //word.ActiveDocument.ActiveWindow.View.ShowFieldCodes = false;
                        //word.ActiveDocument.StoryRanges.Fields.Update();
                        //word.ActiveDocument.StoryRanges.Fields.Update();


                        document.SaveAs2(newFileName, WdSaveFormat.wdFormatXMLDocument,
                                     CompatibilityMode: WdCompatibilityMode.wdWord2010);
                        //document.SaveAs(newFileName, WdSaveFormat.wdFormatXMLDocument);
                        word.ActiveDocument.Close();
                        word.Quit();
                        try
                        {
                            var current = DateTime.Now;
                            System.IO.File.SetCreationTime(newFileName, current);
                            System.IO.File.SetLastWriteTime(newFileName, current);
                            System.IO.File.SetLastAccessTime(newFileName, current);

                            Console.WriteLine(System.IO.File.GetCreationTime(newFileName));
                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }

                    }
                    else
                    {
                        //the file exists
                        if (!useTemp)
                        {
                            System.IO.File.Delete(newFileName);
                            //document = word.Documents.Open(sourceFile.FullName);
                            if (System.IO.File.Exists(sourceFile.FullName))
                            {
                                System.IO.File.Copy(sourceFile.FullName, sourceFile.FullName + ".temp");
                                try
                                {
                                    document = word.Documents.Open(sourceFile.FullName + ".temp");
                                }
                                catch (Exception x)
                                {
                                    throw new Exception(sourceFile.FullName + " cannot be loaded: " + x.StackTrace);
                                }
                            }
                            else
                            {
                                throw new Exception("tempFile" + sourceFile.FullName + "cannot be created");
                            }
                            //var project = document.VBProject;
                            //var module = project.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
                            // module.CodeModule.AddFromString("CMSUpdateFields");
                            try
                            {
                                word.Run("CMSUpdateFields");
                            }
                            catch (Exception macroEx)
                            {
                                //;
                            }
                            word.ActiveDocument.ActiveWindow.View.ShowFieldCodes = false;
                            //word.ActiveDocument.ActiveWindow.View.ShowFieldCodes = false;
                            //word.ActiveDocument.StoryRanges.Fields.Update();
                            //word.ActiveDocument.StoryRanges.Fields.Update();


                            document.SaveAs2(newFileName, WdSaveFormat.wdFormatXMLDocument,
                                         CompatibilityMode: WdCompatibilityMode.wdWord2010);

                            //document.SaveAs(newFileName, WdSaveFormat.wdFormatXMLDocument);

                            word.ActiveDocument.Close();
                            word.Quit();
                            wordExit = false;

                            try
                            {
                                var current = DateTime.Now;
                                System.IO.File.SetCreationTime(newFileName, current);
                                System.IO.File.SetLastWriteTime(newFileName, current);
                                System.IO.File.SetLastAccessTime(newFileName, current);

                                Console.WriteLine(System.IO.File.GetCreationTime(newFileName));
                            }
                            catch (Exception ex)
                            {
                                throw ex;
                            }

                        }
                        else
                        {
                            //leave it alone
                        }

                    }


                    System.IO.File.Delete(path);

                }
                else
                {
                    //TODO
                    ;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (!wordExit && wordPointer != null)
                {
                    //if (wordPointer.get_IsObjectValid(wordPointer.ActiveDocument) && wordPointer.ActiveDocument != null) wordPointer.ActiveDocument.Close();
                    // wordPointer.Quit();
                    ;
                }

            }
        }


        public void DOCXconvertDocToDocx(string path, Boolean useTemp = false)
        {
            //Microsoft.Office.Interop.Word.Application wordPointer = null;
            bool wordExit = true;
            try
            {
                if (path.ToLower().EndsWith(WopiOptions.Value.Word2010Ext))
                {

                    Microsoft.Office.Interop.Word.Document document = null;
                    var sourceFile = new FileInfo(path);
                    string newFileName = sourceFile.FullName.Replace(WopiOptions.Value.Word2010Ext, WopiOptions.Value.WordExt);

                    if (!System.IO.File.Exists(newFileName))
                    {
                        System.IO.File.Delete(newFileName);

                        if (System.IO.File.Exists(sourceFile.FullName))
                        {
                            System.IO.File.Copy(sourceFile.FullName, sourceFile.FullName + ".temp");
                            try
                            {
                                document = word.Documents.Open(sourceFile.FullName + ".temp");
                            }
                            catch (Exception x)
                            {
                                throw new Exception(sourceFile.FullName + " cannot be loaded: " + x.StackTrace);
                            }
                        }
                        else
                        {
                            throw new Exception("tempFile" + sourceFile.FullName + "cannot be created");
                        }


                        //var project = document.VBProject;
                        //var module = project.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
                        // module.CodeModule.AddFromString("CMSUpdateFields");
                        if (document != null)
                        {
                            ;
                        }
                        else
                        {
                            throw new Exception(sourceFile.FullName + " cannot be loaded");
                        }
                        try
                        {
                            word.Run("CMSUpdateFields");
                        }
                        catch (Exception macroEx)
                        {
                            //throw macroEx;
                        }
                        //word.ActiveDocument.ActiveWindow.View.ShowFieldCodes = false;
                        //word.ActiveDocument.ActiveWindow.View.ShowFieldCodes = false;
                        //word.ActiveDocument.StoryRanges.Fields.Update();
                        //word.ActiveDocument.StoryRanges.Fields.Update();


                        document.SaveAs2(newFileName, WdSaveFormat.wdFormatXMLDocument,
                                     CompatibilityMode: WdCompatibilityMode.wdWord2010);
                        //document.SaveAs(newFileName, WdSaveFormat.wdFormatXMLDocument);
                        word.ActiveDocument.Close();
                        word.Quit();
                        try
                        {
                            var current = DateTime.Now;
                            System.IO.File.SetCreationTime(newFileName, current);
                            System.IO.File.SetLastWriteTime(newFileName, current);
                            System.IO.File.SetLastAccessTime(newFileName, current);

                            Console.WriteLine(System.IO.File.GetCreationTime(newFileName));
                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }

                    }
                    else
                    {
                        //the file exists
                        if (!useTemp)
                        {
                            System.IO.File.Delete(newFileName);
                            //document = word.Documents.Open(sourceFile.FullName);
                            if (System.IO.File.Exists(sourceFile.FullName))
                            {
                                System.IO.File.Copy(sourceFile.FullName, sourceFile.FullName + ".temp");
                                try
                                {
                                    document = word.Documents.Open(sourceFile.FullName + ".temp");
                                }
                                catch (Exception x)
                                {
                                    throw new Exception(sourceFile.FullName + " cannot be loaded: " + x.StackTrace);
                                }
                            }
                            else
                            {
                                throw new Exception("tempFile" + sourceFile.FullName + "cannot be created");
                            }
                            //var project = document.VBProject;
                            //var module = project.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
                            // module.CodeModule.AddFromString("CMSUpdateFields");
                            try
                            {
                                word.Run("CMSUpdateFields");
                            }
                            catch (Exception macroEx)
                            {
                                //;
                            }
                            word.ActiveDocument.ActiveWindow.View.ShowFieldCodes = false;
                            //word.ActiveDocument.ActiveWindow.View.ShowFieldCodes = false;
                            //word.ActiveDocument.StoryRanges.Fields.Update();
                            //word.ActiveDocument.StoryRanges.Fields.Update();


                            document.SaveAs2(newFileName, WdSaveFormat.wdFormatXMLDocument,
                                         CompatibilityMode: WdCompatibilityMode.wdWord2010);

                            //document.SaveAs(newFileName, WdSaveFormat.wdFormatXMLDocument);

                            word.ActiveDocument.Close();
                            word.Quit();
                            wordExit = false;

                            try
                            {
                                var current = DateTime.Now;
                                System.IO.File.SetCreationTime(newFileName, current);
                                System.IO.File.SetLastWriteTime(newFileName, current);
                                System.IO.File.SetLastAccessTime(newFileName, current);
                                Console.WriteLine(System.IO.File.GetCreationTime(newFileName));
                            }
                            catch (Exception ex)
                            {
                                throw ex;
                            }

                        }
                        else
                        {
                            //leave it alone
                        }

                    }


                    System.IO.File.Delete(path);

                }
                else
                {
                    //TODO
                    ;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (!wordExit && wordPointer != null)
                {
                    //if (wordPointer.get_IsObjectValid(wordPointer.ActiveDocument) && wordPointer.ActiveDocument != null) wordPointer.ActiveDocument.Close();
                    // wordPointer.Quit();
                    ;
                }

            }
        }


        public void convertDocxToDoc(string path)
        {
            var word = new Microsoft.Office.Interop.Word.Application();

            if (path.ToLower().EndsWith(WopiOptions.Value.WordExt))
            {
                var sourceFile = new FileInfo(path);
                Microsoft.Office.Interop.Word.Document document;
                new System.Threading.Tasks.Task(() => document = word.Documents.Open(sourceFile.FullName)).Start();
                //word.Documents.Open(sourceFile.FullName);

                string newFileName = sourceFile.FullName.Replace(WopiOptions.Value.WordExt, WopiOptions.Value.Word2010Ext);
                // document.SaveAs2(newFileName, WdSaveFormat.wdFormatXMLDocument,
                //                 CompatibilityMode: WdCompatibilityMode.wdWord2010);

                word.ActiveDocument.Close();
                word.Quit();

                //System.IO.File.Delete(path);
            }
        }

        public void convertDocxToDocSpire(string path)
        {
            try
            {
                if (path.ToLower().EndsWith(WopiOptions.Value.WordExt))
                {
                    var sourceFile = new FileInfo(path);
                    string newFileName = sourceFile.FullName.Replace(WopiOptions.Value.WordExt, WopiOptions.Value.Word2010Ext);
                    string macroSource = sourceFile.FullName.Replace(WopiOptions.Value.WordExt, "_saved.docm");
                    string xmlSource = sourceFile.FullName.Replace(WopiOptions.Value.WordExt, "_saved.xml");

                    //ComponentInfo.SetLicense("FREE-LIMITED-KEY");

                    // In order to convert Word to PDF, we just need to:
                    // 1. Load DOC or DOCX file into DocumentModel object.
                    // 2. Save DocumentModel object to PDF file.
                    //DocumentModel document = DocumentModel.Load(sourceFile.FullName);
                    //document.Save(newFileName);

                    if (System.IO.File.Exists(newFileName))
                    {
                        System.IO.File.Delete(newFileName);
                    }

                    Document document = new Document();
                    //document.LoadFromFile(sourceFile.FullName, FileFormat.Docx2013);
                    document.LoadFromFile(sourceFile.FullName);
                    document.SaveToFile(newFileName, FileFormat.WordML);
                    document.SaveToFile(xmlSource, FileFormat.WordXml);
                    System.IO.File.Move(sourceFile.FullName, macroSource);
                }
            }
            catch (Exception ex)
            {
                throw new WordWebException(ex.StackTrace);
            }
        }



        public void convertDocxToDocAspose(string path)
        {
            try
            {

                if (path.ToLower().EndsWith(WopiOptions.Value.WordExt))
                {
                    var sourceFile = new FileInfo(path);
                    //string sourceFileName = path.Replace(WopiOptions.Value.WordExt, "_saved" + WopiOptions.Value.WordExt);
                    string sourceFileName = sourceFile.FullName;

                    runMacroY(sourceFileName);

                    string newFileName = sourceFile.FullName.Replace(WopiOptions.Value.WordExt, WopiOptions.Value.Word2010Ext);
                    string newFileNameWithMacro = sourceFile.FullName.Replace(WopiOptions.Value.WordExt, WopiOptions.Value.WordMacroExt);

                    OoxmlSaveOptions ooxmlSaveOptions = new OoxmlSaveOptions
                    {
                        //Compliance = OoxmlCompliance.Iso29500_2008_Strict,
                        SaveFormat = SaveFormat.FlatOpc
                        //SaveFormat = SaveFormat.WordML
                    };


                    //SaveOptions saveOptions = SaveOptions.CreateSaveOptions(SaveFormat.WordML);
                    WordML2003SaveOptions saveOptions = new WordML2003SaveOptions();
                    saveOptions.MemoryOptimization = true;
                    //saveOptions.PrettyFormat = true;




                    //ComponentInfo.SetLicense("FREE-LIMITED-KEY");

                    // In order to convert Word to PDF, we just need to:
                    // 1. Load DOC or DOCX file into DocumentModel object.
                    // 2. Save DocumentModel object to PDF file.
                    //DocumentModel document = DocumentModel.Load(sourceFile.FullName);
                    //document.Save(newFileName);
                    //System.IO.File.Copy(path, path.Replace(WopiOptions.Value.WordExt, "_saved"+ WopiOptions.Value.WordExt));

                    if (System.IO.File.Exists(newFileName))
                    {
                        System.IO.File.Delete(newFileName);
                    }

                    if (System.IO.File.Exists(newFileNameWithMacro))
                    {
                        System.IO.File.Delete(newFileNameWithMacro);
                    }
                    //Document document = new Document();
                    //document.LoadFromFile(sourceFile.FullName, FileFormat.Docx2013);
                    //document.LoadFromFile(sourceFile.FullName);
                    //System.IO.File.Copy(sourceFileName, newFileNameWithMacro);
                    //System.IO.File.Copy(sourceFileName, newFileName);


                    var disableRemoteResourcesOptions = new Aspose.Words.Loading.LoadOptions
                    {
                        ResourceLoadingCallback = new DisableRemoteResourcesHandler()
                    };

                    Aspose.Words.Document document = null;
                    if (!(null == WopiOptions.Value.ConversionEngineDisableExternalResources || WopiOptions.Value.ConversionEngineDisableExternalResources.Contains(FALSE)))
                        document = new Aspose.Words.Document(sourceFileName, disableRemoteResourcesOptions);
                    else
                        document = new Aspose.Words.Document(sourceFileName);

                    document.Save(newFileNameWithMacro, SaveFormat.Docm);
                    var copyResult = copyMacro(Path.GetFileNameWithoutExtension(newFileName) + WopiOptions.Value.WithMacroSuffix + WopiOptions.Value.WordMacroExt,
                              newFileNameWithMacro);

                    Aspose.Words.Loading.LoadOptions lopt = new Aspose.Words.Loading.LoadOptions
                    {
                        LoadFormat = LoadFormat.Docm
                    };

                    Aspose.Words.Document document2 = null;
                    document2 = new Aspose.Words.Document(newFileNameWithMacro, lopt);
                    //new Aspose.Words.Document(newFileNameWithMacro);

                    //document.Save(newFileName, SaveFormat.WordML);

                    document2.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2003);
                    var docProtKey = Path.GetFileNameWithoutExtension(newFileName);

                    var newXml = newFileName.Replace(WopiOptions.Value.Word2010Ext, ".xml");
                    var newWordDoc = newFileName.Replace(WopiOptions.Value.Word2010Ext, WopiOptions.Value.PreXmlParseSuffix + WopiOptions.Value.Word2010Ext);

                    //if (null != docProtKey && !(docProtection.ContainsKey(docProtKey))) document.Unprotect();
                    document2.Save(newFileName, saveOptions);

                    var xmlString = System.IO.File.ReadAllText(newFileName, Encoding.UTF8);
                    XDocument xmlDoc = XDocument.Parse(xmlString);

                    //XDocument xmlDoc = XDocument.Load(newFileName, System.Xml.Linq.LoadOptions.PreserveWhitespace);
                    //xmlDoc.Declaration = new XDeclaration("1.0", "utf-8", null);


                    /*var attributes = from r in xmlDoc.Descendants("documentProtection")
                                   select new
                                   {
                                       edit = r.Attribute("edit").Value,
                                       enforcement = r.Attribute("enforcement").Value,
                                       pwd = r.Attribute("unprotectPassword").Value,
                                   };*/
                    XElement docProtNode = null;
                    if (null != xmlDoc && null != xmlDoc.Descendants()) docProtNode = xmlDoc.Descendants().SingleOrDefault(p => p.Name.LocalName == "documentProtection");
                    /*foreach (var r in attributes)
                    {
                        //Console.WriteLine("AUTHOR = " + r.Author + Environment.NewLine + "DESCRIPTION = " + r.Description);
                        if (!String.IsNullOrEmpty(r.pwd))
                        {
                            XAttribute _pwd = xmlDoc.Element("docPr").Element("w:documentProtection").Attribute("unprotectPassword");
                            _pwd.Remove();
                        }
                        if (String.IsNullOrEmpty(r.enforcement))
                        {
                            XElement _docProt = xmlDoc.Element("w:docPr").Element("w:documentProtection");
                            XAttribute _enforcement = new XAttribute("enforcement", "off");
                            _docProt.Add(_enforcement);                            
                        }
                    }*/
                    XAttribute uPwd = null;
                    if (null != docProtNode) uPwd = docProtNode?.Attributes().SingleOrDefault(e => e.Name.LocalName == "unprotectPassword");
                    if (null != uPwd) uPwd.Remove();
                    XAttribute enforcement = null;
                    if (null != docProtNode) enforcement = docProtNode?.Attributes().SingleOrDefault(e => e.Name.LocalName == "enforcement");
                    XAttribute _enforcement = null;
                    if (null != docProtNode) _enforcement = new XAttribute(docProtNode.GetNamespaceOfPrefix("w") + "enforcement", "0");
                    if (null != docProtNode && null == enforcement && null != _enforcement) docProtNode.Add(_enforcement);
                    //StringWriter writer = new Utf8StringWriter();
                    //if (null != xmlDoc) xmlDoc.Save(newXml, System.Xml.Linq.SaveOptions.None);

                    //MemoryStream ms = new MemoryStream();
                    System.Xml.XmlWriterSettings settings = new System.Xml.XmlWriterSettings();
                    settings.Encoding = new UTF8Encoding(false);
                    settings.ConformanceLevel = ConformanceLevel.Document;
                    settings.Indent = false;
                    settings.NewLineHandling = NewLineHandling.None;
                    //System.Xml.Linq.SaveOptions.DisableFormatting
                    //settings.Indent = true;
                    //ToString(SaveOptions.DisableFormatting)
                    if (null != xmlDoc)
                    {
                        using (System.Xml.XmlWriter xw = System.Xml.XmlTextWriter.Create(newXml, settings))
                        {
                            xmlDoc.Save(xw);
                            xw.Flush();
                        }
                    }

                    if (System.IO.File.Exists(newFileName)) System.IO.File.Move(newFileName, newWordDoc);
                    if (System.IO.File.Exists(newXml)) System.IO.File.Move(newXml, newFileName);

                }
            }
            catch (Exception ex)
            {
                throw new WordWebException(ex.StackTrace);
            }
        }


        /*public void replaceFieldCodes(string file)
        {
            using (WordprocessingDocument document = WordprocessingDocument.Open(file, true))
            {
                MainDocumentPart main = document.MainDocumentPart;

                foreach (FooterPart foot in main.FooterParts)
                {
                    foreach (var fld in foot.RootElement.Descendants<FieldCode>())
                    {
                        if (fld != null && fld.InnerText.Contains("REF NG_MACRO"))
                        {
                            Run rFldCode = (Run)fld.Parent;

                            // Get the three (3) other Runs that make up our merge field
                            Run rBegin = rFldCode.PreviousSibling<Run>();
                            Run rSep = rFldCode.NextSibling<Run>();
                            Run rText = rSep.NextSibling<Run>();
                            Run rEnd = rText.NextSibling<Run>();

                            // Get the Run that holds the Text element for our merge field
                            // Get the Text element and replace the text content 
                            Text t = rText.GetFirstChild<Text>();
                            //t.Text = replacementText;

                            // Remove all the four (4) Runs for our merge field
                            rFldCode.Remove();
                            rBegin.Remove();
                            rSep.Remove();
                            rEnd.Remove();
                        }
                    }

                    foot.Footer.Save();
                }
                document.MainDocumentPart.Document.Save();
                document.Close();
            }
        }*/

        /*public void convertDocxToDocGlue(string path)
        {
            try
            {
                if (path.ToLower().EndsWith(WopiOptions.Value.WordExt))
                {
                    var sourceFile = new FileInfo(path);
                    string newFileName = sourceFile.FullName.Replace(WopiOptions.Value.WordExt, WopiOptions.Value.Word2010Ext);

                    //ComponentInfo.SetLicense("FREE-LIMITED-KEY");

                    // In order to convert Word to PDF, we just need to:
                    // 1. Load DOC or DOCX file into DocumentModel object.
                    // 2. Save DocumentModel object to PDF file.
                    //DocumentModel document = DocumentModel.Load(sourceFile.FullName);
                    //document.Save(newFileName);

                    if (System.IO.File.Exists(newFileName))
                    {
                        System.IO.File.Delete(newFileName);
                    }

                    Document document = new Document();
                    //document.LoadFromFile(sourceFile.FullName, FileFormat.Docx2013);
                    document.LoadFromFile(sourceFile.FullName);
                    document.SaveToFile(newFileName, FileFormat.WordXml);

                    using (Doc doc = new Doc(sourceFile.FullName))
                        doc.SaveAs(newFileName);

                    //System.IO.File.Delete(path);
                }
            }
            catch (Exception ex)
            {
                throw new WordWebException(ex.StackTrace);
            }
        }*/


        /*
        KEY	                              VALUE	  DESCRIPTION
        wdFormatDocument	                  0	  Microsoft Word format.
        wdFormatDocument97	                  0	  Microsoft Word 97 document format.
        wdFormatDocumentDefault	             16	  Word default document file format. For Microsoft Office Word 2007, this is the DOCX format.
        wdFormatDOSText	                      4	  Microsoft DOS text format.
        wdFormatDOSTextLineBreaks	          5	  Microsoft DOS text with line breaks preserved.
        wdFormatEncodedText	                  7	  Encoded text format.
        wdFormatFilteredHTML	             10	  Filtered HTML format.
        wdFormatFlatXML	                     19	  Reserved for internal use.
        wdFormatFlatXMLMacroEnabled	         20	  Reserved for internal use.
        wdFormatFlatXMLTemplate	             21	  Reserved for internal use.
        wdFormatFlatXMLTemplateMacroEnabled	 22	  Reserved for internal use.
        wdFormatHTML	                      8	  Standard HTML format.
        wdFormatOpenDocumentText	         23	
        wdFormatPDF	                         17	  PDF format.
        wdFormatRTF	                          6	  Rich text format (RTF).
        wdFormatStrictOpenXMLDocument	     24	  Strict Open XML document format.
        wdFormatTemplate	                  1	 Microsoft Word template format.
        wdFormatTemplate97	                  1	 Word 97 template format.
        wdFormatText	                      2	 Microsoft Windows text format.
        wdFormatTextLineBreaks	              3	 Microsoft Windows text format with line breaks preserved.
        wdFormatUnicodeText	                  7	 Unicode text format.
        wdFormatWebArchive	                  9	 Web archive format.
        wdFormatXML	                         11	 Extensible Markup Language (XML) format.
        wdFormatXMLDocument	                 12	 XML document format.
        wdFormatXMLDocumentMacroEnabled	     13	 XML template format with macros enabled.
        wdFormatXMLTemplate	                 14	 XML template format.
        wdFormatXMLTemplateMacroEnabled	15	XML template format with macros enabled.
        wdFormatXPS	18	XPS format.
        */

        public void convertType(string path, string outpath)
        {
            var word = new Microsoft.Office.Interop.Word.Application();

            if (!(path.ToLower().EndsWith(WopiOptions.Value.WordExt)))
            {
                var sourceFile = new FileInfo(path);
                var document = word.Documents.Open(sourceFile.FullName);

                string newFileName = sourceFile.FullName.Replace(path, outpath);
                //document.SaveAs2(newFileName, WdSaveFormat.wdFormatXMLDocument,
                //                 CompatibilityMode: WdCompatibilityMode.wdWord2010);
                document.SaveAs(newFileName, WdSaveFormat.wdFormatXMLDocument);

                word.ActiveDocument.Close();
                word.Quit();

                System.IO.File.Delete(path);
            }
        }

        /*
         
            {
              "errors" : [ {
                "code" : 500,
                "message" : "An un-handled server exception occurred. Please contact your administrator.",
                "level" : "error"
              } ]
            }

        */

        public async Task<ActionResult> handleErrorMessage(ErrorMessage err)
        {
            return handleErrorMessage(err.code, err.message, err.level).Result;
        }


        public string getMimeType(string fileName)
        {
            var provider = new FileExtensionContentTypeProvider();
            string contentType;
            var extension = System.IO.Path.GetExtension(fileName);
            if (extension == WopiOptions.Value.Word2010Ext)
            {
                return "application/msword";
            }
            if (extension == WopiOptions.Value.WordExt)
            {
                return "application/vnd.openxmlformats-officedocument.wordprocessing";
            }
            if (!provider.TryGetContentType(fileName, out contentType))
            {
                contentType = "application/octet-stream";
            }
            return contentType;
        }

        public async Task<ActionResult> handleErrorMessage(string errCode = "500", string errMessage = "An un-handled server exception occurred. Please contact your administrator.", string errLevel = "error")
        {
            ErrorMessage errorMsg = new ErrorMessage
            {
                code = "200",
                message = "Successful",
                level = "Error"
            };

            APIResponse response = new APIResponse
            {
                status = "0",
                timestamp = DateTime.Now,
                errors = new List<ErrorMessage> { errorMsg }
            };

            var apiResponse = new
            {
                status = "0",
                timestamp = DateTime.Now,
                errors = new List<ErrorMessage> { errorMsg }
            };

            //var payload = JsonConvert.SerializeObject(apiResponse);
            var serializer = new Newtonsoft.Json.JsonSerializer();
            var stringWriter = new StringWriter();
            using (var writer = new JsonTextWriter(stringWriter))
            {
                writer.QuoteName = true;
                serializer.Serialize(writer, errorMsg);
            }
            var json = stringWriter.ToString();

            /*var message = string.Format("Product with id = {0} not found", id);
            HttpError err = new HttpError(message);
            return Request.CreateResponse(HttpStatusCode.NotFound, err);*/

            HttpContext.Response.Headers["content-type"] = "application/json";
            switch (errCode)
            {
                case "200":
                case "201":
                    var p200 = new
                    {
                        status = "0",
                        timpstamp = String.Format("{0:G}", DateTime.Now)
                    };
                    return Ok(p200);
                    break;
                case "400":
                    var p400 = new
                    {
                        status = errCode,
                        timpstamp = String.Format("{0:G}", DateTime.Now),
                        errors = new List<Dictionary<string, string>> { new Dictionary<string, string> { { "code", errCode }, { "message", errMessage }, { "level", "error" } } }
                    };
                    return Ok(p400);
                    break;
                case "404":
                    var p404 = new
                    {
                        status = "404",
                        timpstamp = String.Format("{0:G}", DateTime.Now),
                        errors = new List<Dictionary<string, string>> { new Dictionary<string, string> { { "code", "404" }, { "message", "HTTP 404 Not Found" }, { "level", "error" } } }
                    };
                    return Ok(p404);
                    break;
                case "405":
                    var p405 = new
                    {
                        status = "404",
                        timpstamp = String.Format("{0:G}", DateTime.Now),
                        errors = new List<Dictionary<string, string>> { new Dictionary<string, string> { { "code", "405" }, { "message", "HTTP 405 Method Not Allowund" }, { "level", "error" } } }
                    };
                    return Ok(p405);
                    break;
                case "500":
                default:
                    var p500 = new
                    {
                        status = "500",
                        timpstamp = String.Format("{0:G}", DateTime.Now),
                        errors = new List<Dictionary<string, string>> { new Dictionary<string, string> { { "code", "500" }, { "message", "An un-handled server exception occurred. Please contact your administrator." }, { "level", "error" } } }
                    };
                    Response.StatusCode = (int)HttpStatusCode.InternalServerError;
                    //return new JsonResult(p500);
                    return Ok(p500);

            }




        }

        [System.Web.Http.Route("/error")]
        [Microsoft.AspNetCore.Mvc.Route("/error")]
        public IActionResult handleError(ErrorMessage error = null)
        {
            if (error != null && !(String.IsNullOrEmpty(error.code) && String.IsNullOrEmpty(error.message) && String.IsNullOrEmpty(error.level)))
            {
                return handleErrorMessage(error).Result;
            }
            else
            {
                var context = HttpContext.Features.Get<IExceptionHandlerFeature>();
                var code = Response.StatusCode.ToString();
                if (context != null)
                {
                    if (context.Error != null)
                    {
                        if (context.Error.InnerException != null)
                        {
                            if (context.Error.InnerException != null)
                            {
                                WordWebException wwerror = ((WordWebException)context.Error.InnerException);
                                if (wwerror.errMessage != null)
                                {
                                    if (wwerror.errMessage.area == ErrorArea.VIEW)
                                    {
                                        return ErrorView(wwerror).Result;
                                    }
                                    else
                                    {
                                        return handleErrorMessage(wwerror.errMessage).Result;
                                    }
                                }
                            }
                        }

                        if (context.Error.Source != null && context.Error.Message != null)
                        {
                            return handleErrorMessage(code, context.Error.Message, context.Error.Source).Result;
                        }
                    }
                    return handleErrorMessage().Result;
                }
                else
                {
                    return handleErrorMessage().Result;
                }
            }
        }


        protected string[] buildCORSOrigins(string aFile)
        {

            string EXE = Assembly.GetExecutingAssembly().GetName().Name;
            string iniPath = new FileInfo(aFile).FullName;


            ArrayList originsArray = new ArrayList();
            originsArray.Add("http://localhost");


            //lines[0] = "localhost";
            if (System.IO.File.Exists(iniPath))
            {
                StringBuilder originsString = new StringBuilder();

                var lines = System.IO.File.ReadAllLines(iniPath);
                for (var i = 0; i < lines.Length; i += 1)
                {
                    var line = lines[i];
                    line = line.Trim();
                    line = line.Replace("\\n", "");
                    line = line.Replace("\\r", "");

                    if (!string.IsNullOrEmpty(line))
                    {
                        originsArray.Add(line);
                        //originsString.Append(line);
                        /*if ((i < lines.Length - 1) && (lines.Length > 0))
                        {
                            originsString.Append(",");
                        }*/
                    }
                }
            }
            return (string[])originsArray.ToArray(typeof(string));
        }

        /*

        public Action doUpConversion(string engine, string doc)
        {
            if (conversionEngine.ContainsKey(engine))
            {
                return conversionEngine[engine](doc);
            }            
        }

        public Action doDownConversion(string engine, string doc)
        {
            if (conversionEngine.ContainsKey(engine))
            {
                return conversionEngine[engine](doc);
            }
        } */

        public enum CustomPropertyTypes : int
        {
            YesNo,
            Text,
            DateTime,
            NumberInteger,
            NumberDouble
        }


        protected string getCustomProperty(
                WordprocessingDocument document,
                string propertyName
                )
        {
            // Given a document name, a property name/value, and the property type, 
            // add a custom property to a document. The method returns the original
            // value, if it existed.

            string returnValue = null;

            if (null != document)
            {
                var customProps = document.CustomFilePropertiesPart;

                if (null == customProps) return null;

                var props = customProps.Properties;
                if (props != null)
                {
                    // This will trigger an exception if the property's Name 
                    // property is null, but if that happens, the property is damaged, 
                    // and probably should raise an exception.
                    var prop = props.Where(
                        p => ((DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty)p).Name.Value
                            == propertyName).FirstOrDefault();

                    // Does the property exist? If so, get the return value, 
                    // and then delete the property.
                    if (null != prop)
                    {
                        if (!String.IsNullOrWhiteSpace(prop.InnerText))
                            returnValue = prop.InnerText.Trim();
                    }
                }
            }
            return returnValue;
        }


        protected string setCustomProperty(
            WordprocessingDocument document,
            string propertyName,
            object propertyValue,
            CustomPropertyTypes propertyType)
        {
            // Given a document name, a property name/value, and the property type, 
            // add a custom property to a document. The method returns the original
            // value, if it existed.

            string returnValue = null;

            var newProp = new DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty();
            bool propSet = false;

            // Calculate the correct type.
            switch (propertyType)
            {
                case CustomPropertyTypes.DateTime:

                    // Be sure you were passed a real date, 
                    // and if so, format in the correct way. 
                    // The date/time value passed in should 
                    // represent a UTC date/time.
                    if ((propertyValue) is DateTime)
                    {
                        newProp.VTFileTime =
                            new VTFileTime(string.Format("{0:s}Z",
                                Convert.ToDateTime(propertyValue)));
                        propSet = true;
                    }

                    break;

                case CustomPropertyTypes.NumberInteger:
                    if ((propertyValue) is int)
                    {
                        newProp.VTInt32 = new VTInt32(propertyValue.ToString());
                        propSet = true;
                    }

                    break;

                case CustomPropertyTypes.NumberDouble:
                    if (propertyValue is double)
                    {
                        newProp.VTFloat = new VTFloat(propertyValue.ToString());
                        propSet = true;
                    }

                    break;

                case CustomPropertyTypes.Text:
                    newProp.VTLPWSTR = new VTLPWSTR(propertyValue.ToString());
                    propSet = true;

                    break;

                case CustomPropertyTypes.YesNo:
                    if (propertyValue is bool)
                    {
                        // Must be lowercase.
                        newProp.VTBool = new VTBool(
                          Convert.ToBoolean(propertyValue).ToString().ToLower());
                        propSet = true;
                    }
                    break;
            }

            if (!propSet)
            {
                // If the code was not able to convert the 
                // property to a valid value, throw an exception.
                throw new InvalidDataException("propertyValue");
            }

            // Now that you have handled the parameters, start
            // working on the document.
            newProp.FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";
            newProp.Name = propertyName;

            //using (var document = WordprocessingDocument.Open(fileName, true))
            //{
            var customProps = document.CustomFilePropertiesPart;
            if (customProps == null)
            {
                // No custom properties? Add the part, and the
                // collection of properties now.
                customProps = document.AddCustomFilePropertiesPart();
                customProps.Properties =
                    new DocumentFormat.OpenXml.CustomProperties.Properties();
            }

            var props = customProps.Properties;
            if (props != null)
            {
                // This will trigger an exception if the property's Name 
                // property is null, but if that happens, the property is damaged, 
                // and probably should raise an exception.
                var prop =
                    props.Where(
                    p => ((DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty)p).Name.Value
                        == propertyName).FirstOrDefault();

                // Does the property exist? If so, get the return value, 
                // and then delete the property.
                if (prop != null)
                {
                    returnValue = prop.InnerText;
                    prop.Remove();
                }

                // Append the new property, and 
                // fix up all the property ID values. 
                // The PropertyId value must start at 2.
                props.AppendChild(newProp);
                int pid = 2;
                foreach (DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty item in props)
                {
                    item.PropertyId = pid++;
                }
                props.Save();
                //}
            }
            return returnValue;
        }


        protected void changeCompatibilityModeOfDocumentPart(MainDocumentPart part)
        {
            DocumentSettingsPart settingsPart = part.DocumentSettingsPart;
            if (settingsPart == null)
                settingsPart = part.AddNewPart<DocumentSettingsPart>();
            settingsPart.Settings = new Settings(
              new DocumentFormat.OpenXml.Wordprocessing.Compatibility(
                new CompatibilitySetting()
                {
                    Name = new EnumValue<CompatSettingNameValues>
                           (CompatSettingNameValues.CompatibilityMode),
                    Val = new StringValue("15"),
                    Uri = new StringValue
                           ("http://schemas.microsoft.com/office/word")
                }
               )
             );
            settingsPart.Settings.Save();
        }


        private string getCMSServicePwd()
        {
            if (null == cmsServicePwd)
            {
                if (null != WopiOptions.Value.CMSServicePwd)
                {
                    cmsServicePwd = PasswordHelper.Decrypt(WopiOptions.Value.CMSServicePwd);
                    return cmsServicePwd;
                }
                else
                    return WopiOptions.Value.CMSServicePwd;
            }
            else
                return cmsServicePwd;
        }


        private string getTTL()
        {
            var tokenExpiry = Double.Parse(WopiOptions.Value.SessionTimeout);
            if (tokenExpiry <= (double)(0))
            {
                tokenExpiry = (double)(0);
                return new string("0");
            }
            var seconds = System.Math.Abs((DateTime.Now - pointInTime).TotalSeconds);
            tokenExpiry = (seconds + tokenExpiry) * 1000;
            var tokenExpiryStr = tokenExpiry.ToString();
            return tokenExpiryStr;
        }


        private string formatText(string inText)
        {
            if (null == inText || String.IsNullOrWhiteSpace(inText) || inText.Contains("BOOKMARK_UNDEFINED"))
                return new string(" ");
            else
                return inText;
        }

        private ArrayList getIDPList()
        {
            if (ignoreDocumentProtection.Count > 0) return ignoreDocumentProtection;
            string[] idp = WopiOptions.Value.IgnoreDocumentProtection;
            if (null == idp) return new ArrayList();
            else
            {
                if (idp.Length > 0)
                    ignoreDocumentProtection.AddRange(idp);
            }
            return ignoreDocumentProtection;
        }

        public string GenerateRandomUuid()
        {
            int _min = 1000;
            int _max = 9999;
            Random _rdm = new Random();
            return _rdm.Next(_min, _max).ToString();
        }

        public bool useServiceIdAsProxyUser()
        {
            if (null == WopiOptions.Value.CMSServiceIdAsProxyUser)
                return true;
            else
            {
                var useServiceId = true;
                if (Boolean.TryParse(WopiOptions.Value.CMSServiceIdAsProxyUser, out useServiceId))
                {
                    return useServiceId;
                }
                else
                {
                    return true;
                }
            }

        }

        public string getUserForCMSAPI(string userID)
        {
            if (useServiceIdAsProxyUser()) return WopiOptions.Value.CMSServiceId;
            else
            {
                return userID;
            }

        }


        public static void DiffDictionaries<T, U>(
            Dictionary<T, U> dicA,
            Dictionary<T, U> dicB,
            Dictionary<T, U> dicAdd,
            Dictionary<T, U> dicDel)
        {
            // dicDel has entries that are in A, but not in B, 
            // ie they were deleted when moving from A to B
            diffDicSub<T, U>(dicA, dicB, dicDel);

            // dicAdd has entries that are in B, but not in A,
            // ie they were added when moving from A to B
            diffDicSub<T, U>(dicB, dicA, dicAdd);
        }

        private static void diffDicSub<T, U>(
            Dictionary<T, U> dicA,
            Dictionary<T, U> dicB,
            Dictionary<T, U> dicAExceptB)
        {
            // Walk A, and if any of the entries are not
            // in B, add them to the result dictionary.
            foreach (KeyValuePair<T, U> kvp in dicA)
            {
                if (!dicB.Contains(kvp))
                {
                    dicAExceptB[kvp.Key] = kvp.Value;
                }
            }
        }



        private class Utf8StringWriter : StringWriter
        {
            public override Encoding Encoding { get { return Encoding.UTF8; } }
        }


        public bool copyMacro(string src, string dest)
        {
            if (!System.IO.File.Exists(src))
                return false;
            if (!System.IO.File.Exists(dest))
                return false;

            try
            {
                using (WordprocessingDocument destDoc = WordprocessingDocument.Open(dest, true))
                {
                    WordprocessingDocument srcDoc = WordprocessingDocument.Open(src, false);
                    MainDocumentPart destMain = destDoc.MainDocumentPart;
                    MainDocumentPart srcMain = srcDoc.MainDocumentPart;
                    // Create the document structure and add some text.
                    //mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();
                    //DocumentFormat.OpenXml.Wordprocessing.Body body = mainPart.Document.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Body());
                    //Paragraph para = body.AppendChild(new Paragraph());
                    //Run run = para.AppendChild(new Run());
                    //run.AppendChild(new Text("This is a macro enabled doc. Hit Ctrl+Insert Now."));
                    // Get VBA parts from source document
                    if (null != srcMain && null != destMain)
                    {
                        VbaProjectPart vbaSrc = srcMain?.VbaProjectPart;
                        VbaDataPart vbaDatSrc = null;
                        if (null != vbaSrc) vbaDatSrc = vbaSrc?.VbaDataPart;
                        CustomizationPart keymapSrc = srcMain?.CustomizationPart;
                        // Create VBA parts in destination document
                        VbaProjectPart vbaProjectPart1 = destMain?.AddNewPart<VbaProjectPart>("rId9");
                        VbaDataPart vbaDataPart1 = vbaProjectPart1?.AddNewPart<VbaDataPart>("rId1");
                        CustomizationPart customKeyMapPart = destMain?.AddNewPart<CustomizationPart>("rId10");
                        // Copy part contents
                        if (null != vbaSrc && null != vbaProjectPart1) vbaProjectPart1.FeedData(vbaSrc.GetStream());
                        if (null != vbaDatSrc && null != vbaDataPart1) vbaDataPart1.FeedData(vbaDatSrc.GetStream());
                        if (null != keymapSrc && null != customKeyMapPart) customKeyMapPart.FeedData(keymapSrc.GetStream());
                        destDoc.Save();
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.Out.WriteLine("ERROR: Cannot copy macro...");
                throw new WordWebException("WWA ERROR: Cannot copy macro from retreived document...", ex);
            }
            return false;
        }


        public string getHostUrl()
        {
            if (null != this.hostUrl) return this.hostUrl;
            string protocol = null;
            string hostname = null;
            string port = null;
            string serverName = null;
            string fqdn = null;

            Regex hostFilter = new Regex(@"\s*(?<protocol>)\://(?<hostname>):(?<port>)\/*");
            //"HostUrl": "http://ld449820.wcbbc.wcbmain.com:5000"
            var result = hostFilter.Matches(WopiOptions.Value.HostUrl);
            foreach (Match match in result)
            {
                protocol = match.Groups["protocol"].Value;
                hostname = match.Groups["hostname"].Value;
                port = match.Groups["port"].Value;
            }

            if (!String.IsNullOrWhiteSpace(WopiOptions.Value.Protocol)) protocol = WopiOptions.Value.Protocol;

            if (String.IsNullOrWhiteSpace(hostname)) hostname = "localhost";

            if (!String.IsNullOrWhiteSpace(hostname))
            {
                if (hostname.ToLower() == "localhost")
                {
                    serverName = System.Environment.MachineName; //host name sans domain
                    if (!String.IsNullOrEmpty(serverName)) fqdn = System.Net.Dns.GetHostEntry(serverName)?.HostName;
                    if (!String.IsNullOrEmpty(fqdn)) hostname = fqdn.ToLower();
                }
                else
                {
                    hostname = hostname.ToLower();
                }
            }
            if (String.IsNullOrWhiteSpace(port)) port = "5000";

            this.hostUrl = $"{protocol}://{hostname}:{port}";

            return this.hostUrl;

        }

    }
}