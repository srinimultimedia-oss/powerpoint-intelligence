using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
namespace PowerPointIntelligence.Library
{
    internal sealed class LibraryRefreshNotice
    {
        internal string LibraryId {get;set;}
        internal string ErrorCode {get;set;}
        internal LibraryRefreshResult Result {get;set;}
    }
    internal sealed class LibraryRefreshMonitor : IDisposable
    {
        readonly PiidxCatalog catalog;
        readonly int seconds,watchLimit;
        readonly Action<LibraryRefreshNotice> notify;
        readonly ConcurrentDictionary<string,byte> dirty=new ConcurrentDictionary<string,byte>();
        readonly List<FileSystemWatcher> watchers=new List<FileSystemWatcher>();
        readonly CancellationTokenSource cancellation=new CancellationTokenSource();
        readonly object gate=new object();
        readonly Timer timer;
        DateTime lastPoll=DateTime.UtcNow;
        long lastChange;
        int busy,disposed;
        Task current=Task.FromResult(0);
        internal Task Current {get{lock(gate)return current;}}
        internal bool IsBusy {get{return Volatile.Read(ref busy)!=0;}}
        internal LibraryRefreshMonitor(PiidxCatalog catalog,int refreshSeconds,int watchRootLimit,Action<LibraryRefreshNotice> notify)
        {
            if(catalog==null)throw new ArgumentNullException("catalog");if(refreshSeconds<5||refreshSeconds>86400||watchRootLimit<0||watchRootLimit>128)throw new ArgumentOutOfRangeException("refreshSeconds");
            this.catalog=catalog;seconds=refreshSeconds;watchLimit=watchRootLimit;this.notify=notify;timer=new Timer(_=>Tick(),null,500,500);Reset();
        }
        internal void Reset()
        {
            if(Volatile.Read(ref disposed)!=0)return;
            lock(gate)
            {
                foreach(var watcher in watchers)watcher.Dispose();watchers.Clear();
                foreach(var library in catalog.Libraries().Where(l=>l.Enabled).Take(watchLimit))
                {
                    string id=library.Id;
                    try
                    {
                        if(!Directory.Exists(library.Root))continue;
                        var watcher=new FileSystemWatcher(library.Root){IncludeSubdirectories=library.Recursive,NotifyFilter=NotifyFilters.FileName|NotifyFilters.DirectoryName|NotifyFilters.LastWrite|NotifyFilters.Size};
                        FileSystemEventHandler changed=(s,e)=>{string ext=Path.GetExtension(e.FullPath);if(String.IsNullOrEmpty(ext)||ext.Equals(".pptx",StringComparison.OrdinalIgnoreCase)||ext.Equals(".potx",StringComparison.OrdinalIgnoreCase))Mark(id);};
                        watcher.Changed+=changed;watcher.Created+=changed;watcher.Deleted+=changed;watcher.Renamed+=(s,e)=>Mark(id);watcher.Error+=(s,e)=>{Mark(id);Notice(new LibraryRefreshNotice{LibraryId=id,ErrorCode="watcher-unavailable-polling-retained"});};watcher.EnableRaisingEvents=true;watchers.Add(watcher);
                    }
                    catch(Exception ex)when(ex is IOException||ex is UnauthorizedAccessException||ex is ArgumentException){Notice(new LibraryRefreshNotice{LibraryId=id,ErrorCode="watcher-unavailable-polling-retained"});}
                }
            }
        }
        void Notice(LibraryRefreshNotice value){if(Volatile.Read(ref disposed)==0&&notify!=null)try{notify(value);}catch(ObjectDisposedException){}catch(InvalidOperationException){}}
        void Mark(string id){if(Volatile.Read(ref disposed)==0){dirty[id]=0;Interlocked.Exchange(ref lastChange,DateTime.UtcNow.Ticks);}}
        internal void Request(string id){Mark(id);}
        internal void RequestAll(){foreach(var library in catalog.Libraries().Where(l=>l.Enabled))Mark(library.Id);}
        void Tick()
        {
            if(Volatile.Read(ref disposed)!=0)return;
            try
            {
                if((DateTime.UtcNow-lastPoll).TotalSeconds>=seconds){lastPoll=DateTime.UtcNow;RequestAll();}
                if(dirty.IsEmpty||(DateTime.UtcNow.Ticks-Interlocked.Read(ref lastChange))<TimeSpan.FromMilliseconds(750).Ticks||Interlocked.CompareExchange(ref busy,1,0)!=0)return;
                lock(gate)
                {
                    if(Volatile.Read(ref disposed)!=0){Interlocked.Exchange(ref busy,0);return;}
                    current=Task.Run(()=>
                {
                    try
                    {
                        var known=new HashSet<string>(catalog.Libraries().Where(l=>l.Enabled).Select(l=>l.Id));
                        foreach(string id in dirty.Keys.ToArray())
                        {
                            byte ignored;if(!dirty.TryRemove(id,out ignored)||!known.Contains(id))continue;cancellation.Token.ThrowIfCancellationRequested();
                            try{var result=catalog.Refresh(id,cancellation.Token);Notice(new LibraryRefreshNotice{LibraryId=id,Result=result});}
                            catch(OperationCanceledException){throw;}
                            catch(Exception ex)when(ex is IOException||ex is UnauthorizedAccessException||ex is System.Xml.XmlException||ex is ArgumentException){Notice(new LibraryRefreshNotice{LibraryId=id,ErrorCode=ex.GetType().Name});}
                        }
                    }
                    catch(OperationCanceledException){}
                    catch(Exception ex)when(ex is IOException||ex is UnauthorizedAccessException){Notice(new LibraryRefreshNotice{ErrorCode=ex.GetType().Name});}
                    finally{Interlocked.Exchange(ref busy,0);}
                });
                }
            }
            catch(Exception ex)when(ex is IOException||ex is UnauthorizedAccessException||ex is ObjectDisposedException){Notice(new LibraryRefreshNotice{ErrorCode=ex.GetType().Name});}
        }
        public void Dispose()
        {
            if(Interlocked.Exchange(ref disposed,1)!=0)return;timer.Dispose();cancellation.Cancel();
            Task completion;lock(gate){foreach(var watcher in watchers)watcher.Dispose();watchers.Clear();completion=current;}
            completion.ContinueWith(_=>cancellation.Dispose(),TaskScheduler.Default);
        }
    }
}