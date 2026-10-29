using System;
using System.Collections.Generic;
namespace PowerPointIntelligence.Library
{
    internal sealed class LibraryRegistration
    {
        internal string Id { get; set; }
        internal string Root { get; set; }
        internal string Name { get; set; }
        internal bool Recursive { get; set; }
        internal bool Enabled { get; set; }
        internal DateTime? LastRefreshUtc { get; set; }
    }
    internal sealed class LibraryIndexProgress
    {
        internal string LibraryId { get; set; }
        internal string Source { get; set; }
        internal int FilesSeen { get; set; }
        internal int FilesUpdated { get; set; }
        internal int SlidesIndexed { get; set; }
    }
    internal sealed class LibraryRefreshResult
    {
        internal string JobId { get; set; }
        internal int FilesSeen { get; set; }
        internal int FilesUpdated { get; set; }
        internal int SlidesIndexed { get; set; }
        internal int MissingFiles { get; set; }
    }
    internal sealed class LibrarySlideReference
    {
        internal string SlideId { get; set; }
        internal string PresentationId { get; set; }
        internal string LibraryId { get; set; }
        internal string SourcePath { get; set; }
        internal string SourceFingerprint { get; set; }
        internal string SlideFingerprint { get; set; }
        internal string NativeId { get; set; }
        internal string Title { get; set; }
        internal string Text { get; set; }
        internal string Category { get; set; }
        internal string Author { get; set; }
        internal string Layout { get; set; }
        internal string Master { get; set; }
        internal int Position { get; set; }
        internal double Width { get; set; }
        internal double Height { get; set; }
        internal double Rank { get; set; }
        internal DateTime ModifiedUtc { get; set; }
    }
    internal enum LibrarySearchMode { AllWords,Phrase,Prefix,ExactText }
    internal sealed class LibrarySearchRequest
    {
        internal string Query { get; set; } = "";
        internal string LibraryId { get; set; }
        internal string PresentationId { get; set; }
        internal string Category { get; set; }
        internal string Tag { get; set; }
        internal string Author { get; set; }
        internal string Layout { get; set; }
        internal LibrarySearchMode Mode { get; set; } = LibrarySearchMode.AllWords;
        internal int Offset { get; set; }
        internal int Limit { get; set; } = 100;
    }
    internal sealed class LibrarySearchPage
    {
        internal long Total { get; set; }
        internal IReadOnlyList<LibrarySlideReference> Slides { get; set; }
    }
}