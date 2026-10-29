# PIIDX file format specification — schema 1 foundation

Status: database foundation implemented and tested; complete index writer/search/gallery migration remains in progress. No shipped PIIDX compatibility promise is made before integration acceptance.

PIIDX is a SQLite database, not a renamed proprietary file. The `.piidx` extension identifies this application's format. SQLite application_id is `0x50494958`; user_version is 1. INDEX_METADATA.format_id must equal `PowerPointIntelligence.PIIDX`, and schema_version must agree with user_version. Engine version is currently 0.1.0. A foreign, future-version or nonempty unversioned database is rejected without adoption or downgrade.

## Exact schema

`src/Library.Core/PiidxSchema.cs` contains the authoritative DDL below. Initialization is the version-0-to-1 migration and runs under BEGIN IMMEDIATE/COMMIT, rolling back all schema changes on failure. Future migrations must be explicit, transactional and tested; unknown versions are never rewritten. Validation compares required ordered columns AND exact authored table/index/trigger definitions, header identifiers, SQLite integrity_check and foreign_key_check.

```sql
CREATE TABLE INDEX_METADATA(singleton INTEGER PRIMARY KEY CHECK(singleton=1),format_id TEXT NOT NULL,schema_version INTEGER NOT NULL,engine_version TEXT NOT NULL,created_utc TEXT NOT NULL,last_successful_update TEXT,library_id TEXT,source_fingerprint TEXT,integrity_digest TEXT);
CREATE TABLE LIBRARIES(library_id TEXT PRIMARY KEY NOT NULL,root_path TEXT NOT NULL UNIQUE,name TEXT NOT NULL,recursive INTEGER NOT NULL CHECK(recursive IN(0,1)),registered_utc TEXT NOT NULL,last_refresh_utc TEXT,enabled INTEGER NOT NULL DEFAULT 1 CHECK(enabled IN(0,1)));
CREATE TABLE PRESENTATIONS(presentation_id TEXT PRIMARY KEY NOT NULL,library_id TEXT NOT NULL REFERENCES LIBRARIES(library_id) ON DELETE CASCADE,source_path TEXT NOT NULL,relative_path TEXT NOT NULL,file_type TEXT NOT NULL CHECK(file_type IN('.pptx','.potx')),file_size INTEGER NOT NULL CHECK(file_size>=0),modified_utc TEXT NOT NULL,content_hash TEXT NOT NULL,index_status TEXT NOT NULL,slide_count INTEGER NOT NULL CHECK(slide_count>=0),UNIQUE(library_id,relative_path));
CREATE TABLE SLIDES(slide_id TEXT PRIMARY KEY NOT NULL,presentation_id TEXT NOT NULL REFERENCES PRESENTATIONS(presentation_id) ON DELETE CASCADE,position INTEGER NOT NULL CHECK(position>0),native_slide_id TEXT NOT NULL,part_uri TEXT NOT NULL,title TEXT NOT NULL,text TEXT NOT NULL,keywords TEXT NOT NULL,category TEXT NOT NULL,author TEXT NOT NULL,layout TEXT NOT NULL,master TEXT NOT NULL,width REAL NOT NULL CHECK(width>0),height REAL NOT NULL CHECK(height>0),content_fingerprint TEXT NOT NULL,UNIQUE(presentation_id,native_slide_id),UNIQUE(presentation_id,position));
CREATE TABLE SLIDE_OBJECTS(slide_id TEXT NOT NULL REFERENCES SLIDES(slide_id) ON DELETE CASCADE,object_id TEXT NOT NULL,kind TEXT NOT NULL,name TEXT NOT NULL,text TEXT NOT NULL,bounds_json TEXT NOT NULL,metadata_json TEXT NOT NULL,PRIMARY KEY(slide_id,object_id));
CREATE TABLE SLIDE_TAGS(slide_id TEXT NOT NULL REFERENCES SLIDES(slide_id) ON DELETE CASCADE,tag TEXT NOT NULL,normalized_tag TEXT NOT NULL,PRIMARY KEY(slide_id,normalized_tag));
CREATE TABLE SLIDE_PREVIEWS(slide_id TEXT NOT NULL REFERENCES SLIDES(slide_id) ON DELETE CASCADE,width INTEGER NOT NULL CHECK(width>0 AND width<=4096),height INTEGER NOT NULL CHECK(height>0 AND height<=4096),mime TEXT NOT NULL CHECK(mime='image/png'),image BLOB NOT NULL,content_hash TEXT NOT NULL,source_fingerprint TEXT NOT NULL,rendered_utc TEXT NOT NULL,PRIMARY KEY(slide_id,width,height));
CREATE VIRTUAL TABLE SLIDE_SEARCH USING fts5(slide_id UNINDEXED,title,text,keywords,tokenize='unicode61 remove_diacritics 2');
CREATE TABLE INDEX_JOBS(job_id TEXT PRIMARY KEY NOT NULL,library_id TEXT REFERENCES LIBRARIES(library_id) ON DELETE SET NULL,started_utc TEXT NOT NULL,completed_utc TEXT,state TEXT NOT NULL CHECK(state IN('running','completed','cancelled','failed','interrupted')),files_seen INTEGER NOT NULL DEFAULT 0,files_updated INTEGER NOT NULL DEFAULT 0,slides_indexed INTEGER NOT NULL DEFAULT 0);
CREATE TABLE INDEX_ERRORS(error_id INTEGER PRIMARY KEY,job_id TEXT REFERENCES INDEX_JOBS(job_id) ON DELETE CASCADE,source_path TEXT,stage TEXT NOT NULL,error_code TEXT NOT NULL,message TEXT NOT NULL,recorded_utc TEXT NOT NULL);
CREATE INDEX IX_PRESENTATIONS_LIBRARY ON PRESENTATIONS(library_id,index_status);
CREATE INDEX IX_SLIDES_CATEGORY ON SLIDES(category);
CREATE INDEX IX_SLIDE_TAGS_NORMALIZED ON SLIDE_TAGS(normalized_tag,slide_id);
CREATE TRIGGER SLIDES_DELETE_SEARCH AFTER DELETE ON SLIDES BEGIN DELETE FROM SLIDE_SEARCH WHERE slide_id=OLD.slide_id; END;
```

## Entity meaning and compatibility

- INDEX_METADATA records format/schema/engine identity, creation and success timestamps, library identity, source fingerprint and logical integrity digest. Writer-populated integrity fields are reserved in this foundation; they are not yet computed or claimed as verified.
- LIBRARIES supports multiple registered roots, display names, recursion, enabled state and refresh timestamps.
- PRESENTATIONS records stable presentation identity, library ownership, original source/relative paths, type, size, UTC modification time, content hash, status and slide count. Original sources are never modified by database initialization.
- SLIDES stores stable identity separately from slide position/native ID, title/text/keywords/category, author/layout/master, dimensions and content fingerprint. Stable-ID assignment and extraction are subsequent engine work.
- SLIDE_OBJECTS stores typed object identity, kind/name/text and versioned JSON geometry/metadata; extraction/DTO validation is not implemented by the schema itself.
- SLIDE_TAGS stores original and normalized tags, unique per slide.
- SLIDE_PREVIEWS stores bounded-size PNG variants with source fingerprints, content hashes and rendering timestamps. PNG generation/content validation remains integration work.
- SLIDE_SEARCH is FTS5 with Unicode tokenization and diacritic removal. Search-term parsing/ranking/filter services remain subsequent work; raw user query syntax must never become SQL.
- INDEX_JOBS and INDEX_ERRORS track lifecycle, progress and sanitized diagnostics. A slide deletion removes its FTS entry through the schema trigger; child object/tag/preview rows cascade with foreign keys.

## Provider and resource policy

The foundation uses the Windows inbox `winsqlite3.dll`, loaded only from System32 with process-appropriate system redirection. No vendor DLL, COM registration or TCIDX parser is referenced. The x64 machine probe verified SQLite 3.53.4 and FTS5 SQL execution. This is not certification of every Windows/Office build. Missing provider/FTS5 must produce a setup failure rather than an XML-cache fallback. Older Windows and the x86 execution matrix remain unverified.

Connections use FULLMUTEX, a 5-second busy timeout, foreign_keys=ON and trusted_schema=OFF. Read-only opening is supported. SQL values are bound as UTF-8 text, blobs, signed 64-bit integers, booleans, finite doubles or null. Text/blob values are bounded to 16 MiB. Native handles/statements use SafeHandle; query rows stream through a disposable iterator. SQL calls reject additional statements before executing the first one. Errors omit SQL text and bound content.

## Evidence and remaining gates

Run `scripts/Test-PiidxFoundation.ps1` from normal supported PowerShell. Current foundation has 26 x64 assertions, including Unicode/injection isolation, FTS5, rollback/reopen, read-only enforcement, exact schema rejection and safe statement handling. It does not prove PPTX/POTX indexing, refresh, thumbnails, search service, gallery operation, insertion, interrupted rebuild recovery, installer packaging or native Office acceptance.

References: [Microsoft SQLite provider options](https://learn.microsoft.com/en-us/dotnet/standard/data/sqlite/custom-versions), [SQLite FTS5](https://www.sqlite.org/fts5.html). Runtime capability evidence is in artifacts/piidx, not inferred from documentation.