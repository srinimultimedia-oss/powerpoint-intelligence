# TCIDX library behavior — isolated research

## Verified

REA `inspect_byte_entropy` inspected the available Agendas sample read-only. File: `Agendas, Schedules, Timetables.potx.tcidx`; bytes: 34118; SHA256: c054734ab08964ba21e4ec55bde8d8ea74e590c151ee985eda22866672f6b3bd. Whole-file entropy: 7.9898. Retained evidence: `tcidx-entropy-evidence.json`, REA evidence ev_6003071a1fcba1503cd34ceea62b185402c64c14feb823d1f2d50afa93a606f6, parent ev_e142c0dedc16d43f80b96c1ca40365658c4be67602831ea0365bc3c709c1497a.

The original POTX is available in the copied local source library. Existing plugin source independently reads PPTX/POTX ZIP/XML text, titles, dimensions, authors and layout metadata. Existing gallery supports navigation, filtering and native preview/insertion; it does not read this TCIDX file.

## Unknown / unverified

Entropy cannot distinguish compression from encryption. No complete TCIDX format, source mapping, thumbnail storage, invalidation algorithm or semantic search behavior is established. A separate bounded LZMA experiment independently found 11 XZ signatures and decoded 11 complete streams (42,286–117,390 bytes each); the original POTX has 11 slides. See tcidx-stream-observations.json for exact offsets, lengths and hashes. The decoded streams are not UTF-8 documents; no complete schema or block-to-slide mapping is inferred. Recoverable slide-text semantics remain unverified. Entropy alone did not establish these compression observations. Static inspection is not runtime observation of the vendor UI.

The replacement uses original presentations and our documented SQLite schema. This sample and research tools are not production dependencies or deployment content. No vendor binary was executed, disassembled or copied as implementation.