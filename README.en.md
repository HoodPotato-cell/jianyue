# JianYue (简阅)

A minimal, fully offline document reader for Android and Windows.

Built as a WebView (Android) / Electron (Windows) shell around a shared
rendering engine: PDF, Word, Excel, PowerPoint and Markdown, including the
legacy binary formats `.doc` / `.xls` / `.ppt`.

Everything is parsed on-device. No network calls, no uploads.

## Features

- Scans local storage, filters by type, searches by file name or folder
- PDF: vector, block-wise rendering — sharp at any zoom
- OOXML (docx / xlsx / pptx): layout reconstruction
- Legacy formats: custom parsers for Word 97, BIFF8 and PowerPoint 97
  - `.ppt` renders text, images and shape fills at their original coordinates
- Markdown: headings, code blocks, tables, quotes, lists
- Shelves (collections), multi-select, delete, drag & drop (Windows)
- Split view (side by side / stacked), light & dark frosted-glass themes

## Layout

```
android/   Web assets running inside a WebView + build scripts
windows/   Electron desktop app (reuses the same assets)
```

## Build

### Android

Requires `aapt2`, `apksigner`, `zipalign`, Python 3 and a base APK to patch.

```bash
export JY_KS=path/to/your.jks
export JY_KS_PASS=****** JY_KEY_PASS=******
bash android/build.sh
```

Signing keys are not included.

### Windows

```bash
cd windows
npm install
npm start      # dev
npm run dist   # package exe
```

## Known limitations

- Legacy `.ppt`: EMF/WMF vector graphics, gradients and animations are not reproduced
- The Android `.ppt` renderer is an approximation, not pixel-identical to PowerPoint

## License

MIT for this project. Third-party notices in [NOTICE.md](NOTICE.md).
