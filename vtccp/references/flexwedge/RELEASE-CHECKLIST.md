# Versioning and release checklist

**Document version:** 1.0.0  
**Revision date:** 2026-09-06

Release version is `1.0.0` in project, assembly, file, package ZIP, and operator documentation.

- [ ] Obtain and approve redistribution of the AsReader SDK DLL.
- [ ] Place `AsReaderP3xU.dll` in the documented SDK folder.
- [ ] Run `release/flexwedge/Publish-FlexWedge.ps1`.
- [ ] Confirm the generated ZIP contains the DLL, GCP XML, sample configuration, install/uninstall scripts, and FlexWedge executable.
- [ ] Install as a standard user and confirm desktop shortcut operation.
- [ ] Record vendor SDK version and package hash in release notes.