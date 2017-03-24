Make My Manifest Version 0.9.305 Read Me - July 2011

This is the final distribution of MMM.  The software is
copyrighted but permission is granted for broad usage by
anyone, including the development of deritivate works
without requesting permission.  

Robert Riemersma, July 23, 2011


The changes since the "0.6.7" release:

NEWER CHANGES CAN BE FOUND IN THE SOURCE CODE RELATED
DOCUMENT:

   (Project root)\Related\MMMChangeLog.txt

OLDER CHANGE HISTORY:

** Build 0.7.300 August 2009 **

o  Code Refactoring

   Significant internal changes have been made to the logic
   within MMM itself as part of progress toward some new
   features.  Those will include the ability to reprocess
   an MMMP file from the command line for automated builds
   or to re-open an MMMP in the GUI to make and save changes.
   These features are not ready yet, but preparation has
   required lots of code rearrangement.  Be sure to keep a
   known-reliable 0.6.x version of MMM around for now in
   case of problems.

o  Version Numbering

   Why this matters to people I can't be sure, but there have
   been many requests to expose the MMM "build" number as part
   of its versioning.  Beginning with this release MMM will
   use just two-part versions, allowing the third part of the
   VB6 version to be the free-running auto incremented "build"
   value.

   The continuing beta will be numbered 0.x, where x increases
   for each released version.  Thus the successor to 0.6.7 is
   0.7.300, and the next major beta version would be 0.8.xxx,
   and so on.

   Most people won't care a bit.

o  Manifest Cleanup

   I have stripped some unnecessary text from the MMM-created
   application manifests.  The trustInfo section is a little
   more compact and cleaner now.

   Please make sure this doesn't cause you problems and let
   me know if it does.  Keep 0.6.7 or 0.6.6 around just in
   case!

   MMM adds a short XML comment to the manifest which notes
   the MMM version used to create it.

o  Use Correct Decimal Point Character

   There are a few places where MMM creates fractional number
   values as Strings and later converts them to Single for
   use.  These could cause a fatal exception in some locales,
   making MMM useless.  This has been corrected.

o  Capture Path32

   Some people do not compile their EXEs into the Project
   folder.  MMM will now extract and use the Path32 key's
   value from .VBPs that have it.

o  MMM Log Format

   Very minor cosmetic changes to the format of the
   information MMM writes to its activity log.

** Build 0.7.302 September 2009 **

o  Changed the logic for processing Path32 to work on the
   unsupported non-Unicode systems (Win9x).  Cautions have
   been posted about this on MSDN Community Comments so new
   problems may arise.  So far in my testing everything
   seems ok.

   People, *please* move off Win9x ASAP.  It is becoming
   more difficult every year to keep programs running
   properly on these unsupported OSs.

   Note: Windows 2000 SP4 falls off support July 2010 as
         well.

         Even now (September 2009) Microsoft has declined to
         produce a fix for three newly discovered security
         vulnerabilities in Win2K SP4!

o  Fixed Path32 handling to allow for cases where the EXE
   location is on a different drive than the .VBP file
   itself!

o  The KB921337 "two XML schema properties" issue was
   reintroduced along with the <dpiAware> manifest section.
   This appears to have been corrected now, but needs more
   testing on WinXP SP2.

** Build 0.7.303 September 2009 **

o  Added UPNP.DLL to exclusions in INI.

o  Resolved a question re. exclusion of msscript.ocx (no
   change made, it will not be hard-excluded by MMM).

** Build 0.8.304 March 2011 **

o  UIAccess true/false were being localized.  Now only the
   explicit "true" or "false" are generated in the XML
   manifest.

o  Non-COM and no DEPS Folder.  When people tell MMM not to
   use a dependencies folder (always a bad idea - may result
   in a VB6 program corrupting the target system's registry)
   non-COM DLLs should not be redirected because they'll
   appear with no path since they are in the EXE's path.
   Now MMM will not generate <FILE> entries for such DLLs
   located "next to" the EXE.

o  The Log panel TextBox now uses BigTextBox semantics
   making it harder to overflow the Log text.

** Build 0.9.305 April 2011 **

o  Finally closing in on the problem with "comment" strings
   in type library info that produces XML syntax errors.
   Most of the time these appear to be faulty typelibs in
   poor quality libraries (no names).  One can only wonder
   how many *serious* bugs are lrking in such shoddy
   software however.  Maybe the MMM manifest syntax errors
   should be taken as a warning!

   Such strings are now trimmed at the first NUL, and any
   character outside &H20 - &H7E is zapped to an "_"
   (underscore).  It is just too clumsy trying to figure
   out which accented characters are valid.

** Build 0.10.306 October 2011 **

o  DPI-Aware manifest fragement revised after persistent
   complaints that I haven't been looking at the *actual
   problem* correctly.

   More exhaustive testing showed that not all versions of
   Windows handle the manifest namespaces the same way.  I
   should have been more aware of this myself since we've
   seen similar issues in other parts of the manifest.

   Now using asmv3 for the <windowsSettings/> tag but asmv1
   for the actual <dpiAware/> tag.  This *appears* to
   resolve the issue at last.

** Build 0.11.307 February 2013 **

o  More asmv3 changes to manifest related to the DPI-Aware
   node.  Changes to 3 manifest fragments:

      manifest.apphead.txt
      manifest.dpiaware.txt
      manifest.trustinfo.txt

o  Changed LoadRefByProjFileRef() in LibData.cls module.
   Actual .TLB files are now handled better and are
   marked "IncludedNever" and NOT "Included" and generally
   blocked from accidental or intentional inclusion in
   the resulting package.

   If you need a TLB at runtime for some reason, copy it
   into the package yourself as an included file.
