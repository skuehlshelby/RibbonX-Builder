<img src="ribbons.png" alt="RibbonX Builder">
A series of wrapper objects for RibbonX controls. These wrapper objects provide a convenient .NET API for manipulating custom ribbon items. Also provides RibbonX callback signatures in the form of delegates, a working Bitmap to IPictureDisp converter, among other ribbon-related conveniences.

## What's in the box

The library takes you end-to-end — from describing a ribbon to loading it in Office:

- **Control wrappers & a fluent builder.** Compose tabs, groups, buttons, galleries, drop-downs, and the rest through a strongly-typed builder API instead of hand-writing RibbonX XML. Callback signatures are exposed as delegates, and a Bitmap → IPictureDisp converter handles runtime images.

- **A ready-made add-in entry point.** `CustomRibbonBase` implements the COM interfaces Office expects (`IDTExtensibility2` for the add-in lifecycle and `IRibbonExtensibility` for `GetCustomUI`), and wires the ribbon callbacks for you. Inherit it, override `BuildRibbon`, and you have a working add-in. It also composes any number of `IDTExtensibility2` modules, fanning each lifecycle event out to all of them.

- **A registration tool.** A small WPF utility (`RegistrationUtility`) that points at your compiled DLL, reflects out the COM-visible add-in classes, and writes the COM + Office add-in registry entries for you — no `regasm`, no hand-edited registry keys. It registers under the current user (`HKEY_CURRENT_USER`), so no administrator rights are required, and offers a matching un-register.
