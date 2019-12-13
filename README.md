# vba-walidacja-podatnikow

Checking taxpayers based on the NIP number using the WL Register API

### Prerequisites
Import VBA-JSON module from https://github.com/VBA-tools/VBA-JSON

### Usage
Import VBA-JSON and `walidacja-podatnikow.bas` modules to your Excel file or just use `walidacja-podatnikow.xlsm` file instead.

When the import of modules is completed you can use `Podmiot*` functions like any other in this Excel file i.e.
```
=PodmiotStatus([@NIP])
=PodmiotNazwa([@NIP])
=PodmiotRegon([@NIP])
=PodmiotPesel([@NIP])
=PodmiotKrs([@NIP])
=PodmiotDataRejestracji([@NIP])
=PodmiotKontaBankowe([@NIP])
```

### How to use it in all workbooks?
If you want to use these functions in all workbooks:
1. Download `walidacja-podatnikow.xlam` file.
2. Add downloaded file to `C:\Users\<username>\AppData\Roaming\Microsoft\AddIns` folder.
3. Run Excel application.
4. Click the File tab, click Options, and then click the Add-Ins category.
5. In the Manage box, click Excel Add-ins, and then click Go. The Add-Ins dialog box appears.
6. In the Add-Ins available box, select the check box next to the add-in that you want to activate, and then click OK.
7. Enjoy.