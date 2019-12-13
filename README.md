# vba-walidacja-podatnikow

Checking taxpayers based on the NIP number using the WL Register API

### Prerequisites
Import VBA-JSON module from https://github.com/VBA-tools/VBA-JSON

### Usage
Import VBA-JSON and vba-walidacja-podatnikow.bas modules to your Excel file or just use vba-walidacja-podatnikow.xlsm file instead.

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
