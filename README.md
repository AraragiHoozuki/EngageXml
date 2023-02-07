# EngageXml
An extracting, rebundling, xml2xlsx and xlsx2xml tool for Fire Emblem Engage game data.

## Usage

**If your files are not in the same folder as EngageXml.exe, type their full path in cmd. Eg. `EngageXml.exe -out "C:\Users\[UserName]\AppData\Roaming\Ryujinx\mods\contents\0100a6301214e000\ModName\romfs\Data\StreamingAssets\aa\Switch\fe_assets_gamedata\skill.xml.bundle"`**

- `EngageXml.exe *.xml` : convert \*.xml to \*.xml.xlsx
  
- `EngageXml.exe *.xml.xlsx` : convert *.xml.xlsx to *.xml
  
- `EngageXml.exe -out *.(xml/txt).bundle` : extract bundle to \*.xml or \*.txt
  
- `EngageXml.exe -out -xlsx *.xml.bundle` : extract bundle to *.xml.xlsx
  
- `EngageXml.exe -in *.(xml/txt) *.bundle` : insert \*.xml or \*.txt to bundle, replacing its original TextAsset
  
- `EngageXml.exe -in *.xlsx *.xml.bundle` : convert *.xlsx to xml and insert it to bundle, replacing its original TextAsset


## Do not change sheets names, but you can create new sheets with a name starts with "#", these sheets are ignored when converted to xml.
