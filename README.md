# EngageXml
An extracting, rebundling, xml2xlsx and xlsx2xml tool for Fire Emblem Engage game data.

## Usage

- `EngageXml.exe *.xml` : convert \*.xml to \*.xml.xlsx
  
- `EngageXml.exe *.xml.xlsx` : convert *.xml.xlsx to *.xml
  
- `EngageXml.exe -out *.(xml/txt).bundle` : extract bundle to \*.xml or \*.txt
  
- `EngageXml.exe -out -xlsx *.xml.bundle` : extract bundle to *.xml.xlsx
  
- `EngageXml.exe -in *.(xml/txt) *.bundle` : insert \*.xml or \*.txt to bundle, replacing its original TextAsset
  
- `EngageXml.exe -in *.xlsx *.xml.bundle` : convert *.xlsx to xml and insert it to bundle, replacing its original TextAsset


## Do not change sheets names, but you can create new sheets with a name starts with "#", these sheets are ignored when converted to xml.
