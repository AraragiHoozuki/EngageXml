# EngageXml
火焰纹章 Engage 解包，打包工具

## 用法

- **NEW** `EngageXml.exe -update [-o] xlsx_or_xml_path bundle_path`: 用给定的 xml 或 xlsx 更新 bundle，以 Param 的第二个 Attribute（一般为 XID） 为依据，将 xml 或 xlsx 中存在，但 bundle 中不存在的条目添加进 bundle； 如果加上 `-o` 选项，则还会覆盖同 ID 的条目。  

- **NEW** `EngageXml.exe *.bundle` （或将文件拖动到程序上）: 解包 bundle，支持 xml, txt, bytes(message).

- `EngageXml.exe *.xml` （或将文件拖动到程序上）: 将 xml 转换为 xlsx
  
- `EngageXml.exe *.xml.xlsx` （或将文件拖动到程序上）: 将 xlsx 转换为 xml
  
- `EngageXml.exe -out *.(xml/txt).bundle` : 同 `EngageXml.exe *.bundle`
  
- `EngageXml.exe -out -xlsx *.xml.bundle` : 解包 bundle 到 xml 并转换成 xlsx
  
- `EngageXml.exe -in *.(xml/txt/xlsx/bytes) *.bundle` : 将文件插入到 bundle，支持 xml, txt, xlsx(程序会自动将其转换成xml再插入), bytes(message), 替换原本的数据
