# EngageXml
火焰纹章 Engage 解包，打包工具。
程序未经全面测试，自行斟酌使用

## 用法

- **NEW** `EngageXml.exe *.bundle` （或将文件拖动到程序上）: 解包 bundle，支持 xml, txt, bytes(message).

- `EngageXml.exe *.xml` （或将文件拖动到程序上）: 将 xml 转换为 xlsx
  
- `EngageXml.exe *.xml.xlsx` （或将文件拖动到程序上）: 将 xlsx 转换为 xml
  
- `EngageXml.exe -out *.(xml|txt).bundle` : 同 `EngageXml.exe *.bundle`
  
- `EngageXml.exe -out -xlsx *.xml.bundle` : 解包 bundle 到 xml 并转换成 xlsx
  
- `EngageXml.exe -in *.(xml|txt|xlsx|bytes) *.bundle` : 将文件插入到 bundle，支持 xml, txt, xlsx(程序会自动将其转换成xml再插入), bytes(message), 替换原本的数据
