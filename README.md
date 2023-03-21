# EngageXml
火焰纹章 Engage 解包，打包工具。
程序未经全面测试，自行斟酌使用

## 新增指令 update
`EngageXml.exe -update [-o] *.xml|xlsx|bundle  *.xml.bundle [ids...]` 将指定的 xml、xlsx 或 bundle 更新到目标 bundle

### 参数说明：
 - `-o` 可选, 表示覆盖重复的条目，不写该参数则维持原样
 - ids, 条目判断依据，可以是一个或多个字符串，程序将根据该项判定你的条目和目标bundle中的条目是否为同一个
 
### 举例：
`EngageXml.exe -update -o skill.xml.xlsx  skill.xml.bundle Sid` : 
将 skill.xml.xlsx 中的条目全部写入 skill.xml.bundle, 如果遇到 Sid 相同的条目，则视为重复，覆盖原来的。

`EngageXml.exe -update -o assettable.xml.xlsx  assettable.xml.bundle Mode Conditions` : 
将 assettable.xml.xlsx 中的条目全部写入 assettable.xml.bundle, 如果遇到 Mode 和 Conditions 均相同的条目，则视为重复，覆盖原来的。

如果指定的 ids 中，有某项为空，则程序会视为继承之前的，如：
|  Ggid   | Level  |
|  ----  | ----  |
|  GGID_マルス  | |
|   | 1 |
|   | 2 |

则 Level 为 1，2的条目的 Ggid 也视为 GGID_マルス 

## 新增快速 Message 修改
`EngageXml.exe -in *.csv *.bundle` 可以将 csv 文件插入 message bundle 中。csv 文件应该如下书写：

```csv
KEY, VALUE
```

如果 bundle 中已经有 KEY 存在，会覆盖原条目。如果在 csv 中不输入 VALUE, 会删除 bundle 中的 KEY 条目。

分隔符号 `,` 前后的空格会被无视

MSBT 文件解析来自: https://github.com/IcySon55/3DLandMSBTeditor

## 用法

- **NEW** `EngageXml.exe *.bundle` （或将文件拖动到程序上）: 解包 bundle，支持 xml, txt, bytes(message).

- `EngageXml.exe *.xml` （或将文件拖动到程序上）: 将 xml 转换为 xlsx
  
- `EngageXml.exe *.xml.xlsx` （或将文件拖动到程序上）: 将 xlsx 转换为 xml
  
- `EngageXml.exe -out *.(xml|txt).bundle` : 同 `EngageXml.exe *.bundle`
  
- `EngageXml.exe -out -xlsx *.xml.bundle` : 解包 bundle 到 xml 并转换成 xlsx
  
- `EngageXml.exe -in *.(csv|xml|txt|xlsx|bytes) *.bundle` : 将文件插入到 bundle，支持 xml, txt, xlsx(程序会自动将其转换成xml再插入), bytes(message), 替换原本的数据

