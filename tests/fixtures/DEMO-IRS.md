# 归档演示文件接口测试基线

编号：DEMO-IRS；版本：1.0；日期：2026-09-15；编写单位：公开示例。

本文件是虚构回归试验输入，范围由DEMO-SSS第1章定义。

## 1 范围

规定文件提供端、DEMO-ARCHIVE与结果接收端的接口关系及可解析数据格式。本文件不代表现实软件具有这些接口，也不为其他项目新增约束。

## 2 引用

DEMO-SSS《归档演示系统需求测试基线》，1.0版，2026-09-15，公开示例。获取位置：同目录DEMO-SSS.md。

## 3 接口关系和固定定义

```mermaid
flowchart LR
    Provider[输入文件提供端] -->|IF-INGEST 归档文件输入| CSCI[DEMO-ARCHIVE]
    CSCI -->|IF-EXPORT 处理结果输出| Consumer[结果文件接收端]
```

两个接口均固定为测试基线1.0。本地完整文件交换无网络会话；一份输入对应一次提交，输出文件通过试验提交记录与输入关联。

### 3.1 IF-INGEST

输入为无BOM的UTF-8 JSON文件，每个文件恰含一个对象且仅含两个字符串字段：record_id与payload，字段顺序不限。record_id不作空白裁剪，空字符串表示空编号；payload允许空字符串。合法归档命令要求record_id非空且尚未存在于归档集合。

合法示例：`{"record_id":"A","payload":"正文"}`。空编号示例：`{"record_id":"","payload":"正文"}`。空编号通过IF-EXPORT返回EMPTY_ID，不创建归档记录。

### 3.2 IF-EXPORT

输出为无BOM的UTF-8 JSON文件，每个文件恰含一个对象且仅含两个字符串字段：record_id与status，字段顺序不限。record_id保持原始输入值；成功归档的status为SUCCESS，空编号拒绝的status为EMPTY_ID。

成功示例：`{"record_id":"A","status":"SUCCESS"}`。空编号示例：`{"record_id":"","status":"EMPTY_ID"}`。BUSY期间被拒绝的当前编号重复命令不产生额外结果文件；原命令仍完成一次。结果目录可写，写入故障不在该测试基线的验收范围，未定义WRITE_ERROR返回承诺。

### 3.3 内部归档数据与接口的一致性

成功归档时保存record_id和status=SUCCESS；空编号或重复命令的拒绝不新增记录。记录编号在本次运行内唯一且可读取保存值。本文件未分配跨运行持久化或保留时长要求。
