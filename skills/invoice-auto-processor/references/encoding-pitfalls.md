# MIME 邮件头编码坑点

## 问题

从 QQ 邮箱下载邮件时，部分邮件主题使用非 UTF-8 编码（如 GB2312、GBK），Python `email.header.decode_header()` 返回 `(bytes, charset)` 元组。若强制用 UTF-8 解码，中文变乱码，导致关键词匹配失败。

## 案例：12306 火车票邮件

原始 MIME 头：
```
Subject: =?gb2312?B?zfvJz7m5xrHnzbfV7LeisN+0prG...?=
```

正确解码后：`网上购票系统-电子发票通知`
强制 UTF-8 解码后：`���Ϲ�Ʊϵͳ-���ӷ�Ʊ֪ͨ`（「发票」二字被破坏）

## 正确解法：`safe_decode_header()`

```python
from email.header import decode_header

def safe_decode_header(header_value):
    """正确解码 MIME 邮件头（支持 GB2312/GBK/UTF-8 等编码）"""
    if not header_value:
        return ""
    try:
        parts = decode_header(header_value)
        result = []
        for part, charset in parts:
            if isinstance(part, bytes):
                if charset:
                    try:
                        result.append(part.decode(charset))
                    except (LookupError, UnicodeDecodeError):
                        try:
                            result.append(part.decode('gbk'))
                        except UnicodeDecodeError:
                            result.append(part.decode('utf-8', errors='replace'))
                else:
                    result.append(part.decode('utf-8', errors='replace'))
            else:
                result.append(str(part))
        return ''.join(result)
    except Exception:
        return str(header_value)
```

## 关键教训

- **不要写死 `.decode('utf-8')`**，必须根据 `decode_header()` 返回的 charset 参数选择解码器
- GB2312 和 GBK 是常见的中文邮件编码，需优先尝试
- `decode_header()` 可能返回混合编码的多个片段（如主题一半 UTF-8 一半 GB2312），需逐段处理
