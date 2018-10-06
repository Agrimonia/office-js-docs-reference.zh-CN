# <a name="allowsnapshot-element"></a>AllowSnapshot 元素

指定是否将内容外接程序的快照图像与主机文档一起保存。

**外接程序类型：** 内容

## <a name="syntax"></a>语法

```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```

## <a name="contained-in"></a>包含在

[OfficeApp](officeapp.md)

## <a name="remarks"></a>注解

 > [!IMPORTANT]
 > 默认情况下 **AllowSnapshot** 为 `true` 。 这样，用户使用不支持 Office 外接程序的本地应用版本打开文档时，即可看到该外接程序的图像，或者如果本地应用程序无法连接到托管外接程序的服务器时，会提供该外接程序的静态图像。 但是，这也意味着可以直接从托管该外接程序的文档访问显示在外接程序中的潜在敏感信息。

