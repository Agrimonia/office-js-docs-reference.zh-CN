 

# <a name="office"></a>Office

该 Office 命名空间提供所有 Office 应用中的外接程序所使用的共享接口。此列表仅记录 Outlook 外接程序所使用的接口。有关 Office 命名空间的完整列表，请参阅[共享 API](/javascript/api/office)。

##### <a name="requirements"></a>Requirements

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Compose 或 Read|

##### <a name="members-and-methods"></a>成员和方法

| 成员 | 类型 |
|--------|------|
| [AsyncResultStatus](#asyncresultstatus-string) | 成员 |
| [CoercionType](#coerciontype-string) | 成员 |
| [EventType](#eventtype-string) | 成员 |
| [SourceProperty](#sourceproperty-string) | 成员 |

### <a name="namespaces"></a>命名空间

[上下文](office.context.md)： 用于 Outlook 外接程序 API 提供共享的 Office 加载项 API 上下文命名空间中的接口。

[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype)：包括 ItemType、EntityType、AttachmentType、RecipientType、ResponseType 和 ItemNotificationMessageType 枚举。

### <a name="members"></a>成员

####  <a name="asyncresultstatus-string"></a>AsyncResultStatus :String

指定异步调用的结果。

##### <a name="type"></a>类型：

*   字符串

##### <a name="properties"></a>属性：

|名称| 类型| 说明|
|---|---|---|
|`Succeeded`| String|调用成功。|
|`Failed`| 字符串|调用失败。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 撰写或阅读|

---

####  <a name="coerciontype-string"></a>CoercionType :String

指定如何强制由调用方法返回或设置的数据。

##### <a name="type"></a>类型：

*   字符串

##### <a name="properties"></a>属性：

|名称| 类型| 说明|
|---|---|---|
|`Html`| String|请求以 HTML 格式返回的数据。|
|`Text`| 字符串|请求以文本格式返回的数据。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 撰写或阅读|

---

####  <a name="eventtype-string"></a>EventType :String

指定与事件处理程序相关联的事件。

##### <a name="type"></a>类型：

*   字符串

##### <a name="properties"></a>属性：

| 名称 | 类型 | 说明 | 最低要求集 |
|---|---|---|---|
|`AppointmentTimeChanged`| String | 已更改的日期或时间所选的约会或系列。 | 1.7 |
|`ItemChanged`| 字符串 | 选定的项已更改。 | 1.5 |
|`RecipientsChanged`| String | 已更改所选的项目或约会位置的收件人列表。 | 1.7 |
|`RecurrenceChanged`| String | 所选系列的定期模式已更改。 | 1.7 |

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.5 |
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 撰写或阅读 |

---

####  <a name="sourceproperty-string"></a>SourceProperty :String

指定由调用方法返回的数据源。

##### <a name="type"></a>类型：

*   字符串

##### <a name="properties"></a>属性：

|名称| 类型| 说明|
|---|---|---|
|`Body`| 字符串|数据源来自邮件的正文。|
|`Subject`| String|数据源来自邮件的主题。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 撰写或阅读|