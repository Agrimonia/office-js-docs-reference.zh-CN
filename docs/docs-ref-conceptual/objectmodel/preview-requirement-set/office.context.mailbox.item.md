
# <a name="item"></a>item

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a>[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item

`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) 属性确定 `item` 的类型。

##### <a name="requirements"></a>Requirements

|要求|值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|受限|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Compose 或 Read|

##### <a name="members-and-methods"></a>成员和方法

| 成员 | 类型 |
|--------|------|
| [attachments](#attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | 成员 |
| [bcc](#bcc-recipientsjavascriptapioutlookofficerecipients) | 成员 |
| [body](#body-bodyjavascriptapioutlookofficebody) | 成员 |
| [cc](#cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | 成员 |
| [conversationId](#nullable-conversationid-string) | 成员 |
| [dateTimeCreated](#datetimecreated-date) | 成员 |
| [dateTimeModified](#datetimemodified-date) | 成员 |
| [end](#end-datetimejavascriptapioutlookofficetime) | 成员 |
| [from](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) | 成员 |
| [internetMessageId](#internetmessageid-string) | 成员 |
| [itemClass](#itemclass-string) | 成员 |
| [itemId](#nullable-itemid-string) | 成员 |
| [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) | 成员 |
| [location](#location-stringlocationjavascriptapioutlookofficelocation) | 成员 |
| [normalizedSubject](#normalizedsubject-string) | 成员 |
| [notificationMessages](#notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages) | 成员 |
| [optionalAttendees](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | 成员 |
| [organizer](#organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer) | 成员 |
| [recurrence](#nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence) | 成员 |
| [requiredAttendees](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | 成员 |
| [sender](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) | 成员 |
| [可能指向](#nullable-seriesid-string) | 成员 |
| [start](#start-datetimejavascriptapioutlookofficetime) | 成员 |
| [subject](#subject-stringsubjectjavascriptapioutlookofficesubject) | 成员 |
| [to](#to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | 成员 |
| [addFileAttachmentAsync](#addfileattachmentasyncuri-attachmentname-options-callback) | 方法 |
| [addFileAttachmentFromBase64Async](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | 方法 |
| [addHandlerAsync](#addhandlerasynceventtype-handler-options-callback) | 方法 |
| [addItemAttachmentAsync](#additemattachmentasyncitemid-attachmentname-options-callback) | 方法 |
| [close](#close) | 方法 |
| [displayReplyAllForm](#displayreplyallformformdata) | 方法 |
| [displayReplyForm](#displayreplyformformdata) | 方法 |
| [getEntities](#getentities--entitiesjavascriptapioutlookofficeentities) | 方法 |
| [getEntitiesByType](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | 方法 |
| [getFilteredEntitiesByName](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | 方法 |
| [getInitializationContextAsync](#getinitializationcontextasyncoptions-callback) | 方法 |
| [getRegExMatches](#getregexmatches--object) | 方法 |
| [getRegExMatchesByName](#getregexmatchesbynamename--nullable-array-string-) | 方法 |
| [getSelectedDataAsync](#getselecteddataasynccoerciontype-options-callback--string) | 方法 |
| [getSelectedEntities](#getselectedentities--entitiesjavascriptapioutlookofficeentities) | 方法 |
| [getSelectedRegExMatches](#getselectedregexmatches--object) | 方法 |
| [getSharedPropertiesAsync](#getsharedpropertiesasyncoptions-callback) | 方法 |
| [loadCustomPropertiesAsync](#loadcustompropertiesasynccallback-usercontext) | 方法 |
| [removeAttachmentAsync](#removeattachmentasyncattachmentid-options-callback) | 方法 |
| [removeHandlerAsync](#removehandlerasynceventtype-handler-options-callback) | 方法 |
| [saveAsync](#saveasyncoptions-callback) | 方法 |
| [setSelectedDataAsync](#setselecteddataasyncdata-options-callback) | 方法 |

### <a name="example"></a>示例

以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。

```
// The initialize function is required for all apps.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    var subject = item.subject;
    // Continue with processing the subject of the current item,
    // which can be a message or appointment.
    });
}
```

### <a name="members"></a>成员

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a>attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)>

获取项目的附件数组。仅限阅读模式。

> [!NOTE]
> 某些类型的文件被阻止 Outlook 由于潜在的安全问题，并因此不返回。 有关详细信息，请参阅[在 Outlook 中被阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。

##### <a name="type"></a>类型:

*   Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)>

##### <a name="requirements"></a>要求

|要求|值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|阅读|

##### <a name="example"></a>示例

以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。

```
var _Item = Office.context.mailbox.item;
var outputString = "";

if (_Item.attachments.length > 0) {
  for (i = 0 ; i < _Item.attachments.length ; i++) {
    var _att = _Item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += _att.name;
    outputString += "<BR>ID: " + _att.id;
    outputString += "<BR>contentType: " + _att.contentType;
    outputString += "<BR>size: " + _att.size;
    outputString += "<BR>attachmentType: " + _att.attachmentType;
    outputString += "<BR>isInline: " + _att.isInline;
  }
}

// Do something with outputString
```

####  <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a>bcc :[Recipients](/javascript/api/outlook/office.recipients)

获取一个对象，提供用于获取或更新邮件的密件抄送 （密件抄送副本） 行上的收件人的方法。 仅限撰写模式。

##### <a name="type"></a>类型：

*   [收件人](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a>要求

|要求|值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.1|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|撰写|

##### <a name="example"></a>示例

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlookofficebody"></a>body :[Body](/javascript/api/outlook/office.body)

获取一个提供用于处理项目正文的方法的对象。

##### <a name="type"></a>类型：

*   [Body](/javascript/api/outlook/office.body)

##### <a name="requirements"></a>要求

|要求|值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.1|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|撰写或阅读|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a>抄送： 数组。 <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[收件人](/javascript/api/outlook/office.recipients)

提供对邮件的抄送 (cc) 收件人访问。 对象的类型和访问级别取决于当前项的模式。

##### <a name="read-mode"></a>阅读模式

`cc` 属性返回包含邮件的**抄送**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。

##### <a name="compose-mode"></a>撰写模式

`cc`属性返回`Recipients`提供方法用于获取或更新的邮件**Cc**行上的收件人对象。

##### <a name="type"></a>类型:

*   Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a>要求

|要求|值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|撰写或阅读|

##### <a name="example"></a>示例

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a>(nullable) conversationId :String

获取包含特定消息的电子邮件会话的标识符。

如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。

对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。

##### <a name="type"></a>类型:

*   String

##### <a name="requirements"></a>要求

|要求|值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|撰写或阅读|

#### <a name="datetimecreated-date"></a>dateTimeCreated :Date

获取项目创建的日期和时间。仅限阅读模式。

##### <a name="type"></a>类型：

*   日期

##### <a name="requirements"></a>要求

|要求|值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|阅读|

##### <a name="example"></a>示例

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a>dateTimeModified :Date

获取项目最近一次修改的日期和时间。仅限阅读模式。

> [!NOTE]
> 此成员不支持适用于 iOS 的 Outlook 或 Outlook for Android。

##### <a name="type"></a>类型:

*   日期

##### <a name="requirements"></a>要求

|要求|值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|阅读|

##### <a name="example"></a>示例

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlookofficetime"></a>end :Date|[Time](/javascript/api/outlook/office.time)

获取或设置约会结束的日期和时间。

将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。

##### <a name="read-mode"></a>阅读模式

`end` 属性返回 `Date` 对象。

##### <a name="compose-mode"></a>撰写模式

`end` 属性返回 `Time` 对象。

使用 [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。

##### <a name="type"></a>类型：

*   Date | [Time](/javascript/api/outlook/office.time)

##### <a name="requirements"></a>要求

|要求|值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|撰写或阅读|

##### <a name="example"></a>示例

以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的结束时间。

```
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a>从：[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[从](/javascript/api/outlook/office.from)

获取邮件发件人的电子邮件地址。

`from` 和 [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。

> [!NOTE]
> `recipientType`属性`EmailAddressDetails`对象在`from`属性是`undefined`。

##### <a name="read-mode"></a>阅读模式

`from`属性返回`EmailAddressDetails`对象。

```
var subject = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a>撰写模式

`from`属性返回`From`对象，它提供一个方法来获取值。

```
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a>类型:

*   [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [从](/javascript/api/outlook/office.from)

##### <a name="requirements"></a>要求

|要求|||
|---|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|1.7|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|ReadWriteItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Read|撰写|

#### <a name="internetmessageid-string"></a>internetMessageId :String

获取电子邮件的 Internet 消息标识符。仅限阅读模式。

##### <a name="type"></a>类型：

*   String

##### <a name="requirements"></a>要求

|要求|值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|阅读|

##### <a name="example"></a>示例

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a>itemClass :String

获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。

`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。

|类型|说明|项目类|
|---|---|---|
|约会项目|这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurence` 的日历项目。|`IPM.Appointment`<br />`IPM.Appointment.Occurence`|
|邮件项目|这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。

##### <a name="type"></a>类型：

*   String

##### <a name="requirements"></a>要求

|要求|值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|阅读|

##### <a name="example"></a>示例

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a>(nullable) itemId :String

获取当前项目的 Exchange Web 服务项目标识符。仅限阅读模式。

> [!NOTE]
> `itemId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。 `itemId`属性不是与 Outlook 条目 ID 或使用 Outlook REST API 的 ID。 使用此值的 REST API 调用之前，应将其转换使用[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string)。 有关详细信息，请参阅[使用 Outlook REST Api 从 Outlook 外接程序](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)。

`itemId` 属性在撰写模式下不可用。如果需要项目标识符，[`saveAsync`](#saveasyncoptions-callback) 方法可用于将项目保存到存储，这将在回调函数的 [`AsyncResult.value`](/javascript/api/office/office.asyncresult) 参数中返回项目标识符。

##### <a name="type"></a>类型:

*   String

##### <a name="requirements"></a>要求

|要求|值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|阅读|

##### <a name="example"></a>示例

以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a>itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)

获取实例表示的项的类型。

`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。

##### <a name="type"></a>类型：

*   [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a>要求

|要求|值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|撰写或阅读|

##### <a name="example"></a>示例

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlookofficelocation"></a>location :String|[Location](/javascript/api/outlook/office.location)

获取或设置约会的位置。

##### <a name="read-mode"></a>阅读模式

`location` 属性返回一个包含约会位置的字符串。

##### <a name="compose-mode"></a>撰写模式

`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。

##### <a name="type"></a>类型：

*   String | [Location](/javascript/api/outlook/office.location)

##### <a name="requirements"></a>要求

|要求|值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|撰写或阅读|

##### <a name="example"></a>示例

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a>normalizedSubject :String

获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。

normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) 属性。

##### <a name="type"></a>类型:

*   String

##### <a name="requirements"></a>要求

|要求|值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|阅读|

##### <a name="example"></a>示例

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a>notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)

获取项目的通知邮件。

##### <a name="type"></a>类型：

*   [NotificationMessages](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a>要求

|要求|值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.3|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|撰写或阅读|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a>optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)

提供对事件的可选与会者的访问。 对象的类型和访问级别取决于当前项的模式。

##### <a name="read-mode"></a>阅读模式

`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。

##### <a name="compose-mode"></a>撰写模式

`optionalAttendees`属性返回`Recipients`对象，它提供用于获取或更新会议的可选与会者的方法。

##### <a name="type"></a>类型：

*   Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a>要求

|要求|值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|撰写或阅读|

##### <a name="example"></a>示例

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a>组织者：[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[组织者](/javascript/api/outlook/office.organizer)

获取指定会议组织者的电子邮件地址。

##### <a name="read-mode"></a>阅读模式

`organizer`属性返回一个代表会议组织者的[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)对象。

##### <a name="compose-mode"></a>撰写模式

`organizer`属性返回一个[管理器](/javascript/api/outlook/office.organizer)对象，提供要获取组织者值的方法。

##### <a name="type"></a>类型:

*   [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [组织者](/javascript/api/outlook/office.organizer)

##### <a name="requirements"></a>要求

|要求|||
|---|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|1.7|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|ReadWriteItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Read|撰写|

##### <a name="example"></a>示例

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a>(可以为 null) 的定期：[定期](/javascript/api/outlook/office.recurrence)

获取或设置约会的定期模式。 获取会议请求的定期模式。 阅读和撰写的约会项目的模式。 读取模式的会议请求项目。

`recurrence`属性返回的定期约会或会议请求的[定期](/javascript/api/outlook/office.recurrence)对象，如果项是一系列或一系列中的实例。 `null`将返回单个约会和会议请求的单个约会。 `undefined`将返回不是会议请求的消息。

> 注意： 会议请求具有`itemClass`IPM 的值。Schedule.Meeting.Request。

> 注意： 如果定期对象是`null`，这表明该对象是一个约会或会议请求的一个约会，不属于一系列。

##### <a name="type"></a>类型:

* [定期](/javascript/api/outlook/office.recurrence)

|要求|值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.7|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|撰写或阅读|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a>requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)

提供对事件的必需与会者的访问。 对象的类型和访问级别取决于当前项的模式。

##### <a name="read-mode"></a>阅读模式

`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。

##### <a name="compose-mode"></a>撰写模式

`requiredAttendees`属性返回`Recipients`对象，它提供用于获取或更新会议的必需的与会者的方法。

##### <a name="type"></a>类型：

*   Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a>要求

|要求|值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|撰写或阅读|

##### <a name="example"></a>示例

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a>sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)

获取电子邮件发件人的电子邮件地址。仅限阅读模式。

[`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。

> [!NOTE]
> `recipientType`属性`EmailAddressDetails`对象在`sender`属性是`undefined`。

##### <a name="type"></a>类型:

*   [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a>要求

|要求|值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|阅读|

##### <a name="example"></a>示例

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a>(nullable) 可能指向： 字符串

获取实例所属的系列的 id。

在 OWA 和 Outlook 中，`seriesId`返回父 （系列） 项此项目所属的 Exchange Web Services (EWS) ID。 但是，在 iOS 和 Android，`seriesId`返回父项的 REST ID。

> [!NOTE]
> `seriesId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。 `seriesId`属性不是与 Outlook Id 使用 Outlook REST API 相同。 使用此值的 REST API 调用之前，应将其转换使用[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string)。 有关详细信息，请参阅[使用 Outlook REST Api 从 Outlook 外接程序](https://docs.microsoft.com/outlook/add-ins/use-rest-api)。

`seriesId`属性返回`null`如不具有父项的项单个系列项目、 约会或会议请求，并返回`undefined`不会议请求的其他任何项。

##### <a name="type"></a>类型:

* String

##### <a name="requirements"></a>要求

|要求|值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.7|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|撰写或阅读|

##### <a name="example"></a>示例

```
var seriesId = Office.context.mailbox.item.seriesId;
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlookofficetime"></a>start :Date|[Time](/javascript/api/outlook/office.time)

获取或设置约会开始的日期和时间。

将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) 方法将该值转换为客户端的本地日期和时间。

##### <a name="read-mode"></a>阅读模式

`start` 属性返回 `Date` 对象。

##### <a name="compose-mode"></a>撰写模式

`start` 属性返回 `Time` 对象。

使用 [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。

##### <a name="type"></a>类型：

*   Date | [Time](/javascript/api/outlook/office.time)

##### <a name="requirements"></a>要求

|要求|值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|撰写或阅读|

##### <a name="example"></a>示例

以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。

```
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

####  <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a>subject :String|[Subject](/javascript/api/outlook/office.subject)

获取或设置显示在项目的主题字段中的说明。

`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。

##### <a name="read-mode"></a>阅读模式

`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a>撰写模式

`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a>类型：

*   String | [Subject](/javascript/api/outlook/office.subject)

##### <a name="requirements"></a>要求

|要求|值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|撰写或阅读|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a>到： 数组。 <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[收件人](/javascript/api/outlook/office.recipients)

提供对邮件的**到**行上的收件人访问。 对象的类型和访问级别取决于当前项的模式。

##### <a name="read-mode"></a>阅读模式

`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。

##### <a name="compose-mode"></a>撰写模式

`to`属性返回`Recipients`提供方法用于获取或更新邮件的**到**行上的收件人对象。

##### <a name="type"></a>类型：

*   Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a>要求

|要求|值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|撰写或阅读|

##### <a name="example"></a>示例

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a>方法

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a>addFileAttachmentAsync(uri, attachmentName, [options], [callback])

将文件作为附件添加到邮件或约会。

`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。

你随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。

##### <a name="parameters"></a>参数：
|名称|类型|属性|说明|
|---|---|---|---|
|`uri`|字符串||提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。|
|`attachmentName`|字符串||在附件上载过程中显示的附件名称。最大长度为 255 个字符。|
|`options`|对象|&lt;可选&gt;|包含一个或多个以下属性的对象文本。|
|`options.asyncContext`|对象|&lt;可选&gt;|开发人员可以提供他们想要在回调方法中访问的任何对象。|
|`options.isInline`|布尔值|&lt;可选&gt;|如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。|
|`callback`|函数|&lt;可选&gt;|方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。 <br/>如果成功，附件标识符将在 `asyncResult.value` 属性中提供。<br/>如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。|

##### <a name="errors"></a>错误

|错误代码|说明|
|------------|-------------|
|`AttachmentSizeExceeded`|附件大小超过了允许的大小。|
|`FileTypeNotSupported`|该附件的扩展名不是允许的扩展名。|
|`NumberOfAttachmentsExceeded`|邮件或约会具有的附件过多。|

##### <a name="requirements"></a>要求

|要求|值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.1|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|撰写|

##### <a name="examples"></a>示例

```js
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

以下示例将图像文件添加为内联附件，并在邮件正文中引用该附件。

```js
Office.context.mailbox.item.addFileAttachmentAsync
(
  "http://i.imgur.com/WJXklif.png",
  "cute_bird.png",
  {
    isInline: true
  },
  function (asyncResult) {
    Office.context.mailbox.item.body.setAsync(
      "<p>Here's a cute bird!</p><img src='cid:cute_bird.png'>",
      {
        "coercionType": "html"
      },
      function (asyncResult) {
        
      }
    );
  }
);
```

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a>addFileAttachmentFromBase64Async (base64File，attachmentName，[选项] [回调])

添加从 base64 编码到邮件或约会作为附件的文件。

`addFileAttachmentFromBase64Async`方法从 base64 编码将文件上载，并将其附加到撰写窗体中的项。 此方法返回 AsyncResult.value 对象中的附件标识符。

你随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。

##### <a name="parameters"></a>参数：
|名称|类型|属性|说明|
|---|---|---|---|
|`base64File`|String||以 base64 编码的图像或文件添加到电子邮件或事件的内容。|
|`attachmentName`|字符串||在附件上载过程中显示的附件名称。最大长度为 255 个字符。|
|`options`|对象|&lt;可选&gt;|包含一个或多个以下属性的对象文本。|
|`options.asyncContext`|对象|&lt;可选&gt;|开发人员可以提供他们想要在回调方法中访问的任何对象。|
|`options.isInline`|布尔值|&lt;可选&gt;|如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。|
|`callback`|函数|&lt;可选&gt;|方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。 <br/>如果成功，附件标识符将在 `asyncResult.value` 属性中提供。<br/>如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。|

##### <a name="errors"></a>错误

|错误代码|说明|
|------------|-------------|
|`AttachmentSizeExceeded`|附件大小超过了允许的大小。|
|`FileTypeNotSupported`|该附件的扩展名不是允许的扩展名。|
|`NumberOfAttachmentsExceeded`|邮件或约会具有的附件过多。|

##### <a name="requirements"></a>要求

|要求|值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|预览|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|撰写|

##### <a name="examples"></a>示例

```js
Office.context.mailbox.item.addFileAttachmentFromBase64Async(
  base64String,
  "cute_bird.png",
  {
    isInline: true
  },
  function (asyncResult) {
    Office.context.mailbox.item.body.setAsync(
      "<p>Here's a cute bird!</p><img src='cid:cute_bird.png'>",
      {
        "coercionType": "html"
      },
      function (asyncResult) {
      }
    );
  }
);
```

####  <a name="addhandlerasynceventtype-handler-options-callback"></a>addHandlerAsync(eventType, handler, [options], [callback])

添加支持事件的事件处理程序。

当前受支持的事件类型的`Office.EventType.AppointmentTimeChanged`， `Office.EventType.RecipientsChanged`，和`Office.EventType.RecurrenceChanged`

##### <a name="parameters"></a>参数：

| 名称 | 类型 | 属性 | 说明 |
|---|---|---|---|
| `eventType` | [Office.EventType](office.md#eventtype-string) || 应调用处理程序的事件。 |
| `handler` | 函数 || 用于处理事件的函数。此函数必须接受一个参数，即对象文本。参数上的 `type` 属性将匹配传递给 `addHandlerAsync` 的 `eventType` 参数。 |
| `options` | Object | &lt;可选&gt; | 包含一个或多个以下属性的对象文本。 |
| `options.asyncContext` | 对象 | &lt;可选&gt; | 开发人员可以提供他们想要在回调方法中访问的任何对象。 |
| `callback` | 函数| &lt;可选&gt;|方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。|

##### <a name="requirements"></a>Requirements

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.7 |
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem |
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 撰写或阅读 |

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a>addItemAttachmentAsync(itemId, attachmentName, [options], [callback])

将 Exchange 项目（如邮件）作为附件添加到邮件或约会。

`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。

你随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。

如果您 Office 加载项运行的 Outlook Web App 中`addItemAttachmentAsync`方法可以附加到正在编辑; 的项目之外的项目的项目但是，这不受支持，并且不建议使用。

##### <a name="parameters"></a>参数：

|名称|类型|属性|说明|
|---|---|---|---|
|`itemId`|字符串||要附加的项目的 Exchange 标识符。最大长度为 100 个字符。|
|`attachmentName`|字符串||要附加的项目的主题。最大长度为 255 个字符。|
|`options`|对象|&lt;可选&gt;|包含一个或多个以下属性的对象文本。|
|`options.asyncContext`|对象|&lt;可选&gt;|开发人员可以提供他们想要在回调方法中访问的任何对象。|
|`callback`|函数|&lt;可选&gt;|方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。 <br/>如果成功，附件标识符将在 `asyncResult.value` 属性中提供。<br/>如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。|

##### <a name="errors"></a>错误

|错误代码|说明|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|邮件或约会具有的附件过多。|

##### <a name="requirements"></a>要求

|要求|值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.1|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|撰写|

##### <a name="example"></a>示例

以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。

```
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach
  // (Shortened for readability)
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

####  <a name="close"></a>close()

关闭当前正在撰写的项目。

`close` 方法的行为取决于要撰写的项目的当前状态。如果项目具有未保存的更改，客户端将提示用户保存、放弃或取消关闭操作。

> [!NOTE]
> 在 Outlook 中的 web，如果该项目是约会和它之前已保存使用在`saveAsync`、 提示用户保存、 放弃，或取消，即使未发生更改由于项目上次保存。

在 Outlook 桌面客户端中，如果邮件是内联答复，`close` 方法不起作用。

##### <a name="requirements"></a>要求

|要求|值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.3|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|受限|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|撰写|

#### <a name="displayreplyallformformdata"></a>displayReplyAllForm(formData)

显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。

> [!NOTE]
> 在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。

在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。

如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。

当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。

##### <a name="parameters"></a>参数：

|名称|类型|属性|说明|
|---|---|---|---|
|`formData`|字符串 &#124; 对象||一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。<br/>**或**<br/>包含正文或附件数据和回调函数的对象。对象定义如下。|
|`formData.htmlBody`|字符串|&lt;可选&gt;|一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。
|`formData.attachments`|Array.&lt;Object&gt;|&lt;可选&gt;|JSON 对象（是文件或项目附件）的数组。|
|`formData.attachments.type`|字符串||指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。|
|`formData.attachments.name`|字符串||一个包含附件的名称的字符串，最多包含 255 个字符。|
|`formData.attachments.url`|String||仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。|
|`formData.attachments.isInline`|Boolean||仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。|
|`formData.attachments.itemId`|String||仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。|
|`callback`|函数|&lt;可选&gt;|方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。|

##### <a name="requirements"></a>要求

|要求|值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|阅读|

##### <a name="examples"></a>示例

以下代码将一个字符串传递到 `displayReplyAllForm` 函数。

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

使用空白正文答复。

```
Office.context.mailbox.item.displayReplyAllForm({});
```

仅使用正文答复。

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

使用正文和文件附件答复。

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

使用正文和项目附件答复。

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

使用正文、文件附件、项目附件和回调答复。

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyformformdata"></a>displayReplyForm(formData)

显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。

> [!NOTE]
> 在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。

在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。

如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。

当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。

##### <a name="parameters"></a>参数：

|名称|类型|属性|说明|
|---|---|---|---|
|`formData`|字符串 &#124; 对象||一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。<br/>**或**<br/>包含正文或附件数据和回调函数的对象。对象定义如下。|
|`formData.htmlBody`|字符串|&lt;可选&gt;|一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。
|`formData.attachments`|Array.&lt;Object&gt;|&lt;可选&gt;|JSON 对象（是文件或项目附件）的数组。|
|`formData.attachments.type`|字符串||指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。|
|`formData.attachments.name`|字符串||一个包含附件的名称的字符串，最多包含 255 个字符。|
|`formData.attachments.url`|String||仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。|
|`formData.attachments.isInline`|Boolean||仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。|
|`formData.attachments.itemId`|String||仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。|
|`callback`|函数|&lt;可选&gt;|方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。|

##### <a name="requirements"></a>要求

|要求|值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|阅读|

##### <a name="examples"></a>示例

以下代码将一个字符串传递到 `displayReplyForm` 函数。

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

使用空白正文答复。

```
Office.context.mailbox.item.displayReplyForm({});
```

仅使用正文答复。

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

使用正文和文件附件答复。

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

使用正文和项目附件答复。

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

使用正文、文件附件、项目附件和回调答复。

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a>getEntities() → {[Entities](/javascript/api/outlook/office.entities)}

获取在选定的项的正文中找到的实体。

> [!NOTE]
> 在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。

##### <a name="requirements"></a>要求

|要求|值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|阅读|

##### <a name="returns"></a>返回：

类型：[Entities](/javascript/api/outlook/office.entities)

##### <a name="example"></a>示例

以下示例访问当前项的主体中的联系人实体。

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a>getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}

获取选定的项的正文中找到的指定的实体类型的所有实体的数组。

> [!NOTE]
> 在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。

##### <a name="parameters"></a>参数：

|名称|类型|说明|
|---|---|---|
|`entityType`|[Office.MailboxEnums.EntityType](/javascript/api/outlook/office.mailboxenums.entitytype)|EntityType 枚举值之一。|

##### <a name="requirements"></a>Requirements

|要求|值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|受限|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|阅读|

##### <a name="returns"></a>返回：

如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。 如果没有指定类型的实体的存在于项目的正文中，该方法将返回空数组。 否则，返回的数组中对象的类型取决于 `entityType` 参数中请求实体的类型。

当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。

|`entityType` 的值|返回的数组中对象的类型|所需权限级别|
|---|---|---|
|`Address`|String|**受限**|
|`Contact`|Contact|**ReadItem**|
|`EmailAddress`|String|**ReadItem**|
|`MeetingSuggestion`|MeetingSuggestion|**ReadItem**|
|`PhoneNumber`|PhoneNumber|**受限**|
|`TaskSuggestion`|TaskSuggestion|**ReadItem**|
|`URL`|String|**受限**|

类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>

##### <a name="example"></a>示例

下面的示例演示如何访问表示当前项目的正文中的邮政地址的字符串数组。

```
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    // Get an array of strings that represent postal addresses in the current item's body.
    var addresses = item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
    // Continue processing the array of addresses.
  });
}
```

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a>getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}

返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。

> [!NOTE]
> 在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。

`getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式的实体。

##### <a name="parameters"></a>参数：

|名称|类型|说明|
|---|---|---|
|`name`|字符串|定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。|

##### <a name="requirements"></a>要求

|要求|值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|阅读|

##### <a name="returns"></a>返回：

如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。

类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>

#### <a name="getinitializationcontextasyncoptions-callback"></a>getInitializationContextAsync([options], [callback])

当[可操作邮件激活](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message)加载项时，获取传递的初始化数据。

> [!NOTE]
> 仅支持使用此方法通过 Outlook 2016 或更高版本 Windows （单击即点即用版本高于 16.0.8413.1000） 和 Outlook web 上的 Office 365。

##### <a name="parameters"></a>参数：
|名称|类型|属性|说明|
|---|---|---|---|
|`options`|对象|&lt;可选&gt;|包含一个或多个以下属性的对象文本。|
|`options.asyncContext`|对象|&lt;可选&gt;|开发人员可以提供他们想要在回调方法中访问的任何对象。|
|`callback`|函数|&lt;可选&gt;|方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。 <br/>成功，在初始化数据提供在`asyncResult.value`属性的字符串形式。<br/>如果没有初始化上下文，`asyncResult` 对象包含 `Error` 对象，并将它的 `code` 和 `name` 属性分别设置为 `9020` 和 `GenericResponseError`。|

##### <a name="requirements"></a>要求

|要求|值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|预览|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|阅读|

##### <a name="example"></a>示例

```
// Get the initialization context (if present)
Office.context.mailbox.item.getInitializationContextAsync(
  function(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
      if (asyncResult.value != null && asyncResult.value.length > 0) {
        // The value is a string, parse to an object
        var context = JSON.parse(asyncResult.value);
        // Do something with context
      } else {
        // Empty context, treat as no context
      }
    } else {
      if (asyncResult.error.code == 9020) {
        // GenericResponseError returned when there is
        // no context
        // Treat as no context
      } else {
        // Handle the error
      }
    }
  }
);
```

#### <a name="getregexmatches--object"></a>getRegExMatches() → {Object}

返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。

> [!NOTE]
> 在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。

`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。

例如，考虑一个外接程序清单具有以下 `Rule` 元素：

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) 方法检索整个正文。

##### <a name="requirements"></a>Requirements

|要求|值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|阅读|

##### <a name="returns"></a>返回：

一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。

<dl class="param-type">

<dt>
类型</dt>


<dd>对象</dd>

</dl>

##### <a name="example"></a>示例

以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a>getregexmatchesbyname （name） → (nullable) {数组。 < 字符串 >}

返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。

> [!NOTE]
> 在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。

`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。

如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。

##### <a name="parameters"></a>参数：

|名称|类型|说明|
|---|---|---|
|`name`|字符串|定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。|

##### <a name="requirements"></a>要求

|要求|值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|阅读|

##### <a name="returns"></a>返回：

一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。

<dl class="param-type">

<dt>类型</dt>

<dd>数组。 < 字符串 ></dd>

</dl>

##### <a name="example"></a>示例

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a>getSelectedDataAsync(coercionType, [options], callback) → {String}

以异步方式返回邮件的主题或正文中选定的数据。

如果没有选定内容，但光标位于正文或主题中，此方法将会为所选数据返回 null。如果选定的是字段，而不是正文或主题，则此方法返回 `InvalidSelection` 错误。

##### <a name="parameters"></a>参数：

|名称|类型|属性|说明|
|---|---|---|---|
|`coercionType`|[Office.CoercionType](office.md#coerciontype-string)||请求数据的格式。如果为文本，则此方法返回纯文本作为字符串，删除任何显示的 HTML 标记。如果为 HTML，则此方法返回所选文本，不论是纯文本还是 HTML。|
|`options`|对象|&lt;可选&gt;|包含一个或多个以下属性的对象文本。|
|`options.asyncContext`|对象|&lt;可选&gt;|开发人员可以提供他们想要在回调方法中访问的任何对象。|
|`callback`|函数||方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。<br/><br/>若要从回调方法访问所选数据，请调用 `asyncResult.value.data`。 若要访问所选内容来自源属性，请调用`asyncResult.value.sourceProperty`，其中将为`body`或`subject`。|

##### <a name="requirements"></a>要求

|要求|值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.2|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|撰写|

##### <a name="returns"></a>返回：

作为字符串的所选数据的格式由 `coercionType` 确定。

<dl class="param-type">

<dt>
类型</dt>


<dd>字符串</dd>

</dl>

##### <a name="example"></a>示例

```
// getting selected data
Office.initialize = function () {
    Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
}

function getCallback(asyncResult) {
    var text = asyncResult.value.data;
    var prop = asyncResult.value.sourceProperty;

    Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
    // check for errors
}
```

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a>getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}

获取在用户已选择的突出显示匹配项中找到的实体。突出显示匹配项适用于[上下文外接程序](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)。

> [!NOTE]
> 在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。

##### <a name="requirements"></a>要求

|要求|值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.6|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|阅读|

##### <a name="returns"></a>返回：

类型：[Entities](/javascript/api/outlook/office.entities)

##### <a name="example"></a>示例

以下示例访问用户选择的突出显示匹配项中的地址实体。

```
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a>getSelectedRegExMatches() → {Object}

返回突出显示匹配项中匹配在清单 XML 文件中定义的正则表达式的字符串值。突出显示匹配项适用于[上下文外接程序](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)。

> [!NOTE]
> 在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。

`getSelectedRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。

例如，考虑一个外接程序清单具有以下 `Rule` 元素：

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) 方法检索整个正文。

##### <a name="requirements"></a>Requirements

|要求|值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.6|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|阅读|

##### <a name="returns"></a>返回：

一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。

##### <a name="example"></a>示例

以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。

```
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

#### <a name="getsharedpropertiesasyncoptions-callback"></a>getSharedPropertiesAsync ([选项] 回调)

获取共享的文件夹、 日历或邮箱中的所选的约会或邮件的属性。

##### <a name="parameters"></a>参数：

|名称|类型|属性|说明|
|---|---|---|---|
|`options`|对象|&lt;可选&gt;|包含一个或多个以下属性的对象文本。|
|`options.asyncContext`|对象|&lt;可选&gt;|开发人员可以提供他们想要在回调方法中访问的任何对象。|
|`callback`|函数||方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。<br/><br/>共享的属性提供为[`SharedProperties`](/javascript/api/outlook/office.sharedproperties)对象在`asyncResult.value`属性。 此对象可用于获取项目的共享的属性。|

##### <a name="requirements"></a>要求

|要求|值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|预览|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|撰写或阅读|

##### <a name="example"></a>示例

```js
Office.context.mailbox.item.getSharedPropertiesAsync(callback);
function callback (asyncResult) {
  var context=asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a>loadCustomPropertiesAsync(callback, [userContext])

异步加载所选项目上此外接程序的自定义属性。

自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。

##### <a name="parameters"></a>参数：

|名称|类型|属性|说明|
|---|---|---|---|
|`callback`|函数||方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。<br/><br/>自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook/office.customproperties) 对象提供。 此对象可用于获取、 设置，和从项目中删除自定义属性并将更改保存到自定义属性设置回服务器。|
|`userContext`|对象|&lt;可选&gt;|开发人员可以提供他们想要访问的回调函数中的任何对象。 此对象可以访问`asyncResult.asyncContext`的回调函数中的属性。|

##### <a name="requirements"></a>要求

|要求|值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|撰写或阅读|

##### <a name="example"></a>示例

以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。

```
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
  // After the DOM is loaded, add-in-specific code can run.
  var item = Office.context.mailbox.item;
  item.loadCustomPropertiesAsync(customPropsCallback);
  });
}

function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

####  <a name="removeattachmentasyncattachmentid-options-callback"></a>removeAttachmentAsync(attachmentId, [options], [callback])

将附件从邮件或约会中删除。

`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。在 Outlook Web App 和适用于设备的 OWA 中，附件标识符只在同一个会话中才有效。当用户关闭应用程序，或者如果用户开始在内嵌窗体中撰写，并在随后弹出的内嵌窗体中继续在单独的窗口撰写时，会话即结束。

##### <a name="parameters"></a>参数：

|名称|类型|属性|说明|
|---|---|---|---|
|`attachmentId`|字符串||要删除的附件的标识符。字符串的最大长度为 100 个字符。|
|`options`|对象|&lt;可选&gt;|包含一个或多个以下属性的对象文本。|
|`options.asyncContext`|对象|&lt;可选&gt;|开发人员可以提供他们想要在回调方法中访问的任何对象。|
|`callback`|函数|&lt;可选&gt;|方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。 <br/>如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。|

##### <a name="errors"></a>错误

|错误代码|说明|
|------------|-------------|
|`InvalidAttachmentId`|附件标识符不存在。|

##### <a name="requirements"></a>要求

|要求|值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.1|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|撰写|

##### <a name="example"></a>示例

以下代码删除包含标识符 '0' 的附件。

```
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  <a name="removehandlerasynceventtype-handler-options-callback"></a>removeHandlerAsync (eventType，处理程序中，[选项] [回调])

删除事件处理程序支持的事件。

当前受支持的事件类型的`Office.EventType.AppointmentTimeChanged`， `Office.EventType.RecipientsChanged`，和`Office.EventType.RecurrenceChanged`

##### <a name="parameters"></a>参数：

| 名称 | 类型 | 属性 | 说明 |
|---|---|---|---|
| `eventType` | [Office.EventType](office.md#eventtype-string) || 应调用处理程序的事件。 |
| `handler` | 函数 || 用于处理事件的函数。此函数必须接受一个参数，即对象文本。参数上的 `type` 属性将匹配传递给 `removeHandlerAsync` 的 `eventType` 参数。 |
| `options` | Object | &lt;可选&gt; | 包含一个或多个以下属性的对象文本。 |
| `options.asyncContext` | 对象 | &lt;可选&gt; | 开发人员可以提供他们想要在回调方法中访问的任何对象。 |
| `callback` | 函数| &lt;可选&gt;|方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。|

##### <a name="requirements"></a>Requirements

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.7 |
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem |
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 撰写或阅读 |

####  <a name="saveasyncoptions-callback"></a>saveAsync([options], callback)

异步保存项目。

调用时，此方法将当前邮件保存为草稿，并通过回调方法返回项目 ID。在 Outlook Web App 或 Outlook 联机模式下，该项目被保存到服务器中。在 Outlook 缓存模式下，该项目被保存到本地缓存中。

> [!NOTE]
> 如果加载项调用`saveAsync`中的项目在撰写模式下才能获取`itemId`若要使用 EWS 或 REST API，请注意，缓存模式 Outlook 时，可能需要一些时间才能项目实际同步到服务器。 直到该项目同步，使用`itemId`将返回错误。

由于约会没有草稿状态，如果以撰写模式在约会中调用 `saveAsync`，则该项将被保存为用户日历中的正常约会。对于之前未保存过的新约会，则不会发送邀请。保存现有约会将向添加或删除的与会者发送更新。

> [!NOTE]
> 以下客户端具有不同行为的`saveAsync`上约会的撰写模式下：
>
> - Mac Outlook 不支持`saveAsync`中的会议在撰写模式。 调用`saveAsync`上 Mac Outlook 中的会议，则将返回错误。
> - 在 web 上的 outlook 始终发送的邀请或更新何时`saveAsync`上约会调用在撰写模式。

##### <a name="parameters"></a>参数：

|名称|类型|属性|说明|
|---|---|---|---|
|`options`|对象|&lt;可选&gt;|包含一个或多个以下属性的对象文本。|
|`options.asyncContext`|对象|&lt;可选&gt;|开发人员可以提供他们想要在回调方法中访问的任何对象。|
|`callback`|函数||方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。<br/><br/>中提供的项标识符是成功，`asyncResult.value`属性。|

##### <a name="requirements"></a>要求

|要求|值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.3|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|撰写|

##### <a name="examples"></a>示例

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

下面是传递给回调函数的 `result` 参数的示例。`value` 属性包含的项目的项目 ID。

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a>setSelectedDataAsync(data, [options], callback)

以异步方式将数据插入到邮件的正文或主题中。

`setSelectedDataAsync` 方法将指定字符串插入到项目主题或正文的光标位置，或者，如果在编辑器中已选择文本，则该方法将替换选择的文本。如果光标不在正文或主题字段中，则返回错误。插入之后，光标会位于插入内容的末尾。

##### <a name="parameters"></a>参数：

|名称|类型|属性|说明|
|---|---|---|---|
|`data`|字符串||要插入的数据。数据不得超过 1,000,000 个字符。如果传入的数据超过 1,000,000 个字符，则会引发 `ArgumentOutOfRange` 异常。|
|`options`|Object|&lt;可选&gt;|包含一个或多个以下属性的对象文本。|
|`options.asyncContext`|对象|&lt;可选&gt;|开发人员可以提供他们想要在回调方法中访问的任何对象。|
|`options.coercionType`|[Office.CoercionType](office.md#coerciontype-string)|&lt;optional&gt;|如果为 `text`，则在 Outlook Web App 和 Outlook 中应用当前样式。如果该字段是 HTML 编辑器，则仅插入文本数据，即使数据为 HTML。<br/><br/>如果 `html` 和该字段支持 HTML（主题不支持），则在 Outlook Web App 中应用当前样式，而在 Outlook 中应用默认样式。如果该字段是文本字段，则返回 `InvalidDataFormat` 错误。<br/><br/>如果未设置 `coercionType`，则结果取决于该字段：如果该字段是 HTML，则使用 HTML；如果该字段是文本，则使用纯文本。|
|`callback`|function||方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。|

##### <a name="requirements"></a>Requirements

|要求|值|
|---|---|
|[最低版本的邮箱要求集](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.2|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)|撰写|

##### <a name="example"></a>示例

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```