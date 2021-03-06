# <a name="outlook-javascript-api-requirement-sets"></a>Outlook 的 JavaScript API 要求集

Outlook 加载项声明要求通过使用其[清单](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)中的[Requirements](/javascript/office/manifest/requirements)元素哪些 API 版本。 Outlook 加载项始终包括 [Set](/javascript/office/manifest/set) 元素，其中 `Name` 属性设置为 `Mailbox` 且 `MinVersion` 属性设置为支持加载项方案的 API 最低要求集。

例如，下面的清单段表示 1.1 的最低要求集：

```xml
<Requirements>
  <Sets>
    <Set Name="MailBox" MinVersion="1.1" />
  </Sets>
</Requirements>
```

所有 Outlook API 都属于 `Mailbox` [要求集](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)。 `Mailbox` 要求集具有多个版本，并且我们发布的每组新的 API 都属于更高版本的要求集。 并非所有的 Outlook 客户端都支持最新的 API 集，但如果 Outlook 客户端声明支持要求集，则它支持该要求集中的所有 API。

在清单中设置最低要求集版本可控制外接程序会显示在哪个 Outlook 客户端中。如果客户端不支持最低要求集，则不会加载外接程序。例如，如果指定要求集版本 1.3，则意味着外接程序不会显示在任何不支持 1.3 及以上版本的 Outlook 客户端中。

## <a name="using-apis-from-later-requirement-sets"></a>使用更高版本要求集中的 API

设置要求集不限制外接程序可以使用的可用 Api。 例如，如果外接程序指定要求设置 1.1，但它运行中的 Outlook 客户端支持 1.3 外, 接程序可以使用 Api 要求集中 1.3。

若要使用较新的 Api，开发人员可以只检查它们存在通过使用标准 JavaScript 技术：

```js
if (item.somePropertyOrFunction !== undefined) {
  item.somePropertyOrFunction ...
}
```

对于清单中所指定的要求集版本中的任何 API，无需执行此类检查。

## <a name="choosing-a-minimum-requirement-set"></a>选择最低要求集

开发人员应使用包含其方案关键 API 集的最早要求集，如果不使用该要求集，外接程序将不起作用。

## <a name="clients"></a>客户端

下列客户端支持 Outlook 外接程序。

| 客户端 | 受支持的 API 要求集 |
| --- | --- |
| 用于 Windows 的 outlook 2019 | 1.1、 1.2、 1.3、 1.4、 1.5、 1.6 |
| Outlook 2019 for Mac | 1.1、 1.2、 1.3、 1.4、 1.5、 1.6 |
| Outlook 2016（即点即用）for Windows | 1.1、 1.2、 1.3、 1.4、 1.5、 1.6、 1.7 |
| Outlook 2016 (MSI) for Windows | 1.1、1.2、1.3、1.4 |
| Outlook 2016 for Mac | 1.1、 1.2、 1.3、 1.4、 1.5、 1.6 |
| Outlook 2013 for Windows | 1.1、1.2、1.3、1.4 |
| Outlook for iPhone | 1.1, 1.2, 1.3, 1.4, 1.5 |
| Outlook for Android | 1.1, 1.2, 1.3, 1.4, 1.5 |
| Outlook 网页版（Office 365 和 Outlook.com） | 1.1、 1.2、 1.3、 1.4、 1.5、 1.6 |
| Outlook Web App （Exchange 2013 内部部署） | 1.1 |
| Outlook Web App （Exchange 2016 本地） | 1.1, 1.2. 1.3 |

> [!NOTE]
> 在 Outlook 2013 中的 1.3 支持已添加的[年 12 月 8，2015，Outlook 2013 (KB3114349) 的更新的](https://support.microsoft.com/kb/3114349)一部分。 在 Outlook 2013 中的 1.4 支持已添加的[年 9 月 13，2016，Outlook 2013 (KB3118280) 的更新的](https://support.microsoft.com/help/3118280)一部分。
