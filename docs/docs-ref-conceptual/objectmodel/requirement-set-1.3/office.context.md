
# <a name="context"></a><span data-ttu-id="839c0-101">context</span><span class="sxs-lookup"><span data-stu-id="839c0-101">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="839c0-102">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="839c0-102">[Office](Office.md).context</span></span>

<span data-ttu-id="839c0-p101">Office.context 命名空间提供所有 Office 应用中的外接程序所使用的共享接口。此列表仅记录 Outlook 外接程序所使用的接口。有关 Office.context 命名空间的完整列表，请参阅[共享 API 中的 Office.context 引用](/javascript/api/office/office.context)。</span><span class="sxs-lookup"><span data-stu-id="839c0-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Shared API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="839c0-105">要求</span><span class="sxs-lookup"><span data-stu-id="839c0-105">Requirements</span></span>

|<span data-ttu-id="839c0-106">要求</span><span class="sxs-lookup"><span data-stu-id="839c0-106">Requirement</span></span>| <span data-ttu-id="839c0-107">值</span><span class="sxs-lookup"><span data-stu-id="839c0-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="839c0-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="839c0-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="839c0-109">1.0</span><span class="sxs-lookup"><span data-stu-id="839c0-109">1.0</span></span>|
|[<span data-ttu-id="839c0-110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="839c0-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="839c0-111">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="839c0-111">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="839c0-112">命名空间</span><span class="sxs-lookup"><span data-stu-id="839c0-112">Namespaces</span></span>

<span data-ttu-id="839c0-113">[邮箱](office.context.mailbox.md)： Microsoft Outlook 和 Microsoft Outlook on the web 提供对 Outlook 加载项对象模型的访问。</span><span class="sxs-lookup"><span data-stu-id="839c0-113">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

### <a name="members"></a><span data-ttu-id="839c0-114">成员</span><span class="sxs-lookup"><span data-stu-id="839c0-114">Members</span></span>

####  <a name="displaylanguage-string"></a><span data-ttu-id="839c0-115">displayLanguage :String</span><span class="sxs-lookup"><span data-stu-id="839c0-115">displayLanguage :String</span></span>

<span data-ttu-id="839c0-116">获取用户针对 Office 主机应用程序的 UI 指定的 RFC 1766 语言标记格式的区域设置（语言）。</span><span class="sxs-lookup"><span data-stu-id="839c0-116">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="839c0-117">`displayLanguage` 值反映在 Office 主机应用程序中通过“**文件 > 选项 > 语言**”指定的当前“**显示语言**”设置。</span><span class="sxs-lookup"><span data-stu-id="839c0-117">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="839c0-118">类型：</span><span class="sxs-lookup"><span data-stu-id="839c0-118">Type:</span></span>

*   <span data-ttu-id="839c0-119">String</span><span class="sxs-lookup"><span data-stu-id="839c0-119">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="839c0-120">要求</span><span class="sxs-lookup"><span data-stu-id="839c0-120">Requirements</span></span>

|<span data-ttu-id="839c0-121">要求</span><span class="sxs-lookup"><span data-stu-id="839c0-121">Requirement</span></span>| <span data-ttu-id="839c0-122">值</span><span class="sxs-lookup"><span data-stu-id="839c0-122">Value</span></span>|
|---|---|
|[<span data-ttu-id="839c0-123">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="839c0-123">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="839c0-124">1.0</span><span class="sxs-lookup"><span data-stu-id="839c0-124">1.0</span></span>|
|[<span data-ttu-id="839c0-125">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="839c0-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="839c0-126">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="839c0-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="839c0-127">示例</span><span class="sxs-lookup"><span data-stu-id="839c0-127">Example</span></span>

```js
function sayHelloWithDisplayLanguage() {
  var myDisplayLanguage = Office.context.displayLanguage;
  switch (myDisplayLanguage) {
    case 'en-US':
      write('Hello!');
      break;
    case 'en-NZ':
      write('G\'day mate!');
      break;
  }
}
// Function that writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

####  <a name="officetheme-object"></a><span data-ttu-id="839c0-128">officeTheme :Object</span><span class="sxs-lookup"><span data-stu-id="839c0-128">officeTheme :Object</span></span>

<span data-ttu-id="839c0-129">提供了访问 Office 主题颜色的属性。</span><span class="sxs-lookup"><span data-stu-id="839c0-129">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="839c0-130">此成员不支持适用于 iOS 的 Outlook 或 Outlook for Android。</span><span class="sxs-lookup"><span data-stu-id="839c0-130">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="839c0-p102">通过使用 Office 主题颜色，你可以使外接程序的配色方案与用户（通过 **“文件”>“Office 帐户”>“Office 主题”UI**）选择的当前 Office 主题协调一致，这种做法适用于所有 Office 主机应用程序。使用 Office 主题颜色适用于邮件和任务窗格外接程序。</span><span class="sxs-lookup"><span data-stu-id="839c0-p102">Using Office theme colors let's you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications. Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="839c0-133">类型：</span><span class="sxs-lookup"><span data-stu-id="839c0-133">Type:</span></span>

*   <span data-ttu-id="839c0-134">对象</span><span class="sxs-lookup"><span data-stu-id="839c0-134">Object</span></span>

##### <a name="properties"></a><span data-ttu-id="839c0-135">属性：</span><span class="sxs-lookup"><span data-stu-id="839c0-135">Properties:</span></span>

|<span data-ttu-id="839c0-136">名称</span><span class="sxs-lookup"><span data-stu-id="839c0-136">Name</span></span>| <span data-ttu-id="839c0-137">类型</span><span class="sxs-lookup"><span data-stu-id="839c0-137">Type</span></span>| <span data-ttu-id="839c0-138">描述</span><span class="sxs-lookup"><span data-stu-id="839c0-138">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="839c0-139">String</span><span class="sxs-lookup"><span data-stu-id="839c0-139">String</span></span>|<span data-ttu-id="839c0-140">获取十六进制三原色形式的 Office 主题正文背景色。</span><span class="sxs-lookup"><span data-stu-id="839c0-140">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="839c0-141">String</span><span class="sxs-lookup"><span data-stu-id="839c0-141">String</span></span>|<span data-ttu-id="839c0-142">获取十六进制三原色形式的 Office 主题正文前景色。</span><span class="sxs-lookup"><span data-stu-id="839c0-142">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="839c0-143">字符串</span><span class="sxs-lookup"><span data-stu-id="839c0-143">String</span></span>|<span data-ttu-id="839c0-144">获取十六进制三原色形式的 Office 主题控制背景色。</span><span class="sxs-lookup"><span data-stu-id="839c0-144">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="839c0-145">字符串</span><span class="sxs-lookup"><span data-stu-id="839c0-145">String</span></span>|<span data-ttu-id="839c0-146">获取十六进制三原色形式的 Office 主题正文控制颜色。</span><span class="sxs-lookup"><span data-stu-id="839c0-146">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="839c0-147">要求</span><span class="sxs-lookup"><span data-stu-id="839c0-147">Requirements</span></span>

|<span data-ttu-id="839c0-148">要求</span><span class="sxs-lookup"><span data-stu-id="839c0-148">Requirement</span></span>| <span data-ttu-id="839c0-149">值</span><span class="sxs-lookup"><span data-stu-id="839c0-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="839c0-150">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="839c0-150">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="839c0-151">1.3</span><span class="sxs-lookup"><span data-stu-id="839c0-151">1.3</span></span>|
|[<span data-ttu-id="839c0-152">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="839c0-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="839c0-153">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="839c0-153">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="839c0-154">示例</span><span class="sxs-lookup"><span data-stu-id="839c0-154">Example</span></span>

```js
function applyOfficeTheme(){
  // Get office theme colors.
  var bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
  var bodyForegroundColor = Office.context.officeTheme.bodyForegroundColor;
  var controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor
  var controlForegroundColor = Office.context.officeTheme.controlForegroundColor;

  // Apply body background color to a CSS class.
  $('.body').css('background-color', bodyBackgroundColor);
}
```

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlook13officeroamingsettings"></a><span data-ttu-id="839c0-155">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_3/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="839c0-155">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_3/office.RoamingSettings)</span></span>

<span data-ttu-id="839c0-156">获取一个对象，它表示保存到用户邮箱的邮件外接程序的自定义设置或状态。</span><span class="sxs-lookup"><span data-stu-id="839c0-156">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="839c0-157">`RoamingSettings` 对象允许您存储和访问用户邮箱中存储的邮件外接程序的数据，以便从用于访问该邮箱的任何主机客户端应用程序中运行该外接程序时，该外接程序可以使用该数据。</span><span class="sxs-lookup"><span data-stu-id="839c0-157">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="839c0-158">类型:</span><span class="sxs-lookup"><span data-stu-id="839c0-158">Type:</span></span>

*   [<span data-ttu-id="839c0-159">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="839c0-159">RoamingSettings</span></span>](/javascript/api/outlook_1_3/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="839c0-160">要求</span><span class="sxs-lookup"><span data-stu-id="839c0-160">Requirements</span></span>

|<span data-ttu-id="839c0-161">要求</span><span class="sxs-lookup"><span data-stu-id="839c0-161">Requirement</span></span>| <span data-ttu-id="839c0-162">值</span><span class="sxs-lookup"><span data-stu-id="839c0-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="839c0-163">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="839c0-163">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="839c0-164">1.0</span><span class="sxs-lookup"><span data-stu-id="839c0-164">1.0</span></span>|
|[<span data-ttu-id="839c0-165">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="839c0-165">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="839c0-166">受限</span><span class="sxs-lookup"><span data-stu-id="839c0-166">Restricted</span></span>|
|[<span data-ttu-id="839c0-167">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="839c0-167">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="839c0-168">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="839c0-168">Compose or read</span></span>|