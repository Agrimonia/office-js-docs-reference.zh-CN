# <a name="control-element"></a>控件（Control）元素

定义执行操作或启动任务窗格（Task pane）的 JavaScript 函数。**Control** 元素可以是 `button`，也可以是 `menu`。[Group](group.md) 元素中至少需包括一个 **Control**。

## <a name="attributes"></a>属性

|  属性  |  必需  |  说明  |
|:-----|:-----|:-----|
|**xsi:type**|是|正在定义的控件类型。可以是 `Button`、`Menu` 或 `MobileButton`。 |
|**id**|否|控件元素的 ID。最多可包含 125 个字符。|

> [!NOTE]
> **xsi:type** 的 `MobileButton` 类型是在 VersionOverrides schema 1.1 中的新定义。 它仅适用于 [MobileFormFactor](mobileformfactor.md) 元素中包含的**控件**元素。

## <a name="button-control"></a>Button 控件

当用户选择某个按钮时，将执行一个操作。它可以执行函数或显示任务窗格。每个按钮控件必须具有对清单唯一的 `id`。 

### <a name="child-elements"></a>子元素
|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  **Label**     | 是 |  按钮文本。**resid** 属性必须设置为 **ShortStrings** 元素（位于 [Resources](resources.md) 元素）中 **String** 元素的 **id** 属性的值。        |
|  **ToolTip**  |否|按钮的提示框。**resid** 属性必须设置为 **String** 元素的 **id** 属性的值。**String** 元素是 **LongStrings** 元素的子元素，而 LongStrings 元素是 [Resources](resources.md) 元素的子元素。|        
|  [Supertip](supertip.md)  | 是 |  按钮的 supertip。    |
|  [Icon](icon.md)      | 是 |  按钮的图像。         |
|  [Action](action.md)    | 是 |  指定要执行的操作。  |

### <a name="executefunction-button-example"></a>ExecuteFunction Button 控件示例

```xml
<Control xsi:type="Button" id="msgReadFunctionButton">
  <Label resid="funcReadButtonLabel" />
  <Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="blue-icon-16" />
    <bt:Image size="32" resid="blue-icon-32" />
    <bt:Image size="80" resid="blue-icon-80" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>getSubject</FunctionName>
  </Action>
</Control>
```

### <a name="showtaskpane-button-example"></a>ShowTaskpane Button 控件示例

```xml
<Control xsi:type="Button" id="msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Supertip>
    <Title resid="paneReadSuperTipTitle" />
    <Description resid="paneReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="green-icon-16" />
    <bt:Image size="32" resid="green-icon-32" />
    <bt:Image size="80" resid="green-icon-80" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
  </Action>
</Control>
```

## <a name="menu-dropdown-button-controls"></a> Menu (下拉菜单) 控件

`menu` 定义选项的静态列表。每个 `menu` 项可以执行函数或显示任务窗格。不支持 `Submenus`。

使用 **PrimaryCommandSurface** 或 **ContextMenu** [扩展点](extensionpoint.md)时，`menu control` 定义了：

- 根级别的 `menu` 项。（A root-level menu item.）

- `submenu` 项列表。

与 **PrimaryCommandSurface** 一起使用时，root-menu 项将显示为功能区上的按钮。选择该按钮后，`submenu` 将显示为下拉列表。  
与 **ContextMenu** 一起使用时，具有 `submenu` 的 `menu` 项将被插入到 **ContextMenu** 上。在这两种情况下，每个 `submenu` 项可以执行 JavaScript 函数，也可显示任务窗格。这时仅支持一级子菜单。

下面的示例演示如何定义具有两个 `submenu` 项的 `menu` 项。第一个 `submenu` 项显示任务窗格，而第二个 `submenu` 项运行 JavaScript 函数。

```xml
<Control xsi:type="Menu" id="TestMenu2">
  <Label resid="residLabel3" />
  <Tooltip resid="residToolTip" />
  <Supertip>
    <Title resid="residLabel" />
    <Description resid="residToolTip" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="icon1_32x32" />
    <bt:Image size="32" resid="icon1_32x32" />
    <bt:Image size="80" resid="icon1_32x32" />
  </Icon>
  <Items>
    <Item id="showGallery2">
      <Label resid="residLabel3"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Action xsi:type="ShowTaskpane">
        <TaskpaneId>MyTaskPaneID1</TaskpaneId>
        <SourceLocation resid="residUnitConverterUrl" />
      </Action>
    </Item>
    <Item id="showGallery3">
      <Label resid="residLabel5"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon4_32x32" />
        <bt:Image size="32" resid="icon4_32x32" />
        <bt:Image size="80" resid="icon4_32x32" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getButton</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>

```

### <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  **Label**     | 是 |  按钮文本。**resid** 属性必须设置为 **ShortStrings** 元素（位于 [Resources](resources.md) 元素）中 **String** 元素的 **id** 属性的值。      |
|  **ToolTip**  |否|按钮的提示框。**resid** 属性必须设置为 **String** 元素的 **id** 属性的值。**String** 元素是 **LongStrings** 元素的子元素，而 LongStrings 元素是 [Resources](resources.md) 元素的子元素。|        
|  [Supertip](supertip.md)  | 是 |  此按钮的 supertip。    |
|  [Icon](icon.md)      | 是 |  按钮的图像。         |
|  **Items**     | 是 |  `menu`中显示的按钮的集合。包含每个子`menu`项的 **Item** 元素。每个 **Item** 元素均包含 [按钮控件](#button-control) 的子元素。|

### <a name="menu-control-examples"></a>`menu` 控件示例

```xml
<Control xsi:type="Menu" id="TestMenu2">
  <Label resid="residLabel3" />
  <Tooltip resid="residToolTip" />
  <Supertip>
    <Title resid="residLabel" />
    <Description resid="residToolTip" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="icon1_32x32" />
    <bt:Image size="32" resid="icon1_32x32" />
    <bt:Image size="80" resid="icon1_32x32" />
  </Icon>
  <Items>
    <Item id="showGallery2">
      <Label resid="residLabel3"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Action xsi:type="ShowTaskpane">
        <TaskpaneId>MyTaskPaneID1</TaskpaneId>
        <SourceLocation resid="residUnitConverterUrl" />
      </Action>
    </Item>
    <Item id="showGallery3">
      <Label resid="residLabel5"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon4_32x32" />
        <bt:Image size="32" resid="icon4_32x32" />
        <bt:Image size="80" resid="icon4_32x32" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getButton</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>

```

```xml
<Control xsi:type="Menu" id="msgReadMenuButton">
  <Label resid="menuReadButtonLabel" />
  <Supertip>
    <Title resid="menuReadSuperTipTitle" />
    <Description resid="menuReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="red-icon-16" />
    <bt:Image size="32" resid="red-icon-32" />
    <bt:Image size="80" resid="red-icon-80" />
  </Icon>
  <Items>
    <Item id="msgReadMenuItem1">
      <Label resid="menuItem1ReadLabel" />
      <Supertip>
        <Title resid="menuItem1ReadLabel" />
        <Description resid="menuItem1ReadTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="red-icon-16" />
        <bt:Image size="32" resid="red-icon-32" />
        <bt:Image size="80" resid="red-icon-80" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getItemClass</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>
```

## <a name="mobilebutton-control"></a>MobileButton 控件

当用户选择某个 `MobileButton` 时，将执行单个操作。它可以执行函数或显示任务窗格。 每个移动按钮控件必须具有对清单唯一的 `id`。

在 VersionOverrides 架构 1.1 中定义了 **xsi:type** 的 `MobileButton` 值。包含 [VersionOverrides](versionoverrides.md) 元素的 `VersionOverridesV1_1` 属性值必须为 `xsi:type`。

### <a name="child-elements"></a>子元素
|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  **Label**     | 是 |  按钮文本。**resid** 属性必须设置为 **ShortStrings** 元素（位于 [Resources](resources.md) 元素）中 **String** 元素的 **id** 属性的值。        |
|  [Icon](icon.md)      | 是 |  按钮的图像。         |
|  [Action](action.md)    | 是 |  指定要执行的操作。  |

### <a name="executefunction-mobile-button-example"></a>ExecuteFunction MobileButton 控件示例

```xml
<Control xsi:type="MobileButton" id="msgReadFunctionButton">
  <Label resid="funcReadButtonLabel" />
  <Icon>
    <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
    <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
    <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
    <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
    <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
    <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
    <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
    <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
    <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>getSubject</FunctionName>
  </Action>
</Control>
```

### <a name="showtaskpane-mobile-button-example"></a>ShowTaskpane MobileButton 控件示例

```xml
<Control xsi:type="MobileButton" id="msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Icon>
    <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
    <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
    <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
    <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
    <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
    <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
    <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
    <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
    <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
  </Action>
</Control>
```