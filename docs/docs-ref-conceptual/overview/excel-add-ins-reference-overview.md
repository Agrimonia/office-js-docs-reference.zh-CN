# <a name="excel-javascript-api-overview"></a>Excel 的 JavaScript API 概述

Excel 的 JavaScript API 可用于生成加载项的 Excel 2016 或更高版本。 以下列表显示在 API 中可用的高级 Excel 对象。 每个对象页上链接包含的属性、 事件和方法的对象上的说明。 如需了解详细信息，请从菜单中浏览相应链接。

为了方便起见，下面列出了一些核心 Excel 对象： 

- [工作簿](/javascript/api/excel/excel.workbook)：包含相关 workbook 对象的顶级对象，例如 worksheet、table、range 等。它还可以用于列出相关的参考。

- [Worksheet](/javascript/api/excel/excel.worksheet)：表示工作簿中的工作表。 
    - [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)：工作簿中 **Worksheet** 对象的集合。

- [Range](/javascript/api/excel/excel.range)：表示某一单元格、某一行、某一列、某一单元格选定区域（其中包含一个或多个相邻单元格块）。

- [Table](/javascript/api/excel/excel.table)：表示有组织的单元格的集合，设计用于简化数据管理。
    - [TableCollection](/javascript/api/excel/excel.tablecollection)：工作簿或工作表中的表的集合。
    - [TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)：表中所有列的集合。
    - [TableRowCollection](/javascript/api/excel/excel.tablerowcollection)：表中所有行的集合。

- [Chart](/javascript/api/excel/excel.chart)：表示工作表中的 chart 对象，它是基础数据的可视表示形式。
    - [ChartCollection](/javascript/api/excel/excel.chartcollection)：工作表中的图表的集合。

- [TableSort](/javascript/api/excel/excel.tablesort)：表示管理 **Table** 对象上的排序操作的对象。

- [RangeSort](/javascript/api/excel/excel.rangesort)：表示管理 **Range** 对象上的排序操作的对象。

- [Filter](/javascript/api/excel/excel.filter)：表示管理表格列筛选的对象。

- [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)：表示对 **Worksheet** 对象的保护。

- [NamedItem](/javascript/api/excel/excel.nameditem)：表示单元格区域或值的定义名称。 
    - [NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)：工作簿中 **NamedItem** 对象的集合。

- [Binding](/javascript/api/excel/excel.binding)：表示对工作簿的某一部分的绑定的抽象类。
    - [BindingCollection](/javascript/api/excel/excel.bindingcollection)：工作簿中 **Binding** 对象的集合。

## <a name="excel-javascript-api-open-specifications"></a>Excel 的 JavaScript API 开放性规范

在我们设计和开发针对 Excel 加载项的新 Api，我们将使其可用于您的反馈我们[Open API 规范](../openspec.md)的页面。 了解新增功能中的管道 Excel JavaScript Api，并在我们的设计规范上提供的输入。

## <a name="excel-javascript-api-reference"></a>Excel JavaScript API 参考

有关 Excel JavaScript API 的详细信息，请参阅[Excel 的 JavaScript API 参考文档](/javascript/api/excel)。

## <a name="see-also"></a>另请参阅

- [Excel 加载项概述](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-overview)
- [Office 加载项平台概述](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
- [在 GitHub Excel 加载项示例](https://github.com/OfficeDev?utf8=%E2%9C%93&q=Excel)
