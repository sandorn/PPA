# WPS 支持技术验证示例

本文档提供用于验证 WPS COM API 可用性的示例代码。

## 快速验证步骤

### 1. 创建测试项目

创建一个简单的控制台应用程序来测试 WPS COM 互操作。

### 2. 添加 COM 引用

在 Visual Studio 中：

1. 右键项目 → **添加引用**
2. 选择 **COM** 选项卡
3. 查找并添加 WPS 相关的类型库（如果可用）

### 3. 测试代码示例

```csharp
using System;
using System.Runtime.InteropServices;

namespace WPSComTest
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // 方法1: 通过 ProgID 创建 WPS 应用程序对象
                Type wpsType = Type.GetTypeFromProgID("WPS.Application");
                if (wpsType == null)
                {
                    Console.WriteLine("未找到 WPS 应用程序类型");
                    return;
                }

                dynamic wpsApp = Activator.CreateInstance(wpsType);
                Console.WriteLine($"WPS 应用程序名称: {wpsApp.Name}");
                Console.WriteLine($"WPS 版本: {wpsApp.Version}");

                // 方法2: 尝试获取活动演示文稿
                try
                {
                    dynamic pres = wpsApp.ActivePresentation;
                    Console.WriteLine($"活动演示文稿: {pres.Name}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"获取活动演示文稿失败: {ex.Message}");
                }

                // 方法3: 尝试创建新演示文稿
                try
                {
                    dynamic newPres = wpsApp.Presentations.Add();
                    Console.WriteLine("成功创建新演示文稿");
                    newPres.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"创建演示文稿失败: {ex.Message}");
                }

                Marshal.ReleaseComObject(wpsApp);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"错误: {ex.Message}");
                Console.WriteLine($"堆栈: {ex.StackTrace}");
            }

            Console.WriteLine("按任意键退出...");
            Console.ReadKey();
        }
    }
}
```

### 4. 检测应用程序类型

```csharp
public enum ApplicationType
{
    Unknown,
    PowerPoint,
    WPSPresentation
}

public static class ApplicationDetector
{
    public static ApplicationType DetectType(object app)
    {
        if (app == null) return ApplicationType.Unknown;

        try
        {
            dynamic dynApp = app;
            string name = dynApp.Name;
            string version = dynApp.Version;

            Console.WriteLine($"应用程序名称: {name}");
            Console.WriteLine($"版本: {version}");

            if (name.Contains("PowerPoint") || name.Contains("Microsoft"))
                return ApplicationType.PowerPoint;
            else if (name.Contains("WPS") || name.Contains("Kingsoft"))
                return ApplicationType.WPSPresentation;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"检测失败: {ex.Message}");
        }

        return ApplicationType.Unknown;
    }
}
```

### 5. 测试 WPS 表格操作

```csharp
public static void TestWPSTable(dynamic wpsApp)
{
    try
    {
        // 获取活动演示文稿
        dynamic pres = wpsApp.ActivePresentation;
        if (pres == null)
        {
            Console.WriteLine("没有活动的演示文稿");
            return;
        }

        // 获取第一张幻灯片
        dynamic slide = pres.Slides[1];

        // 尝试查找表格
        int shapeCount = slide.Shapes.Count;
        Console.WriteLine($"幻灯片中的形状数量: {shapeCount}");

        for (int i = 1; i <= shapeCount; i++)
        {
            dynamic shape = slide.Shapes[i];
            if (shape.HasTable != null && shape.HasTable == true)
            {
                Console.WriteLine($"找到表格: {shape.Name}");

                // 尝试访问表格属性
                dynamic table = shape.Table;
                int rows = table.Rows.Count;
                int cols = table.Columns.Count;
                Console.WriteLine($"表格大小: {rows} 行 x {cols} 列");

                // 尝试设置单元格值
                try
                {
                    dynamic cell = table.Cell(1, 1);
                    cell.Shape.TextFrame.TextRange.Text = "测试";
                    Console.WriteLine("成功设置单元格文本");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"设置单元格文本失败: {ex.Message}");
                }
            }
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"测试 WPS 表格操作失败: {ex.Message}");
        Console.WriteLine($"堆栈: {ex.StackTrace}");
    }
}
```

## 验证清单

- [ ] WPS 应用程序对象可以创建
- [ ] 可以获取应用程序名称和版本
- [ ] 可以访问活动演示文稿
- [ ] 可以创建新演示文稿
- [ ] 可以访问幻灯片
- [ ] 可以访问表格对象
- [ ] 可以访问文本对象
- [ ] 可以访问图表对象
- [ ] 可以设置表格样式
- [ ] 可以设置文本格式

## 注意事项

1. **ProgID 可能不同**: WPS 的 ProgID 可能与示例不同，需要查找 WPS 文档
2. **API 差异**: WPS 的 API 可能与 PowerPoint 不完全相同
3. **错误处理**: 需要完善的错误处理机制
4. **COM 对象释放**: 确保正确释放 COM 对象，避免内存泄漏

## 下一步

根据验证结果：

- 如果 WPS COM API 可用 → 继续实施方案
- 如果 WPS COM API 不可用或兼容性差 → 考虑备选方案
