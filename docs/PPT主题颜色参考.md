# PowerPoint 主题颜色与格式化参考

> 本文档记录了 PowerPoint 主题颜色、字体和效果的 API 使用说明，供开发参考。

**相关代码**: `PPA/Formatting/FormatHelper.cs`

---

## 一、PPT 12 种主题颜色与 MsoThemeColorIndex 的对应关系

| 中文名称 (UI 显示) | 英文名称 (设计理念)             | 对应的 MsoThemeColorIndex 枚举值 | 说明                             |
| ------------------ | ------------------------------- | -------------------------------- | -------------------------------- |
| 文字/背景 - 深色 1 | Dark 1 (暗 1)                   | `msoThemeColorDark1`             | 通常用于主要文本或深色背景       |
| 文字/背景 - 浅色 1 | Light 1 (光 1)                  | `msoThemeColorLight1`            | 通常用于幻灯片背景或浅色文本     |
| 文字/背景 - 深色 2 | Dark 2 (暗 2)                   | `msoThemeColorDark2`             | 辅助深色，用于次要文本或背景     |
| 文字/背景 - 浅色 2 | Light 2 (光 2)                  | `msoThemeColorLight2`            | 辅助浅色，用于填充或高亮背景     |
| 着色 1             | Accent 1 (强调 1)               | `msoThemeColorAccent1`           | 主要强调色，通常是最突出的品牌色 |
| 着色 2             | Accent 2 (重音 2)               | `msoThemeColorAccent2`           | 次要强调色                       |
| 着色 3             | Accent 3 (重音 3)               | `msoThemeColorAccent3`           | 第三强调色                       |
| 着色 4             | Accent 4 (重音 4)               | `msoThemeColorAccent4`           | 第四强调色                       |
| 着色 5             | Accent 5 (重音 5)               | `msoThemeColorAccent5`           | 第五强调色                       |
| 着色 6             | Accent 6 (重音 6)               | `msoThemeColorAccent6`           | 第六强调色                       |
| 超链接             | Hyperlink (超链接)              | `msoThemeColorHyperlink`         | 用于未点击的超链接               |
| 已访问的超链接     | Followed Hyperlink (点击超链接) | `msoThemeColorFollowedHyperlink` | 用于已点击的超链接               |

**注意**: `msoThemeColorText1` 和 `msoThemeColorBackground1` 这两个枚举值也存在，它们在内部通常分别指向 `msoThemeColorDark1` 和 `msoThemeColorLight1`。

---

## 二、主题颜色变体 (Tint and Shade)

### 基本用法

```csharp
// 获取形状的填充颜色格式对象
var fillFormat = shape.Fill.ForeColor;

// 设置颜色为"着色 1"
fillFormat.ObjectThemeColor = MsoThemeColorIndex.msoThemeColorAccent1;

// 将其调亮 40% (Tint)
fillFormat.TintAndShade = 0.4f;

// 将其调暗 20% (Shade)
fillFormat.TintAndShade = -0.2f;
```

### 说明

-   **Tint (调亮)**: 值为正数，范围 0.0 - 1.0，数值越大越亮
-   **Shade (调暗)**: 值为负数，范围 -1.0 - 0.0，数值越小越暗

---

## 三、主题字体

### 获取主题字体方案

```csharp
// 获取当前幻灯片的主主题
var theme = slide.Design.Theme;

// 获取主题字体方案
var fontScheme = theme.ThemeFontScheme;

// 获取主要字体（用于标题）
var majorFont = fontScheme.MajorFont; // 这是一个 Font 对象，包含 .Name, .NameFarEast, .NameAscii 等属性
Profiler.LogMessage($"标题字体: {majorFont.Name}");

// 获取次要字体（用于正文）
var minorFont = fontScheme.MinorFont;
Profiler.LogMessage($"正文字体: {minorFont.Name}");
```

### 应用主题字体

```csharp
// 应用主题字体到文本框
shape.TextFrame.TextRange.Font.Name = fontScheme.MajorFont.Name;

// 应用次要字体到文本框
shape.TextFrame.TextRange.Font.Name = fontScheme.MinorFont.Name;
```

### 主题字体别名

PowerPoint 支持使用特殊的字体别名来引用主题字体，这些别名会自动适应主题变化：

```csharp
var tfont = textFrame.TextRange.Font;
tfont.Name = "+mn-lt";      // 拉丁字母使用主题的"次要字体"
tfont.NameFarEast = "+mn-ea"; // 东亚字符（如中文）使用主题的"次要字体"
```

### 字体别名对照表

| 代号     | 含义                        | 对应主题中的角色 | 常用场景           |
| -------- | --------------------------- | ---------------- | ------------------ |
| `+mj-lt` | Major Latin (主要拉丁语)    | 主要字体 (拉丁)  | 标题、页眉中的西文 |
| `+mj-ea` | Major East Asian (主要东亚) | 主要字体 (东亚)  | 标题、页眉中的中文 |
| `+mn-lt` | Minor Latin (次要拉丁语)    | 次要字体 (拉丁)  | 正文、备注中的西文 |
| `+mn-ea` | Minor East Asian (次要东亚) | 次要字体 (东亚)  | 正文、备注中的中文 |

---

## 四、主题效果

### 阴影效果

```csharp
// 应用一个预设的阴影效果，这个效果会与主题颜色协调
shape.Shadow.Type = MsoShadowType.msoShadow21;
```

### 柔化边缘效果

```csharp
// 应用一个预设的柔化边缘效果
shape.SoftEdge.Type = MsoSoftEdgeType.msoSoftEdgeType1;
```

---

## 五、整个主题颜色方案

### 获取和修改主题颜色方案

```csharp
// 获取当前幻灯片的设计主题
var theme = slide.Design.Theme;

// 获取主题颜色方案
var colorScheme = theme.ThemeColorScheme;

// 获取"着色 1"的 RGB 值
// 注意：GetColor 返回的是一个 MsoRGBType，可以转换为 int
int accent1Rgb = colorScheme.GetColor(MsoThemeColorIndex.msoThemeColorAccent1);

// 可以修改颜色方案（会影响到整个使用该主题的幻灯片）
colorScheme.Colors(MsoThemeColorIndex.msoThemeColorAccent1).RGB = RGB(255, 0, 0); // 将着色1改为红色
```

**警告**: 修改主题颜色方案会影响所有使用该主题的幻灯片，请谨慎操作。

---

## 六、功能总结

| 功能     | 对象/属性          | 示例用途                                    |
| -------- | ------------------ | ------------------------------------------- |
| 基础颜色 | `ObjectThemeColor` | 将形状或文本设置为 `msoThemeColorAccent1`   |
| 颜色变体 | `TintAndShade`     | 创建更亮或更暗的强调色                      |
| 主题字体 | `ThemeFontScheme`  | 获取或应用标题/正文字体                     |
| 主题效果 | `EffectFormat`     | 应用与主题协调的阴影、发光等                |
| 颜色方案 | `ThemeColorScheme` | 获取主题中任意颜色的 RGB 值，或修改整个方案 |

---

## 七、相关资源

-   [Microsoft Office Theme Colors Documentation](https://docs.microsoft.com/en-us/office/vba/api/office.msothemecolorindex)
-   [PowerPoint Object Model Reference](https://docs.microsoft.com/en-us/office/vba/api/powerpoint)
-   NetOffice API Documentation

---

**最后更新**: 2024 年 12 月  
**维护者**: PPA 开发团队
