# Application 变量命名规范

## 📋 规范概述

本文档定义了 PPA 项目中 Application 相关变量的统一命名规范，确保代码一致性和可读性。

---

## 1. 命名规范

### 1.1 变量类型与命名对应表

| 变量类型                          | 命名规范              | 示例                                   | 说明               |
| --------------------------------- | --------------------- | -------------------------------------- | ------------------ |
| **IApplication** (抽象接口)       | `abstractApp` / `app` | `IApplication abstractApp`             | 方法参数或局部变量 |
| **IApplication** (私有字段)       | `_abstractApp`        | `private IApplication _abstractApp;`   | 类私有字段         |
| **NETOP.Application** (NetOffice) | `netApp` / `app`      | `NETOP.Application netApp`             | 方法参数或局部变量 |
| **NETOP.Application** (私有字段)  | `_netApp`             | `private NETOP.Application _netApp;`   | 类私有字段         |
| **MSOP.Application** (原生 COM)   | `nativeApp`           | `MSOP.Application nativeApp`           | 方法参数或局部变量 |
| **MSOP.Application** (私有字段)   | `_nativeApp`          | `private MSOP.Application _nativeApp;` | 类私有字段         |

### 1.2 特殊情况

| 场景                   | 命名                                   | 说明                             |
| ---------------------- | -------------------------------------- | -------------------------------- |
| **缓存字段**           | `_cachedAbstractApp` / `_cachedNetApp` | 带缓存的字段                     |
| **方法参数（通用）**   | `app`                                  | 如果上下文明确，可以使用简短名称 |
| **ThisAddIn 中的字段** | `NetApp` / `NativeApp`                 | 公共属性，使用 PascalCase        |

---

## 2. 命名规则详解

### 2.1 抽象接口 (IApplication)

```csharp
// ✅ 推荐：方法参数
public void ExecuteAlignment(IApplication abstractApp, ...)
{
    // ...
}

// ✅ 推荐：局部变量
IApplication abstractApp = GetAbstractApplication();

// ✅ 推荐：私有字段
private IApplication _abstractApp;

// ✅ 如果上下文明确，可以使用简短名称
public void Process(IApplication app) // 在业务方法中，app 通常指抽象接口
{
    // ...
}
```

### 2.2 NetOffice 对象 (NETOP.Application)

```csharp
// ✅ 推荐：方法参数
public void FormatShapes(NETOP.Application netApp)
{
    // ...
}

// ✅ 推荐：局部变量
NETOP.Application netApp = ApplicationHelper.GetNetOfficeApplication(abstractApp);

// ✅ 推荐：私有字段
private NETOP.Application _netApp;

// ⚠️ 如果上下文明确，可以使用简短名称（但建议使用 netApp）
public void Process(NETOP.Application app) // 在底层方法中，如果明确是 NetOffice
{
    // ...
}
```

### 2.3 原生 COM 对象 (MSOP.Application)

```csharp
// ✅ 推荐：方法参数（必须使用 nativeApp）
public static void CropShapesToSlide(MSOP.Application nativeApp)
{
    // ...
}

// ✅ 推荐：局部变量
MSOP.Application nativeApp = ApplicationHelper.GetNativeComApplication();

// ✅ 推荐：私有字段
private MSOP.Application _nativeApp;
```

### 2.4 缓存字段

```csharp
// ✅ 推荐：缓存抽象接口
private IApplication _cachedAbstractApp;
private NETOP.Application _cachedNetApp;

// ✅ 推荐：在 AlignHelper 中的缓存
private IApplication _cachedApp; // 缓存抽象接口
private NETOP.Application _cachedNativeApp; // 缓存 NetOffice 对象（注意：这里命名有歧义，应该改为 _cachedNetApp）
```

---

## 3. 重构建议

### 3.1 需要统一的地方

#### CustomRibbon.cs

- `_app` → `_netApp` (类型：NETOP.Application)
- `_abstractApp` → 保持不变（已符合规范）

#### AlignHelper.cs

- `_cachedApp` → `_cachedAbstractApp` (类型：IApplication)
- `_cachedNativeApp` → `_cachedNetApp` (类型：NETOP.Application，注意：不是 MSOP.Application)

#### 方法参数

- 统一使用 `abstractApp` 表示 IApplication
- 统一使用 `netApp` 表示 NETOP.Application
- 统一使用 `nativeApp` 表示 MSOP.Application

---

## 4. 命名优先级

### 4.1 明确性优先

```csharp
// ✅ 好：明确表示类型
public void Method(IApplication abstractApp, NETOP.Application netApp)
{
    MSOP.Application nativeApp = ApplicationHelper.GetNativeComApplication();
}

// ❌ 不好：不够明确
public void Method(IApplication app1, NETOP.Application app2)
{
    MSOP.Application app3 = ApplicationHelper.GetNativeComApplication();
}
```

### 4.2 上下文明确时可以使用简短名称

```csharp
// ✅ 好：在业务方法中，app 通常指抽象接口
public void ExecuteAlignment(IApplication app, AlignmentType alignment)
{
    // 上下文明确，app 是抽象接口
}

// ✅ 好：在底层方法中，如果方法名已经说明类型
public void ProcessNetOfficeApp(NETOP.Application app)
{
    // 方法名已说明类型，可以使用简短名称
}
```

---

## 5. 示例对比

### 5.1 重构前

```csharp
// CustomRibbon.cs
private NETOP.Application _app; // ❌ 不够明确
private IApplication _abstractApp; // ✅ 符合规范

// AlignHelper.cs
private IApplication _cachedApp; // ⚠️ 可以更明确
private NETOP.Application _cachedNativeApp; // ❌ 命名有歧义（不是 MSOP.Application）

// 方法参数
public void Method(IApplication app, NETOP.Application native) // ⚠️ native 不够明确
{
    MSOP.Application comApp = ...; // ⚠️ comApp 不够明确
}
```

### 5.2 重构后

```csharp
// CustomRibbon.cs
private NETOP.Application _netApp; // ✅ 明确表示 NetOffice
private IApplication _abstractApp; // ✅ 保持不变

// AlignHelper.cs
private IApplication _cachedAbstractApp; // ✅ 明确表示抽象接口
private NETOP.Application _cachedNetApp; // ✅ 明确表示 NetOffice

// 方法参数
public void Method(IApplication abstractApp, NETOP.Application netApp) // ✅ 明确
{
    MSOP.Application nativeApp = ...; // ✅ 明确表示原生 COM
}
```

---

## 6. 实施检查清单

- [ ] CustomRibbon.cs: `_app` → `_netApp`
- [ ] AlignHelper.cs: `_cachedApp` → `_cachedAbstractApp`
- [ ] AlignHelper.cs: `_cachedNativeApp` → `_cachedNetApp`
- [ ] 所有方法参数统一命名
- [ ] 所有局部变量统一命名
- [ ] 更新相关注释和文档

---

**文档版本**：1.0  
**最后更新**：2024 年 11 月 15 日
