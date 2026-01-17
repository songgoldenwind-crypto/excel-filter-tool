# 🚀 最简单的方法：在Windows电脑上打包

## 为什么用这个方法？

- ✅ 最简单：只需双击一个文件
- ✅ 最可靠：不依赖GitHub网络
- ✅ 最快速：3-5分钟完成
- ✅ 一次搞定：打包后可直接使用

---

## 📝 操作流程

### 准备工作（Mac上）

您已经有了这些文件，现在把它们都拷贝到U盘：

```
✅ excel_filter_tool.py
✅ requirements.txt
✅ install_and_build.bat  ← 这个是关键！
```

### 执行（Windows电脑上）

1. **插入U盘**，复制文件到桌面（任意文件夹都行）

2. **双击运行** `install_and_build.bat`

3. 按提示操作：
   - 如果没有Python，会自动打开下载页面
   - 下载Python安装时**必须勾选** "Add Python to PATH"
   - 安装完Python后，**重新双击** `install_and_build.bat`
   - 等待自动完成

4. **完成！**
   - exe文件在 `dist` 文件夹
   - 会自动打开文件夹显示结果

---

## 🎬 完整流程图

```
在Mac上：
├── 把文件拷到U盘
│   ├── excel_filter_tool.py
│   ├── requirements.txt
│   └── install_and_build.bat
└── 带U盘到Windows电脑

在Windows上：
├── 双击 install_and_build.bat
├── 首次运行 → 自动打开Python下载页面
├── 安装Python（记得勾选"Add to PATH"）
├── 重新双击 install_and_build.bat
├── 自动安装依赖
├── 自动打包
└── 得到 Excel数据筛选工具.exe ✓
```

---

## ❓ 常见问题

### Q: 我没有Windows电脑怎么办？
**A:** 可以借朋友的，或者去网吧/公司，5分钟就能搞定。

### Q: Python安装失败？
**A:** 确保下载的是 Windows installer (64-bit)，安装时勾选 "Add Python to PATH"

### Q: 需要多长时间？
**A:** 首次（需要装Python）：10-15分钟
    已有Python：3-5分钟

### Q: 可以打包多次吗？
**A:** 可以！修改代码后重新双击即可。

---

## 📁 最终文件结构

打包完成后：
```
当前文件夹/
├── dist/
│   └── Excel数据筛选工具.exe  ← 这就是你要的文件！
├── build/
└── 其他文件...
```

把 `Excel数据筛选工具.exe` 拷贝出来，就可以在**任何Windows电脑**上使用了！

---

## 🎯 特别说明

这个脚本会自动：
1. ✓ 检测Python是否安装
2. ✓ 未安装时打开下载页面
3. ✓ 自动安装所需依赖
4. ✓ 自动打包生成exe
5. ✓ 自动打开结果文件夹

全程自动化，无需手动输入命令！

---

**现在就试试吧！把文件拷到Windows电脑，双击 `install_and_build.bat`！** 🚀
