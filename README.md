# 简阅 · JianYue

极简的本地文档阅读器：Android（WebView 壳）+ Windows（Electron），
支持 **PDF / Word / Excel / PPT / Markdown**，含旧版 `.doc / .xls / .ppt` 的自行解析。

全部在本机解析，不联网、不上传。

## 功能

- 自动扫描本机文档，按类型筛选、文件名/目录搜索
- PDF：矢量分块渲染，缩放不糊
- Word / Excel / PPT（OOXML）：版式还原
- 旧版 `.doc`（Word 97）、`.xls`（BIFF8）、`.ppt`（PPT97）：自研解析器
  - `.ppt` 按形状坐标还原文字/图片/填充，EMF 图元与渐变不支持
- Markdown：标题、代码块、表格、引用、列表
- 书库（合集）、多选、删除、拖拽（Windows 可拖到资源管理器）
- 分屏阅读（左右/上下）、浅色/深色磨砂主题

## 目录

```
android/   前端资产（WebView 内运行）+ 打包脚本
windows/   Electron 桌面版（复用同一套前端）
```

## 构建

### Android

需要：`aapt2`、`apksigner`、`zipalign`、Python 3，以及一个可用的 base APK（作为壳）。

```bash
export JY_KS=path/to/your.jks
export JY_KS_PASS=****** JY_KEY_PASS=******
bash android/build.sh
```

> 仓库不含签名密钥，请自备；改版本号在 `android/tools/patch_manifest.py` 调用处。

### Windows

```bash
cd windows
npm install
npm start        # 开发运行
npm run dist     # 打包 exe
```

## 已知限制

- 旧版 `.ppt`：EMF/WMF 矢量图元、渐变填充、动画无法还原
- Android 端 `.ppt` 渲染为自研近似实现，达不到 PDF 级 1:1

## 许可

本项目代码 MIT；第三方组件见 [NOTICE.md](NOTICE.md)。
