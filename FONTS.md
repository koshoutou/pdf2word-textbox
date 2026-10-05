# 字体说明

> 本工具 **不含任何字体文件**。为获得最佳 1:1 复刻效果,请按本文档安装字体。

## 为什么字体很重要

PDF 转 Word 时,若目标系统未安装原 PDF 使用的字体(如仿宋、黑体),Word/WPS/LibreOffice
会使用替代字体渲染,导致:

- 文字宽度变化(可能折行或错位)
- 字形外观不一致(宋体被替换成黑体等)
- 标题/正文层级感丢失

本工具虽已通过"文本宽度估算 + 禁止折行"缓解折行问题,但**字体本身的字形差异无法通过
代码弥补**,必须安装原字体才能实现真正的 1:1 视觉复刻。

---

## 推荐字体安装方案

### 方案一:Windows 常用中文字体(推荐)

以下两个 GitHub 仓库收集了 Windows 系统常用的中文字体(仿宋、黑体、宋体、楷体、
华文系列、方正系列),可直接克隆使用:

1. **zhyounger/FontsFromWindows**
   - 仓库:https://github.com/zhyounger/FontsFromWindows/tree/master/fonts
   - 含:仿宋、宋体、黑体、楷体、华文宋体/仿宋/楷体/中宋/行楷、方正小标宋等 13 款

2. **DoveOutland/Common-Chinese-office-fonts-font-library-**
   - 仓库:https://github.com/DoveOutland/Common-Chinese-office-fonts-font-library-
   - 含:仿宋、宋体、黑体、楷体、Times New Roman、方正系列(仿宋/楷体/黑体/大小标宋)等 20 款

**Linux 安装方法**:

```bash
# 克隆字体仓库
git clone https://github.com/zhyounger/FontsFromWindows.git
git clone https://github.com/DoveOutland/Common-Chinese-office-fonts-font-library-.git

# 复制到用户字体目录
mkdir -p ~/.local/share/fonts/cn-office
cp FontsFromWindows/fonts/*.ttf ~/.local/share/fonts/cn-office/
cp FontsFromWindows/fonts/*.ttc ~/.local/share/fonts/cn-office/
cp "Common-Chinese-office-fonts-font-library-/"*.ttf ~/.local/share/fonts/cn-office/
cp "Common-Chinese-office-fonts-font-library-/"*.ttc ~/.local/share/fonts/cn-office/
cp -r "Common-Chinese-office-fonts-font-library-/Times New Roman" ~/.local/share/fonts/cn-office/

# 刷新字体缓存
fc-cache -f
```

**macOS 安装方法**:双击 `.ttf`/`.ttc` 文件,在"字体册"中点击"安装字体"。

**Windows**:字体已在 `C:\Windows\Fonts\` 自带,无需额外安装。

### 方案二:晋城市财政局字体包(官方公文场景)

适用于政府公文、招投标文件等正式场景,字体包来自晋城市财政局官网:

- **下载地址**:http://czj.jcgov.gov.cn/ggfw/xzzx/202506/P020250624600205358001.zip
- **来源**:晋城市财政局(http://czj.jcgov.gov.cn/)
- **说明**:该字体包为官方提供的常用办公字体合集,适合公文排版

下载后解压,将 `.ttf`/`.ttc` 文件复制到系统字体目录即可。

---

## 工具的字体处理机制

本工具的字体处理分为三步:

### 1. 字体名识别(`fonts.py`)

PDF 字体名通常带子集前缀和样式后缀,如 `FAAAAH+FangSong,Bold`:

- 去除子集前缀(`FAAAAH+` → `FangSong,Bold`)
- 去除样式后缀(`FangSong,Bold` → `FangSong`)
- 映射到标准中文字体名(`FangSong` → `仿宋`)

支持识别的字体族:

| 类别 | 字体 |
|------|------|
| 宋体类 | SimSun, NSimSun, STSong, STSong-Light |
| 黑体类 | SimHei, STHeiti, Microsoft YaHei |
| 楷体类 | KaiTi, STKaiti, SimKai |
| 仿宋类 | FangSong, STFangsong, SimFang |
| 华文系列 | STXihei, STZhongsong, STKaiti, STCaiyun |
| 方正系列 | FZFangSong, FZKai, FZHei, FZXiaoBiaoSong |
| 其他 | 隶书, 幼圆 |

### 2. 字号整数化(`normalize_size`)

PDF 字号由变换矩阵计算,可能产生浮点误差(如 `13.99` 本应是 `14.0`)。
本工具将接近整数的字号规整为整数(容差 0.1pt),其余保留 1 位小数:

```
13.99 → 14.0
14.01 → 14.0
10.5  → 10.5
9.95  → 10.0
```

### 3. 系统字体可用性检测(`is_font_available`)

转换时,工具会用 `fc-list`(Linux/macOS)检测目标字体是否在系统已安装:

- **已安装**:直接使用原字体名(最佳渲染)
- **未安装**:回退到系统已有的中文衬线/无衬线字体(避免完全无渲染)

检测带缓存(`lru_cache`),同一字体名只检测一次,不影响性能。

---

## 字体版权与免责声明

> **重要**:字体文件受版权法保护。本工具 **不分发、不内嵌任何字体文件**,
> 仅提供字体名识别与映射功能。用户需自行确保拥有所用字体的合法授权。

### 各字体版权说明

| 字体 | 版权方 | 授权情况 |
|------|--------|----------|
| 仿宋 / 宋体 / 黑体 / 楷体 | 微软(Windows 自带) | 随 Windows 授权 |
| 华文系列 | 字王(zhongguowang.com) | 商业字体,需购买授权 |
| 方正系列 | 北大方正集团 | 商业字体,需购买授权(个人非商用免费) |
| Times New Roman | 微软 | 随 Windows 授权 |
| Noto Sans/Serif CJK | Google | 开源(SIL OFL) |
| 思源系列 | Adobe + Google | 开源(SIL OFL) |

### 免责声明

1. 本工具不提供、不分发任何字体文件,字体版权归各自所有者。
2. 用户通过本工具生成的 docx 文档中,字体名仅作为样式引用,不嵌入字体数据。
3. 用户应自行确保在所在地区使用相关字体已获得合法授权。
4. 上述字体仓库链接与晋城市财政局字体包链接仅为方便用户获取字体,本工具不对
   这些外部资源的可用性、合法性、安全性负责。
5. 商业使用方正系列、华文系列等付费字体前,请向版权方购买授权。
6. 政府公文场景建议使用官方提供的字体包(如晋城市财政局字体包),以确保合规。

---

## 验证字体是否安装成功

```bash
# Linux/macOS
fc-list | grep -iE "fangsong|simhei|simsun|kaiti|仿宋|黑体|宋体|楷体"
```

应看到类似输出:

```
/home/.../仿宋_常规.ttf: FangSong,仿宋:style=Regular
/home/.../黑体_常规.ttf: SimHei,黑体:style=Regular
/home/.../宋体_常规.ttc: SimSun,宋体:style=Regular
/home/.../楷体_常规.ttf: KaiTi,楷体:style=Regular
```

也可用本工具自带的检测函数:

```python
from pdf2word_textbox.fonts import is_font_available, list_system_fonts
print("仿宋:", is_font_available("仿宋"))
print("黑体:", is_font_available("黑体"))
print("系统字体总数:", len(list_system_fonts()))
```
