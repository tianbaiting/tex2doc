# tex2doc
trans latex to docx

this python code allows us to trans tex to docx , using pandoc.

And can transever the pdf img in tex, in order to put the img in docx.


# tex_to_word.py
# 需要pandoc ,ImageMagick，poppler-utils
# 运行例子 python3 tex2word.py input.tex -o out.docx
# tbt 2025-07-18


## 功能特性

- ✅ **LaTeX到Word转换**：基于Pandoc的高质量文档转换
- ✅ **PDF图片自动转换**：自动检测并转换LaTeX文档中的PDF图片为PNG格式
- ✅ **参考文献支持**：支持.bib文件的引用处理
- ✅ **版本兼容性**：自动检测Pandoc版本并适配相应功能
- ✅ **高质量输出**：300 DPI分辨率确保图片质量
- ✅ **智能路径处理**：自动处理相对路径和绝对路径
- ✅ **临时文件管理**：自动清理转换过程中的临时文件

## 依赖要求

### 必需依赖

1. **Python 3.6+**
2. **Pandoc**：文档转换核心工具
   ```bash
   # Ubuntu/Debian
   sudo apt-get install pandoc
   
   # Windows (使用Chocolatey)
   choco install pandoc
   
   # macOS (使用Homebrew)
   brew install pandoc
   ```

3. **pypandoc**：Python的Pandoc接口
   ```bash
   pip install pypandoc
   ```

### 可选依赖（用于PDF图片转换）

选择以下任一工具：

1. **ImageMagick**（推荐）：
   ```bash
   # Ubuntu/Debian
   sudo apt-get install imagemagick
   
   # Windows (Chocolatey)
   choco install imagemagick
   
   # macOS (Homebrew)
   brew install imagemagick
   ```

2. **poppler-utils**：
   ```bash
   # Ubuntu/Debian
   sudo apt-get install poppler-utils
   
   # Windows: 下载预编译二进制文件
   # macOS (Homebrew)
   brew install poppler
   ```

## 安装

1. 克隆或下载此脚本
2. 确保所有依赖已安装
3. 使脚本可执行（Linux/macOS）：
   ```bash
   chmod +x tex2word.py
   ```

## 使用方法

### 基本用法

```bash
# 最简单的用法
python3 tex2word.py input.tex

# 指定输出文件名
python3 tex2word.py input.tex -o output.docx

# 包含参考文献
python3 tex2word.py input.tex -o output.docx -b references.bib

# 禁用PDF图片转换
python3 tex2word.py input.tex -o output.docx --no-pdf-convert

# 指定 word 模板，添加参考文献格式
python3 tex2word.py input.tex -o output.docx -b references.bib --r template.docx --csl style.csl
```

### 命令行参数

| 参数 | 说明 | 必需 |
|------|------|------|
| `input_tex` | 输入的LaTeX文件路径 | ✅ |
| `-o, --output` | 输出的Word文件路径 | ❌ |
| `-b, --bibliography` | 参考文献.bib文件路径 | ❌ |
| `--no-pdf-convert` | 禁用PDF图片转换功能 | ❌ |

### 使用示例

```bash
# 示例1：转换简单的LaTeX文档
python3 tex2word.py report.tex -o report.docx

# 示例2：带参考文献的学术论文
python3 tex2word.py thesis.tex -o thesis.docx -b thesis.bib

# 示例3：保持PDF图片格式
python3 tex2word.py document.tex --no-pdf-convert

# 示例4：使用相对路径
python3 tex2word.py ../papers/paper.tex -o ../output/paper.docx
```

## PDF图片转换功能

### 工作原理

脚本会自动：
1. 扫描LaTeX文件中的`\includegraphics{}`命令
2. 识别PDF格式的图片文件
3. 使用ImageMagick或poppler-utils将PDF转换为PNG
4. 更新LaTeX文件中的图片引用
5. 继续进行文档转换

### 支持的图片引用格式

```latex
% 基本格式
\includegraphics{image.pdf}

% 带选项的格式
\includegraphics[width=0.5\textwidth]{figure.pdf}
\includegraphics[scale=0.8]{diagram.pdf}

% 相对路径
\includegraphics{images/chart.pdf}

% 绝对路径
\includegraphics{/home/user/figures/plot.pdf}
```

### 转换质量设置

- **分辨率**：300 DPI（适合打印质量）
- **格式**：PNG（支持透明背景）
- **质量**：90%（平衡文件大小和图片质量）

## 常见问题

### Q: 出现"Permission denied"错误怎么办？

**A:** 这通常是因为输出文件被其他程序占用（如Word正在打开该文件）：
1. 关闭可能打开输出文件的程序
2. 选择不同的输出文件名
3. 检查文件夹写入权限

### Q: PDF图片转换失败怎么办？

**A:** 请确保安装了以下工具之一：
- ImageMagick：`which convert` 应该有输出
- poppler-utils：`which pdftoppm` 应该有输出

如果都没有安装，可以使用 `--no-pdf-convert` 参数跳过转换。

### Q: 数学公式显示不正确怎么办？

**A:** 确保安装了完整的LaTeX发行版：
- Ubuntu/Debian：`sudo apt-get install texlive-full`
- Windows：安装MiKTeX或TeX Live
- macOS：`brew install --cask mactex`

### Q: 中文字符显示乱码怎么办？

**A:** 
1. 确保LaTeX文件使用UTF-8编码
2. 在LaTeX文档中使用适当的中文包（如ctex）
3. 检查系统是否安装了中文字体

### Q: 转换很慢怎么办？

**A:** 转换速度取决于：
- 文档大小和复杂性
- 图片数量（特别是PDF图片）
- 系统性能

可以通过以下方式优化：
- 使用 `--no-pdf-convert` 跳过图片转换
- 预先将PDF图片转换为PNG
- 简化复杂的LaTeX结构

## 技术细节

### 支持的Pandoc版本

- **Pandoc 2.11+**：支持 `--citeproc` 选项
- **Pandoc 2.5-2.10**：基本转换功能
- **更早版本**：可能存在兼容性问题

### 处理流程

1. **预处理阶段**：
   - 扫描LaTeX文件
   - 转换PDF图片
   - 生成临时LaTeX文件

2. **转换阶段**：
   - 检测Pandoc版本
   - 配置转换参数
   - 执行文档转换

3. **清理阶段**：
   - 删除临时文件
   - 输出结果信息

