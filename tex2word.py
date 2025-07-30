# tex_to_word.py
# 需要pandoc ,ImageMagick，poppler-utils
# 运行例子 python3 tex2word.py input.tex -o out.docx
# tbt 2025-07-18


import pypandoc
import sys
import argparse
import subprocess
import shutil
import re
from pathlib import Path

def convert_pdf_to_png(pdf_path: Path, output_dir: Path = None):
    """
    将PDF文件转换为PNG格式
    
    Args:
        pdf_path (Path): PDF文件路径
        output_dir (Path): 输出目录，如果为None则使用PDF文件同目录
    
    Returns:
        Path: 转换后的PNG文件路径
    """
    if output_dir is None:
        output_dir = pdf_path.parent
        
    png_path = output_dir / (pdf_path.stem + '.png')
    
    try:
        # 使用ImageMagick的convert命令
        subprocess.run([
            'convert', 
            '-density', '300',  # 高分辨率
            '-quality', '90',   # 高质量
            str(pdf_path),
            str(png_path)
        ], check=True, capture_output=True)
        print(f"PDF转PNG成功: {pdf_path} -> {png_path}")
        return png_path
    except subprocess.CalledProcessError:
        try:
            # 尝试使用pdftoppm (poppler-utils)
            subprocess.run([
                'pdftoppm',
                '-png',
                '-r', '300',  # 300 DPI
                str(pdf_path),
                str(png_path.with_suffix(''))
            ], check=True, capture_output=True)
            
            # pdftoppm会添加页码后缀，找到生成的文件
            generated_files = list(output_dir.glob(f"{png_path.stem}-*.png"))
            if generated_files:
                final_png = generated_files[0]
                final_png.rename(png_path)
                print(f"PDF转PNG成功: {pdf_path} -> {png_path}")
                return png_path
        except subprocess.CalledProcessError:
            print(f"警告: 无法转换PDF {pdf_path}，将保持原格式")
            return pdf_path

def preprocess_tex_file(input_file: Path, output_file: Path):
    """
    预处理LaTeX文件，将PDF图片转换为PNG格式
    
    Args:
        input_file (Path): 原始LaTeX文件
        output_file (Path): 处理后的LaTeX文件
    
    Returns:
        Path: 处理后的LaTeX文件路径
    """
    with open(input_file, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 查找所有图片引用
    # 匹配 \includegraphics{filename} 和 \includegraphics[options]{filename}
    img_pattern = r'\\includegraphics(?:\[.*?\])?\{([^}]+)\}'
    
    def replace_pdf_images(match):
        img_path_str = match.group(1)
        img_path = Path(img_path_str)
        
        # 检查是否为PDF文件
        if img_path.suffix.lower() == '.pdf':
            # 确定完整路径
            if not img_path.is_absolute():
                full_img_path = input_file.parent / img_path
            else:
                full_img_path = img_path
                
            if full_img_path.exists():
                # 转换PDF为PNG
                png_path = convert_pdf_to_png(full_img_path)
                # 返回相对于LaTeX文件的路径
                if png_path != full_img_path:  # 转换成功
                    relative_png = png_path.relative_to(input_file.parent)
                    return match.group(0).replace(img_path_str, str(relative_png))
        
        return match.group(0)  # 不是PDF或转换失败，保持原样
    
    # 替换PDF图片引用
    processed_content = re.sub(img_pattern, replace_pdf_images, content)
    
    # 写入处理后的文件
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(processed_content)
    
    print(f"LaTeX文件预处理完成: {output_file}")
    return output_file

def convert_tex_to_docx(
    input_file: Path,
    output_file: Path,
    bib_file: Path = None,
    convert_pdf_images: bool = True,
    reference_doc: Path = None,
    csl_file: Path = None
):
    """
    Converts a .tex file to a .docx file using pandoc, with optional bibliography support.

    Args:
        input_file (Path): The path to the input .tex file.
        output_file (Path): The path for the output .docx file.
        bib_file (Path, optional): The path to the .bib bibliography file. Defaults to None.
        convert_pdf_images (bool): Whether to convert PDF images to PNG. Defaults to True.
    """
    if not input_file.is_file():
        print(f"错误: 输入文件不存在 -> {input_file}")
        sys.exit(1)

    # 如果启用PDF图片转换，先预处理LaTeX文件
    if convert_pdf_images:
        temp_tex_file = input_file.parent / f"temp_{input_file.name}"
        processed_input = preprocess_tex_file(input_file, temp_tex_file)
    else:
        processed_input = input_file

    print(f"开始转换: {processed_input} -> {output_file}")

    # Pandoc 的额外参数列表
    # 检查Pandoc版本以决定使用哪些参数
    extra_args = []

    # —— 检测并启用 pandoc-crossref 过滤器（如果已安装）
    if shutil.which('pandoc-crossref'):
        extra_args.append('--filter=pandoc-crossref')
        print("已检测到 pandoc-crossref，已启用图表编号和交叉引用功能")
        # —— 使用中文“图 1”这样的前缀，加上这个 metadata
        extra_args.extend([
            '-MfigureTitle=图',
            '-MtableTitle=表',
            '-MequationTitle=式'
        ])
        print("已设置中文前缀：图、表、式")
    else:
        print("未检测到 pandoc-crossref，跳过启用该过滤器")
    
    # 检查是否支持--citeproc（较新版本的pandoc）
    try:
        result = subprocess.run(['pandoc', '--help'], capture_output=True, text=True)
        if '--citeproc' in result.stdout:
            extra_args.append('--citeproc')
            print("使用新版Pandoc的--citeproc选项")
        else:
            print("检测到旧版Pandoc，跳过--citeproc选项")
    except:
        print("无法检测Pandoc版本，使用基本选项")
    
    # 如果指定了 .bib 文件，添加到参数中
    if bib_file:
        if not bib_file.is_file():
            print(f"警告: 参考文献文件不存在 -> {bib_file}，将忽略引用处理。")
        else:
            print(f"使用参考文献文件: {bib_file}")
            extra_args.extend([f'--bibliography={bib_file}'])

    # 如果指定了 Word 模板，添加 --reference-doc
    if reference_doc:
        if not reference_doc.is_file():
            print(f"警告: 模板文件不存在 -> {reference_doc}，忽略模板设置。")
        else:
            print(f"使用 Word 模板: {reference_doc}")
            extra_args.extend([f'--reference-doc={reference_doc}'])

    # 如果指定了 CSL 文件，添加 --csl
    if csl_file:
        if not csl_file.is_file():
            print(f"警告: CSL 文件不存在 -> {csl_file}，忽略引用样式设置。")
        else:
            print(f"使用 CSL 样式文件: {csl_file}")
            extra_args.extend([f'--csl={csl_file}'])

    try:
        # 调用 pypandoc 执行转换
        pypandoc.convert_file(
            source_file=str(processed_input),
            to='docx',
            outputfile=str(output_file),
            extra_args=extra_args
        )
        print(f"转换成功！输出文件已保存至: {output_file}")
        
        # 清理临时文件
        if convert_pdf_images and processed_input != input_file:
            processed_input.unlink()
            print("已清理临时文件")
            
    except Exception as e:
        print(f"转换过程中发生错误: {e}")
        print("\n请确保您已正确安装 Pandoc(或者版本不对) 和一个完整的 LaTeX 发行版 (如 MiKTeX, TeX Live)。")
        print("对于PDF图片转换，还需要安装 ImageMagick 或 poppler-utils。")
        
        # 清理临时文件
        if convert_pdf_images and processed_input != input_file and processed_input.exists():
            processed_input.unlink()
        sys.exit(1)

def main():
    """主函数，用于解析命令行参数。"""
    parser = argparse.ArgumentParser(
        description="一个使用 Pandoc 将 .tex 文件转换为 .docx 的 Python 脚本。",
        epilog="示例: python tex_to_word.py my_paper.tex -o my_paper.docx -b references.bib"
    )
    parser.add_argument(
        "input_tex",
        type=str,
        help="需要转换的 .tex 文件的路径。"
    )
    parser.add_argument(
        "-o", "--output",
        type=str,
        help="输出的 .docx 文件的路径。如果未指定，将自动生成文件名。"
    )
    parser.add_argument(
        "-b", "--bibliography",
        type=str,
        help="可选的 .bib 参考文献文件的路径。"
    )
    parser.add_argument(
        "--no-pdf-convert",
        action="store_true",
        help="禁用PDF图片转换功能，保持原始PDF格式。"
    )
    parser.add_argument(
        "-r", "--reference-doc",
        type=str,
        help="可选的 Word 模板文件 (.docx)，用来定义页边距、标题编号等样式。"
    )
    parser.add_argument(
        "--csl",
        type=str,
        help="可选的 CSL 样式文件 (.csl)，用来控制参考文献的引用风格（如数字序号）。"
    )

    args = parser.parse_args()

    input_path = Path(args.input_tex)
    
    # 如果未指定输出路径，则自动在输入文件同目录下生成
    if args.output:
        output_path = Path(args.output)
    else:
        output_path = input_path.with_suffix('.docx')
        
    # 处理参考文献文件路径
    bib_path = Path(args.bibliography) if args.bibliography else None
    # 处理 Word 模板路径
    template_path = Path(args.reference_doc) if args.reference_doc else None
    # 处理 CSL 样式文件路径
    csl_path = Path(args.csl) if args.csl else None
    
    # 确定是否转换PDF图片
    convert_pdf_images = not args.no_pdf_convert

    convert_tex_to_docx(
        input_path,
        output_path,
        bib_file=bib_path,
        convert_pdf_images=convert_pdf_images,
        reference_doc=template_path,
        csl_file=csl_path
    )

if __name__ == "__main__":
    main()