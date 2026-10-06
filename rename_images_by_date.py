#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
按「修改日期」批量重命名图片文件
=================================

功能：
    扫描指定文件夹中的所有图片文件，读取每个文件的最后修改时间（mtime），
    并把文件重命名为  所在文件夹名_YYYY-MM-DD.原扩展名  的形式。
    例如「2024旅行」文件夹里一张 2024 年 1 月 15 日修改的图片：
        IMG_2049.jpg  ->  2024旅行_2024-01-15.jpg
    日期只精确到「日」，并且用短横线隔开写，方便肉眼辨认。

安全设计：
    1. 默认只做「预览」，会先列出全部改名计划，等你输入 y 确认后才会真正改名；
       加上 --yes 参数可跳过确认（适合脚本调用）。
    2. 加上 --dry-run 参数则只预览、绝不改名。
    3. 同一天有多张图片时，自动追加 _1、_2 序号，保证不会覆盖任何已有文件。
    4. 已经按规则命名好的文件会被跳过，脚本可以安全地重复执行；
       但若文件在改名后又被编辑过（修改日期变了），仍会按新日期重新命名。

用法示例：
    py rename_images_by_date.py                            # 不带参数：运行后提示你输入/拖入文件夹
    py rename_images_by_date.py D:\\照片\\2024旅行          # 预览 + 确认后改名
    py rename_images_by_date.py D:\\照片\\2024旅行 --dry-run  # 只看结果，不改名
    py rename_images_by_date.py D:\\照片\\2024旅行 --yes      # 直接改名，不询问
    py rename_images_by_date.py D:\\照片 --recursive          # 连同所有子文件夹一起处理

在 PyCharm 里使用：
    直接点绿色三角运行即可（不需要配置任何参数），程序会在 Run 窗口里
    提示你输入文件夹路径，把文件夹拖进窗口再回车就行。
"""

import argparse          # 解析命令行参数
import os                # 文件路径、重命名等操作系统接口
import re                # 用正则判断文件名是否已符合目标格式
import sys               # 用于退出程序、处理控制台编码
from datetime import datetime   # 把修改时间转换成 YYYY-MM-DD 日期字符串
from pathlib import Path        # 面向对象的路径操作，比 os.path 更清晰


# ---------------------------------------------------------------------------
# 配置项
# ---------------------------------------------------------------------------

# 需要处理的图片扩展名（统一小写，比较时会先把实际扩展名转成小写）。
# 如需支持其他格式，直接往这个集合里加即可。
IMAGE_EXTENSIONS = {
    ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp",
    ".tif", ".tiff", ".heic", ".heif", ".avif", ".svg", ".ico",
    ".raw", ".cr2", ".nef", ".arw", ".dng", ".orf", ".rw2",
}

# 日期格式：只保留年-月-日，用短横线隔开，方便肉眼辨认（不精确到时分秒）
DATE_FORMAT = "%Y-%m-%d"

# 「所在文件夹名」和「日期」之间的连接符
NAME_JOINER = "_"


# ---------------------------------------------------------------------------
# 工具函数
# ---------------------------------------------------------------------------

def is_image_file(path: Path) -> bool:
    """
    判断给定路径是不是「我们要处理的图片文件」。

    条件有两个：
        1. 它是一个普通文件（排除文件夹、符号链接指向的目录等）；
        2. 它的扩展名在 IMAGE_EXTENSIONS 白名单里（忽略大小写）。
    """
    if not path.is_file():
        return False
    # suffix 返回带点的扩展名，例如 ".JPG"；统一转小写后再比较
    return path.suffix.lower() in IMAGE_EXTENSIONS


def collect_image_files(folder: Path, recursive: bool) -> list[Path]:
    """
    收集文件夹下所有的图片文件。

    参数：
        folder    : 要扫描的文件夹
        recursive : True 表示连同所有层级的子文件夹一起扫描

    返回：
        按「修改时间」从早到晚排序好的图片路径列表。
        先排序再改名，可以让生成的 _1、_2 冲突序号顺序稳定、可预期。
    """
    # rglob("*") 递归遍历，glob("*") 只遍历当前一层
    walker = folder.rglob("*") if recursive else folder.glob("*")

    images = [p for p in walker if is_image_file(p)]

    # 按修改时间排序；同一时刻的文件再按文件名排序，保证结果完全确定
    images.sort(key=lambda p: (p.stat().st_mtime, p.name.lower()))
    return images


def build_target_name(mtime: float, extension: str, folder_name: str,
                      prefix: str = "") -> str:
    """
    生成新的文件名（不含所在目录）。

    命名规则：  [自定义前缀] + 所在文件夹名 + _ + 年-月-日 + 原扩展名
    例如文件夹叫「2024旅行」，图片修改于 2024 年 1 月 15 日，则生成：
        2024旅行_2024-01-15.jpg

    参数：
        mtime       : os.stat().st_mtime 得到的时间戳（浮点秒数）
        extension   : 原扩展名，例如 ".jpg"（保持原样大小写，不强制转换）
        folder_name : 图片所在文件夹的名字（即父目录名），放在文件名开头
        prefix      : 可选的自定义前缀，会拼在最前面，用于区分不同批次

    返回：
        形如 "2024旅行_2024-01-15.jpg" 的字符串
    """
    # fromtimestamp 把 Unix 时间戳转成本地时间的 datetime 对象，
    # 再按 DATE_FORMAT 只取年月日部分
    date_text = datetime.fromtimestamp(mtime).strftime(DATE_FORMAT)

    head = f"{prefix}{folder_name}"
    # 文件夹名有可能为空（例如图片直接放在 D:\ 这样的盘符根目录下），
    # 这种情况下不要留下一个孤零零的下划线开头
    joiner = NAME_JOINER if head else ""
    return f"{head}{joiner}{date_text}{extension}"


def is_already_renamed(path: Path, prefix: str) -> bool:
    """
    判断文件是否「已经按本脚本的规则命名好了」，是则应当跳过。

    这一步是脚本幂等性的关键。光比较「目标名 == 原名」是不够的：
    现在只精确到「日」，同一天里的几十张照片算出来的基础名完全相同，
    于是会生成 xxx_1.jpg、xxx_2.jpg 这样的序号名。第二次运行时，这些文件
    算出来的基础目标名仍然是不带序号的 xxx.jpg，该名已被占用，就会又被
    改成 xxx_3.jpg、xxx_4.jpg……序号无限增长。

    因此这里用正则匹配「文件夹名 + 日期 + 可选序号 + 原扩展名」的整体格式。
    由于基础名是由文件自身的修改时间反推出来的，也就顺带校验了名字里的日期
    确实等于该文件自己的修改日期——如果文件在改名后又被编辑过（跨天了），
    仍会按新日期重新命名。
    """
    stat = path.stat()

    # 用当前规则重新算一次「它本来应该叫什么」
    base_name = build_target_name(stat.st_mtime, path.suffix, path.parent.name, prefix)

    # 把基础名拆成「不含扩展名的主干」和「扩展名」两部分
    suffix = path.suffix
    stem = base_name[: -len(suffix)] if suffix else base_name

    pattern = (
        r"^"
        + re.escape(stem)      # 文件夹名_日期，必须与文件自身的修改日期一致
        + r"(_\d+)?"           # 可选的冲突序号，如 _1、_2
        + re.escape(suffix)    # 原扩展名，保留大小写
        + r"$"
    )
    return re.match(pattern, path.name) is not None


def resolve_conflict(target: Path, taken: set[str]) -> Path:
    """
    解决文件名冲突：如果目标文件名已经被占用，就在末尾追加 _1、_2、_3 ...
    直到找到一个从未使用过的名字为止。

    参数：
        target : 期望的目标路径（可能已存在）
        taken  : 本次运行中「已经计划要用」的小写文件名集合。
                 光靠 target.exists() 是不够的：日期只精确到「日」，
                 同一天拍的几十张图片会算出完全相同的目标名，它们之间
                 互相撞名，而此时改名还没真正发生，磁盘上根本查不到。

    返回：
        一个保证不与磁盘上任何文件、也不与本批次其他文件冲突的路径。
    """
    # Windows / macOS 文件系统默认不区分大小写，所以统一用小写做比较
    parent = target.parent
    key = target.name.lower()

    if not target.exists() and key not in taken:
        return target                      # 名字没被占用，直接用

    stem, suffix = target.stem, target.suffix   # 拆成「主名」和「扩展名」
    counter = 1
    while True:
        candidate = parent / f"{stem}_{counter}{suffix}"
        if not candidate.exists() and candidate.name.lower() not in taken:
            return candidate
        counter += 1


def plan_renames(folder: Path, recursive: bool, prefix: str) -> list[tuple[Path, Path]]:
    """
    生成完整的「改名计划」，但不会真正执行任何改名操作。

    返回：
        [(原路径, 新路径), ...] 的列表，只包含确实需要改名的文件。
        已经符合目标命名的文件不会出现在列表里（幂等性保证）。
    """
    images = collect_image_files(folder, recursive)

    # 本批次「已经占用」的新文件名集合（统一小写，适配 Windows/macOS
    # 这类不区分大小写的文件系统）。
    taken: set[str] = set()

    # --- 第一轮：先把已经命名规范的文件登记为「已占用」 -------------------
    # 必须分两轮处理，否则后改名的文件可能算出一个已被前面文件占用、
    # 但尚未登记的名字，导致互相覆盖。
    pending: list[Path] = []
    for path in images:
        if is_already_renamed(path, prefix):
            taken.add(path.name.lower())
        else:
            pending.append(path)

    # --- 第二轮：为剩下的文件计算新名字 -----------------------------------
    plan: list[tuple[Path, Path]] = []
    for src in pending:
        # src.parent.name 就是「图片所在文件夹的名字」，递归模式下每个子文件夹
        # 会各自取自己的名字，因此不同子目录的图片前缀不会混在一起
        new_name = build_target_name(
            src.stat().st_mtime, src.suffix, src.parent.name, prefix
        )
        dst = resolve_conflict(src.with_name(new_name), taken)
        taken.add(dst.name.lower())
        plan.append((src, dst))

    return plan


def execute_plan(plan: list[tuple[Path, Path]]) -> tuple[int, int]:
    """
    真正执行改名操作。

    返回：
        (成功数量, 失败数量)
    """
    ok, failed = 0, 0
    for src, dst in plan:
        try:
            # os.rename 在同一磁盘分区内是原子的、瞬间完成的
            os.rename(src, dst)
            print(f"  [OK]   {src.name}  ->  {dst.name}")
            ok += 1
        except OSError as exc:
            # 常见失败原因：文件被其他程序占用、没有写权限、跨设备等
            print(f"  [失败] {src.name}  ->  {dst.name}  原因: {exc}")
            failed += 1
    return ok, failed


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------

def main() -> int:
    # Windows 的控制台默认编码可能是 GBK，直接 print 中文/特殊字符会报错，
    # 这里强制把标准输出切到 UTF-8，让脚本在中文 Windows 上也能正常显示。
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8")
            except (ValueError, OSError):
                pass   # 输出被重定向到不支持的地方时，忽略即可

    parser = argparse.ArgumentParser(
        description="把文件夹里的图片重命名为「所在文件夹名_年-月-日」格式，"
                    "例如 2024旅行_2024-01-15.jpg（按文件的最后修改日期）。",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="示例：\n"
               "  py rename_images_by_date.py                     # 不带参数，运行后按提示拖入文件夹\n"
               "  py rename_images_by_date.py D:\\照片\\2024旅行 --dry-run\n"
               "  py rename_images_by_date.py D:\\照片\\2024旅行 --prefix 精选\n"
               "  py rename_images_by_date.py D:\\照片 --recursive --yes\n",
    )
    parser.add_argument("folder", nargs="?", default=None,
                        help="要处理的图片文件夹路径。不传这个参数时，运行后会提示你输入")
    parser.add_argument("-r", "--recursive", action="store_true",
                        help="同时处理所有子文件夹中的图片（每个子文件夹各自取自己的名字作前缀）")
    parser.add_argument("-p", "--prefix", default="",
                        help="额外加在最前面的自定义前缀，例如 --prefix 精选 生成 精选2024旅行_2024-01-15.jpg")
    parser.add_argument("-n", "--dry-run", action="store_true",
                        help="只预览改名计划，不实际修改任何文件")
    parser.add_argument("-y", "--yes", action="store_true",
                        help="跳过确认提示，直接执行改名")
    args = parser.parse_args()

    # --- 第一步：确定要处理哪个文件夹 ---------------------------------------
    # 没有传 folder 参数时（例如在 PyCharm 里直接点绿色三角运行），
    # 不再报错退出，而是改成交互式提问，让你把文件夹拖进窗口即可。
    raw_folder = args.folder
    if not raw_folder:
        print("未指定文件夹。请把目标文件夹用鼠标拖进下面的窗口，")
        print("或者手动粘贴完整路径，然后按回车：")
        print(r"（路径长这样：D:\照片\2024旅行）")
        try:
            raw_folder = input("> ")
        except EOFError:
            # 输入流被关闭（例如被重定向且没有内容），直接结束
            print("没有读到输入，程序退出。")
            return 1

    # 把文件夹拖进终端时，系统会自动在路径两边加上一对引号，这里统一去掉；
    # 再清理首尾空格，避免复制粘贴时带进来多余字符。
    raw_folder = raw_folder.strip().strip('"').strip("'").strip()
    if not raw_folder:
        print("输入为空，程序退出。")
        return 1

    folder = Path(raw_folder).expanduser().resolve()
    if not folder.exists():
        print(f"错误：路径不存在 -> {folder}")
        return 1
    if not folder.is_dir():
        print(f"错误：这不是一个文件夹 -> {folder}")
        return 1

    print(f"目标文件夹：{folder}")
    print(f"递归子目录：{'是' if args.recursive else '否'}")
    print(f"自定义前缀：{args.prefix or '(无)'}")
    print(f"命名规则  ：所在文件夹名{NAME_JOINER}年-月-日.原扩展名")
    print(f"            例如本文件夹下的图片会变成 "
          f"{args.prefix}{folder.name}{NAME_JOINER}2024-01-15.jpg")
    if args.recursive:
        print("            （递归模式下，每个子文件夹各自取自己的名字）")
    print("-" * 60)

    # --- 第二步：扫描并生成改名计划 -----------------------------------------
    try:
        plan = plan_renames(folder, args.recursive, args.prefix)
    except PermissionError as exc:
        print(f"错误：没有权限读取该文件夹 -> {exc}")
        return 1

    if not plan:
        print("没有找到需要重命名的图片文件（可能目录为空，或文件已经是目标命名）。")
        return 0

    print(f"共找到 {len(plan)} 个文件需要重命名：\n")
    # 只展示相对路径，输出更清爽
    for src, dst in plan:
        print(f"  {src.relative_to(folder)}\n      ->  {dst.name}")
    print()

    # --- 第三步：预览模式直接结束 -------------------------------------------
    if args.dry_run:
        print("（--dry-run 预览模式，未对任何文件做出修改）")
        return 0

    # --- 第四步：确认后执行 -------------------------------------------------
    if not args.yes:
        answer = input(f"确认要重命名以上 {len(plan)} 个文件吗？此操作不可自动撤销 (y/N): ")
        if answer.strip().lower() not in ("y", "yes"):
            print("已取消，未修改任何文件。")
            return 0

    print("-" * 60)
    ok, failed = execute_plan(plan)
    print("-" * 60)
    print(f"完成：成功 {ok} 个，失败 {failed} 个。")

    # 有任何失败时返回非 0 退出码，方便被其他脚本/CI 检测到
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
